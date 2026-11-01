/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using NPOI;
using NPOI.OpenXmlFormats.Spreadsheet;
using NUnit.Framework;

namespace TestCases.Util
{
    [TestFixture]
    public class TestXmlWriterEscaping
    {
        private const string Probe = "a&b<c>d\"e'f]]>g";
        private const BindingFlags AnyInstance = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;

        private static readonly HashSet<string> RawXmlProperties = new HashSet<string> { "Any", "InnerXml", "ItemXml", "XmlText" };

        private static readonly HashSet<string> PropertiesNotReadBack = new HashSet<string>
        {
            // Parse fills it from the rgb attribute, and Write writes only the rgb bytes
            "NPOI.OpenXmlFormats.Spreadsheet.CT_RgbColor.rgbHex",
            // Write derives it from Value, as "preserve" for a leading or trailing space or tab
            "NPOI.OpenXmlFormats.Wordprocessing.CT_Text.space",
            // Parse reads o:spt into id
            "NPOI.OpenXmlFormats.Vml.CT_Shapetype.id",
            // Parse looks up "w.fullDate", with a dot, so the w:fullDate attribute that Write emits is never found
            "NPOI.OpenXmlFormats.Wordprocessing.CT_SdtDate.fullDate",
            // Write and Parse skip the attributes inherited from CT_TrackChange
            "NPOI.OpenXmlFormats.Wordprocessing.CT_RunTrackChange.author",
            "NPOI.OpenXmlFormats.Wordprocessing.CT_RunTrackChange.date",
            "NPOI.OpenXmlFormats.Wordprocessing.CT_RunTrackChange.id",
            // Write and Parse skip fldData
            "NPOI.OpenXmlFormats.Wordprocessing.CT_SimpleField.fldData.space",
            "NPOI.OpenXmlFormats.Wordprocessing.CT_SimpleField.fldData.Value"
        };

        /// <summary>
        /// Matches an element or attribute name with a prefix, e.g. "&lt;c:f" or " r:id=", other than the predefined xmlns and xml prefixes.
        /// </summary>
        private static readonly Regex PrefixedName = new Regex(@"[<\s/](?!xmlns:|xml:)([A-Za-z_][\w.-]*):[A-Za-z_]");

        public static IEnumerable<TestCaseData> SerializableTypes()
        {
            // Discovered rather than listed, so a writer added later is checked without being registered here.
            return typeof(CT_Worksheet).Assembly.GetTypes()
                .Where(t => t.IsClass && !t.IsAbstract && !t.ContainsGenericParameters)
                .Where(t => GetWrite(t) != null && GetParse(t) != null && t.GetConstructor(Type.EmptyTypes) != null)
                .Where(t => GetStringProperties(t).Any())
                .OrderBy(t => t.FullName, StringComparer.Ordinal)
                .Select(t => new TestCaseData(t).SetName(t.FullName));
        }

        [TestCaseSource(nameof(SerializableTypes))]
        public void Write_KeepsXmlWellFormed_StringPropertiesWithMarkupCharacters(Type type)
        {
            var instance = CreateInstance(type);
            var probed = SetProbes(instance, null)
                .Concat(GetChildren(type).Where(c => c.GetValue(instance) != null).SelectMany(c => SetProbes(instance, c)))
                .ToList();

            string xml;
            try
            {
                xml = Write(instance);
            }
            catch (TargetInvocationException e) when (e.InnerException?.GetType() == GetWriteFailure(CreateInstance(type)))
            {
                Assert.Ignore($"Write cannot serialize a default instance: {e.InnerException?.GetType().Name}");
                return;
            }

            XmlDocument document;
            try
            {
                document = Load(xml);
            }
            catch (XmlException e)
            {
                Assert.Fail($"{e.Message}{Environment.NewLine}{xml}");
                return;
            }

            var written = document.DocumentElement.FirstChild;
            if (written == null)
            {
                Assert.That(probed.Where(p => p.Child == null), Is.Empty, "Write wrote nothing");
                return;
            }

            var parsed = GetParse(type).Invoke(null, new object[] { written, POIXMLDocumentPart.NamespaceManager });
            var readBack = probed.ToLookup(p => ReadBack(parsed, p) == ProbeFor(p));
            var notReadBack = readBack[false]
                .Where(p => !PropertiesNotReadBack.Contains($"{type.FullName}.{PathOf(p)}"))
                .Select(p => $"{PathOf(p)} = {ReadBack(parsed, p) ?? "null"}")
                .ToList();
            var listedButReadBack = readBack[true]
                .Where(p => PropertiesNotReadBack.Contains($"{type.FullName}.{PathOf(p)}"))
                .Select(PathOf)
                .ToList();
            Assert.Multiple(() =>
            {
                Assert.That(notReadBack, Is.Empty, xml);
                Assert.That(listedButReadBack, Is.Empty, "listed in PropertiesNotReadBack but read back");
            });
        }

        private static MethodInfo GetWrite(Type type)
        {
            return type.GetMethod("Write", AnyInstance, null, new[] { typeof(StreamWriter), typeof(string) }, null);
        }

        private static MethodInfo GetParse(Type type)
        {
            var parse = type.GetMethod("Parse", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic, null, new[] { typeof(XmlNode), typeof(XmlNamespaceManager) }, null);
            return parse != null && parse.ReturnType == type ? parse : null;
        }

        private static IEnumerable<PropertyInfo> GetSettableProperties(Type type)
        {
            return type.GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .Where(p => p.CanRead && p.GetSetMethod() != null && p.GetIndexParameters().Length == 0);
        }

        private static IEnumerable<PropertyInfo> GetStringProperties(Type type)
        {
            return GetSettableProperties(type)
                .Where(p => !RawXmlProperties.Contains(p.Name))
                .Where(p => p.PropertyType == typeof(string) || p.PropertyType == typeof(List<string>) || p.PropertyType == typeof(string[]));
        }

        private static IEnumerable<PropertyInfo> GetChildren(Type type)
        {
            return GetSettableProperties(type)
                .Where(p => p.PropertyType.IsClass && !p.PropertyType.IsAbstract && p.PropertyType.Assembly == typeof(CT_Worksheet).Assembly)
                .Where(p => p.PropertyType.GetConstructor(Type.EmptyTypes) != null);
        }

        private static object CreateInstance(Type type)
        {
            var instance = Activator.CreateInstance(type);
            InitializeNullLists(instance);
            foreach (var child in GetChildren(type).Where(c => c.GetValue(instance) == null))
            {
                var value = Activator.CreateInstance(child.PropertyType);
                InitializeNullLists(value);
                if (GetWrite(child.PropertyType) == null || GetWriteFailure(value) == null)
                {
                    child.SetValue(instance, value);
                }
            }
            return instance;
        }

        private static void InitializeNullLists(object instance)
        {
            var lists = GetSettableProperties(instance.GetType())
                .Where(p => p.PropertyType.IsGenericType && p.PropertyType.GetGenericTypeDefinition() == typeof(List<>));
            foreach (var list in lists)
            {
                if (list.GetValue(instance) == null)
                {
                    list.SetValue(instance, Activator.CreateInstance(list.PropertyType));
                }
            }
        }

        private static Type GetWriteFailure(object instance)
        {
            try
            {
                Write(instance);
                return null;
            }
            catch (TargetInvocationException e)
            {
                return e.InnerException?.GetType();
            }
        }

        private static List<(PropertyInfo Child, PropertyInfo Property)> SetProbes(object instance, PropertyInfo child)
        {
            var target = child == null ? instance : child.GetValue(instance);
            return GetStringProperties(target.GetType())
                .Select(p => (Child: child, Property: p))
                .Where(p => TrySetProbe(target, p))
                .ToList();
        }

        private static bool TrySetProbe(object target, (PropertyInfo Child, PropertyInfo Property) probe)
        {
            var value = ProbeFor(probe);
            object typed = probe.Property.PropertyType == typeof(string) ? value
                : probe.Property.PropertyType == typeof(string[]) ? new[] { value }
                : new List<string> { value };
            try
            {
                probe.Property.SetValue(target, typed);
                return true;
            }
            catch (TargetInvocationException)
            {
                return false;
            }
        }

        private static string PathOf((PropertyInfo Child, PropertyInfo Property) probe)
        {
            return probe.Child == null ? probe.Property.Name : $"{probe.Child.Name}.{probe.Property.Name}";
        }

        private static string ProbeFor((PropertyInfo Child, PropertyInfo Property) probe)
        {
            return Probe + PathOf(probe);
        }

        private static string ReadBack(object parsed, (PropertyInfo Child, PropertyInfo Property) probe)
        {
            var target = probe.Child == null ? parsed : probe.Child.GetValue(parsed);
            var value = target == null ? null : probe.Property.GetValue(target);
            return value as string ?? (value as IEnumerable<string>)?.FirstOrDefault();
        }

        private static string Write(object instance)
        {
            using var stream = new MemoryStream();
            using (var writer = new StreamWriter(stream, new UTF8Encoding(false), 1024, true))
            {
                GetWrite(instance.GetType()).Invoke(instance, new object[] { writer, "node" });
            }
            return Encoding.UTF8.GetString(stream.ToArray());
        }

        private static XmlDocument Load(string xml)
        {
            var namespaces = POIXMLDocumentPart.NamespaceManager;
            var declarations = PrefixedName.Matches(xml).Cast<Match>()
                .Select(m => m.Groups[1].Value)
                .Distinct()
                .Select(p => $" xmlns:{p}=\"{namespaces.LookupNamespace(p) ?? "urn:" + p}\"");
            var document = new XmlDocument();
            document.LoadXml($"<root xmlns=\"{namespaces.DefaultNamespace}\"{string.Concat(declarations)}>{xml}</root>");
            return document;
        }
    }
}
