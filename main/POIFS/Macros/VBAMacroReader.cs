/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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

namespace NPOI.POIFS.Macros
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Text;
    using ICSharpCode.SharpZipLib.Zip;
    using NPOI.POIFS.FileSystem;
    using NPOI.Util;

    /**
     * Finds all VBA Macros in an office file (OLE2/POIFS and OOXML/OPC),
     *  and returns them.
     * 
     * @since 3.15-beta2
     */
    public class VBAMacroReader : ICloseable
    {
        protected static String VBA_PROJECT_OOXML = "vbaProject.bin";
        protected static String VBA_PROJECT_POIFS = "VBA";

        private NPOIFSFileSystem fs;

        public VBAMacroReader(InputStream rstream)
        {
            PushbackInputStream stream = new PushbackInputStream(rstream, 8);
            byte[] header8 = IOUtils.PeekFirst8Bytes(stream);

            if (NPOIFSFileSystem.HasPOIFSHeader(header8))
            {
                fs = new NPOIFSFileSystem(stream);
            }
            else
            {
                OpenOOXML(stream);
            }
        }

        public VBAMacroReader(FileInfo file)
        {
            try
            {
                this.fs = new NPOIFSFileSystem(file);
            }
            catch (OfficeXmlFileException)
            {
                OpenOOXML(file.OpenRead());
            }
        }
        public VBAMacroReader(NPOIFSFileSystem fs)
        {
            this.fs = fs;
        }

        private void OpenOOXML(Stream zipFile)
        {
            ZipInputStream zis = new ZipInputStream(zipFile);
            ZipEntry zipEntry;
            while ((zipEntry = zis.GetNextEntry()) != null)
            {
                if (zipEntry.Name.EndsWith(VBA_PROJECT_OOXML, StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        // Make a NPOIFS from the contents, and close the stream
                        this.fs = new NPOIFSFileSystem(zis);
                        return;
                    }
                    catch (IOException e)
                    {
                        // Tidy up
                        zis.Close();

                        // Pass on
                        throw e;
                    }
                }
            }
            zis.Close();
            throw new ArgumentException("No VBA project found");
        }

        public void Close()
        {
            fs.Close();
            fs = null;
        }

        /**
         * Reads all macros from all modules of the opened office file. 
         * @return All the macros and their contents
         *
         * @since 3.15-beta2
         */
        public Dictionary<String, String> ReadMacros()
        {
            ModuleMap modules = new ModuleMap();
            FindMacros(fs.Root, modules);

            Dictionary<String, String> moduleSources = new Dictionary<String, String>();
            foreach (KeyValuePair<String, Module> entry in modules)
            {
                Module module = entry.Value;
                if (module.buf != null && module.buf.Length > 0)
                { // Skip empty modules
                    moduleSources.Add(entry.Key, ModuleMap.charset.GetString(module.buf));
                }
            }
            return moduleSources;
        }

        protected class Module
        {
            public int? offset;
            public byte[] buf;
            public void Read(Stream in1)
            {
                MemoryStream out1 = new MemoryStream();
                IOUtils.Copy(in1, out1);
                out1.Close();
                buf = out1.ToArray();
            }
        }
        protected class ModuleMap : Dictionary<String, Module>
        {
            static ModuleMap()
            {
#if NETSTANDARD2_1 || NETSTANDARD2_0
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
#endif
                charset =Encoding.GetEncoding(1252);
            }
            public static Encoding charset = null;
            //Charset charset = Charset.ForName("Cp1252"); // default charset
            public Module Get(string key)
            {
                return ContainsKey(key) ? this[key] : null;
            }
            public Module Put(string key, Module value)
            {
                Module oldValue = null;
                if (ContainsKey(key))
                {
                    oldValue = this[key];
                    this[key] = value;
                }
                else
                {
                    Add(key, value);
                }
                return oldValue;
            }
        }

        /**
         * Recursively traverses directory structure rooted at <tt>dir</tt>.
         * For each macro module that is found, the module's name and code are
         * Added to <tt>modules</tt>.
         *
         * @param dir
         * @param modules
         * @throws IOException
         * @since 3.15-beta2
         */
        protected void FindMacros(DirectoryNode dir, ModuleMap modules)
        {
            if (VBA_PROJECT_POIFS.Equals(dir.Name, StringComparison.OrdinalIgnoreCase))
            {
                // VBA project directory, process
                ReadMacros(dir, modules);
            }
            else
            {
                // Check children
                foreach (Entry child in dir)
                {
                    if (child is DirectoryNode)
                    {
                        FindMacros((DirectoryNode)child, modules);
                    }
                }
            }
        }

        /**
         * Read <tt>length</tt> bytes of MBCS (multi-byte character Set) characters from the stream
         *
         * @param stream the inputstream to read from
         * @param length number of bytes to read from stream
         * @param charset the character Set encoding of the bytes in the stream
         * @return a java String in the supplied character Set
         * @throws IOException
         */
        private static String ReadString(InputStream stream, int length, Encoding charset)
        {
            byte[] buffer = new byte[length];
            int count = stream.Read(buffer);
            //return new String(buffer, 0, count, charset);
            return charset.GetString(buffer, 0, count);
        }

        /**
         * Reads module from DIR node in input stream and Adds it to the modules map for decompression later
         * on the second pass through this function, the module will be decompressed
         * 
         * Side-effects: Adds a new module to the module map or Sets the buf field on the module
         * to the decompressed stream contents (the VBA code for one module)
         *
         * @param in the Run-length encoded input stream to read from
         * @param streamName the stream name of the module
         * @param modules a map to store the modules
         * @throws IOException
         */
        private static void ReadModule(RLEDecompressingInputStream in1, String streamName, ModuleMap modules)
        {
            int moduleOffset = in1.ReadInt();
            Module module = modules.Get(streamName);
            if (module == null)
            {
                // First time we've seen the module. Add it to the ModuleMap and decompress it later
                module = new Module();
                module.offset = moduleOffset;
                modules.Put(streamName, module);
                // Would Adding module.Read(in1) here be correct?
            }
            else
            {
                // Decompress a previously found module and store the decompressed result into module.buf
                InputStream stream = new RLEDecompressingInputStream(
                        new MemoryStream(module.buf, moduleOffset, module.buf.Length - moduleOffset)
                );
                module.Read(stream);
                stream.Close();
            }
        }

        private static void ReadModule(DocumentInputStream dis, String name, ModuleMap modules)
        {
            Module module = modules.Get(name);
            // TODO Refactor this to fetch dir then do the rest
            if (module == null)
            {
                // no DIR stream with offsets yet, so store the compressed bytes for later
                module = new Module();
                modules.Put(name, module);
                module.Read(dis);
            }
            else
            {
                if (module.offset == null)
                {
                    //This should not happen. bug 59858
                    throw new IOException("Module offset for '" + name + "' was never Read.");
                }
                // we know the offset already, so decompress immediately on-the-fly
                long skippedBytes = dis.Skip(module.offset.Value);
                if (skippedBytes != module.offset)
                {
                    throw new IOException("tried to skip " + module.offset + " bytes, but actually skipped " + skippedBytes + " bytes");
                }
                InputStream stream = new RLEDecompressingInputStream(dis);
                module.Read(stream);
                stream.Close();
            }

        }

        /**
          * Skips <tt>n</tt> bytes in an input stream, throwing IOException if the
          * number of bytes skipped is different than requested.
          * @throws IOException
          */
        private static void TrySkip(InputStream in1, long n)
        {
            long skippedBytes = in1.Skip(n);
            if (skippedBytes != n)
            {
                if (skippedBytes < 0)
                {
                    throw new IOException(
                        "Tried skipping " + n + " bytes, but no bytes were skipped. "
                        + "The end of the stream has been reached or the stream is closed.");
                }
                else
                {
                    throw new IOException(
                        "Tried skipping " + n + " bytes, but only " + skippedBytes + " bytes were skipped. "
                        + "This should never happen.");
                }
            }
        }

        // Constants from MS-OVBA: https://msdn.microsoft.com/en-us/library/office/cc313094(v=office.12).aspx
        private const int EOF = -1;
        private const int VERSION_INDEPENDENT_TERMINATOR = 0x0010;
        private const int VERSION_DEPENDENT_TERMINATOR = 0x002B;
        private const int PROJECTVERSION = 0x0009;
        private const int PROJECTCODEPAGE = 0x0003;
        private const int STREAMNAME = 0x001A;
        private const int MODULEOFFSET = 0x0031;
        private const int MODULETYPE_PROCEDURAL = 0x0021;
        private const int MODULETYPE_DOCUMENT_CLASS_OR_DESIGNER = 0x0022;
        private const int PROJECTLCID = 0x0002;

        /**
         * Reads VBA Project modules from a VBA Project directory located at
         * <tt>macroDir</tt> into <tt>modules</tt>.
         *
         * @since 3.15-beta2
         */
        protected void ReadMacros(DirectoryNode macroDir, ModuleMap modules)
        {
            foreach (Entry entry in macroDir)
            {
                if (!(entry is DocumentNode)) { continue; }

                String name = entry.Name;
                DocumentNode document = (DocumentNode)entry;
                DocumentInputStream dis = new DocumentInputStream(document);
                try
                {
                    if ("dir".Equals(name, StringComparison.OrdinalIgnoreCase))
                    {
                        // process DIR
                        RLEDecompressingInputStream in1 = new RLEDecompressingInputStream(dis);
                        String streamName = null;
                        int recordId = 0;
                        try
                        {
                            while (true)
                            {
                                recordId = in1.ReadShort();
                                if (EOF == recordId
                                        || VERSION_INDEPENDENT_TERMINATOR == recordId)
                                {
                                    break;
                                }
                                int recordLength = in1.ReadInt();
                                switch (recordId)
                                {
                                    case PROJECTVERSION:
                                        TrySkip(in1, 6);
                                        break;
                                    case PROJECTCODEPAGE:
                                        int codepage = in1.ReadShort();
                                        ModuleMap.charset = Encoding.GetEncoding(codepage); //Charset.ForName("Cp" + codepage);
                                        break;
                                    case STREAMNAME:
                                        streamName = ReadString(in1, recordLength, ModuleMap.charset);
                                        break;
                                    case MODULEOFFSET:
                                        ReadModule(in1, streamName, modules);
                                        break;
                                    default:
                                        TrySkip(in1, recordLength);
                                        break;
                                }
                            }
                        }
                        catch (IOException e)
                        {
                            throw new IOException(
                                    "Error occurred while Reading macros at section id "
                                    + recordId + " (" + HexDump.ShortToHex(recordId) + ")", e);
                        }
                        finally
                        {
                            in1.Close();
                        }
                    }
                    else if (!name.StartsWith("__SRP", StringComparison.OrdinalIgnoreCase)
                          && !name.StartsWith("_VBA_PROJECT", StringComparison.OrdinalIgnoreCase))
                    {
                        // process module, skip __SRP and _VBA_PROJECT since these do not contain macros
                        ReadModule(dis, name, modules);
                    }
                }
                finally
                {
                    dis.Close();
                }
            }
        }
    }

}