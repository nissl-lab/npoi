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

namespace NPOI.XWPF.UserModel
{
    using System;

    using NPOI.OpenXml4Net.OPC;
    using NPOI.Util;
    using System.Reflection;

    /**
     * @author Yegor Kozlov
     */
    public class XWPFFactory : POIXMLFactory
    {

        private static POILogger logger = POILogFactory.GetLogger(typeof(XWPFFactory));

        private XWPFFactory()
        {

        }

        private static XWPFFactory inst = new XWPFFactory();

        public static XWPFFactory GetInstance()
        {
            return inst;
        }


        public override POIXMLDocumentPart CreateDocumentPart(POIXMLDocumentPart parent, PackageRelationship rel, PackagePart part)
        {
            POIXMLRelation descriptor = XWPFRelation.GetInstance(rel.RelationshipType);
            if (descriptor == null || descriptor.RelationClass == null)
            {
                logger.Log(POILogger.DEBUG, "using default POIXMLDocumentPart for " + rel.RelationshipType);
                return new POIXMLDocumentPart(part, rel);
            }

            try
            {
                Type cls = descriptor.RelationClass;
                try
                {
                    ConstructorInfo constructor = cls.GetConstructor(new Type[] { typeof(POIXMLDocumentPart), typeof(PackagePart), typeof(PackageRelationship) });
                    return constructor.Invoke(new object[] { parent, part, rel }) as POIXMLDocumentPart;
                }
                catch (Exception)
                {
                    ConstructorInfo constructor = cls.GetConstructor(new Type[] { typeof(PackagePart), typeof(PackageRelationship) });
                    return constructor.Invoke(new object[] { part, rel }) as POIXMLDocumentPart;
                }
            }
            catch (Exception e)
            {
                throw new POIXMLException(e);
            }
        }


        public override POIXMLDocumentPart CreateDocumentPart(POIXMLRelation descriptor)
        {
            try
            {
                Type cls = descriptor.RelationClass;
                ConstructorInfo constructor = cls.GetConstructor(BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public, 
                    null, Type.EmptyTypes, null);
                return constructor.Invoke(new object[] { }) as POIXMLDocumentPart;
            }
            catch (Exception e)
            {
                throw new POIXMLException(e);
            }
        }

    }
}
