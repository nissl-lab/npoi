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

using NPOI.Util;
using NPOI.XSSF.UserModel;
using System;
using NPOI.OpenXml4Net.OPC;
using System.Reflection;
namespace NPOI.XSSF.UserModel
{

    /**
     * Instantiates sub-classes of POIXMLDocumentPart depending on their relationship type
     *
     * @author Yegor Kozlov
     */
    public class XSSFFactory : POIXMLFactory
    {
        private static POILogger logger = POILogFactory.GetLogger(typeof(XSSFFactory));

        private XSSFFactory()
        {

        }

        private static XSSFFactory inst = new XSSFFactory();

        public static XSSFFactory GetInstance()
        {
            return inst;
        }


        public override POIXMLDocumentPart CreateDocumentPart(POIXMLDocumentPart parent, PackageRelationship rel, PackagePart part)
        {
            POIXMLRelation descriptor = XSSFRelation.GetInstance(rel.RelationshipType);
            if (descriptor == null || descriptor.RelationClass == null)
            {
                logger.Log(POILogger.DEBUG, "using default POIXMLDocumentPart for " + rel.RelationshipType);
                return new POIXMLDocumentPart(part, rel);
            }

            try
            {
                Type cls = descriptor.RelationClass;
                ConstructorInfo constructor = cls.GetConstructor(BindingFlags.NonPublic | BindingFlags.Instance,null, new Type[] { typeof(PackagePart), typeof(PackageRelationship) },null);
                return (POIXMLDocumentPart)constructor.Invoke(new object[] { part, rel });
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
                //Console.WriteLine(cls.ToString());
                ConstructorInfo constructor = cls.GetConstructor(new Type[] { });
                return (POIXMLDocumentPart)constructor.Invoke(new object[] { });
            }
            catch (Exception e)
            {
                throw new POIXMLException(e);
            }
        }

    }


}