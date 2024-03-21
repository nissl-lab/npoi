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

namespace NPOI.SS.Extractor

{
    using static NPOI.Util.StringUtil;


    using NPOI.HPSF;
    using NPOI.POIFS.FileSystem;
    using NPOI.SS.UserModel;
    using NPOI.Util;
    using NPOI.HSSF.UserModel;
    using System.Collections.Generic;
    using System;
    using System.IO;
    using System.Text;
    using System.Collections;
    using NPOI.XSSF.UserModel;

    /// <summary>
    /// This extractor class tries to identify various embedded documents within Excel files
    /// and provide them via a common interface, i.e. the EmbeddedData instances
    /// </summary>
    public class EmbeddedExtractor : IEnumerable<EmbeddedExtractor>
    {
        private static POILogger LOG = POILogFactory.GetLogger(typeof(EmbeddedExtractor));

        // contentType
        private static String CONTENT_TYPE_BYTES = "binary/octet-stream";
        private static String CONTENT_TYPE_PDF = "application/pdf";
        private static String CONTENT_TYPE_DOC = "application/msword";
        private static String CONTENT_TYPE_XLS = "application/vnd.ms-excel";

        /// <summary>
        /// </summary>
        /// <return>list of known extractors, if you provide custom extractors, override this method/// </return>
        public IEnumerator<EmbeddedExtractor> GetEnumerator()
        {
            EmbeddedExtractor[] ee = {
            new Ole10Extractor(), new PdfExtractor(), new BiffExtractor(), new OOXMLExtractor(), new FsExtractor()
        };
            return Arrays.AsList(ee).GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            throw new NotImplementedException();
        }

        public virtual EmbeddedData ExtractOne(DirectoryNode src)
        {

            foreach (EmbeddedExtractor ee in this)
            {
                if (ee.CanExtract(src))
                {
                    return ee.Extract(src);
                }
            }
            return null;
        }

        public virtual EmbeddedData ExtractOne(IPicture src)
        {

            foreach (EmbeddedExtractor ee in this)
            {
                if (ee.CanExtract(src))
                {
                    return ee.Extract(src);
                }
            }
            return null;
        }

        public virtual List<EmbeddedData> ExtractAll(ISheet sheet)
        {

            IDrawing<IShape> patriarch = sheet.DrawingPatriarch;
            if (null == patriarch)
            {
                return new List<EmbeddedData>();
            }
            List<EmbeddedData> embeddings = new List<EmbeddedData>();
            ExtractAll(patriarch, embeddings);
            return embeddings;
        }

        protected virtual void ExtractAll<T>(IShapeContainer<T> parent, List<EmbeddedData> embeddings) where T : class, IShape
        {
            foreach (IShape shape in parent)
            {
                EmbeddedData data = null;
                if (shape is IObjectData)
                {
                    IObjectData od = (IObjectData)shape;
                    try
                    {
                        if (od.HasDirectoryEntry())
                        {
                            data = ExtractOne((DirectoryNode)od.Directory);
                        }
                        else
                        {
                            String contentType = CONTENT_TYPE_BYTES;
                            if (od is XSSFObjectData)
                            {
                                contentType = ((XSSFObjectData)od).GetObjectPart().ContentType;
                            }
                            data = new EmbeddedData(od.FileName, od.ObjectData, contentType);
                        }
                    }
                    catch (Exception e)
                    {
                        LOG.Log(POILogger.WARN, "Entry not found / Readable - ignoring OLE embedding", e);
                    }
                }
                else if (shape is IPicture)
                {
                    data = ExtractOne((IPicture)shape);
                }
                else if (shape is IShapeContainer<T>)
                {
                    ExtractAll((IShapeContainer<T>)shape, embeddings);
                }

                if (data == null)
                {
                    continue;
                }

                data.Shape = (/*setter*/shape);
                String filename = data.Filename;
                String extension = (filename == null || filename.LastIndexOf('.') == -1) ? ".bin" : filename.Substring(filename.LastIndexOf('.'));

                // try to find an alternative name
                if (filename == null || "".Equals(filename) || filename.StartsWith("MBD") || filename.StartsWith("Root Entry"))
                {
                    filename = shape.ShapeName;
                    if (filename != null)
                    {
                        filename += extension;
                    }
                }
                // default to dummy name
                if (filename == null || "".Equals(filename))
                {
                    filename = "picture_" + embeddings.Count + extension;
                }
                filename = filename.Trim();
                data.Filename = (/*setter*/filename);

                embeddings.Add(data);
            }
        }


        public virtual bool CanExtract(DirectoryNode source)
        {
            return false;
        }

        public virtual bool CanExtract(IPicture source)
        {
            return false;
        }

        public virtual EmbeddedData Extract(DirectoryNode dn)
        {

            //assert(canExtract(dn));
            POIFSFileSystem dest = new POIFSFileSystem();
            copyNodes(dn, dest.Root);
            // start with a reasonable big size
            MemoryStream bos = new MemoryStream(20000);
            dest.WriteFileSystem(bos);
            dest.Close();

            return new EmbeddedData(dn.Name, bos.ToArray(), CONTENT_TYPE_BYTES);
        }

        public virtual EmbeddedData Extract(IPicture source)
        {

            return null;
        }

        public class Ole10Extractor : EmbeddedExtractor
        {
            public override bool CanExtract(DirectoryNode dn)
            {
                ClassID clsId = dn.StorageClsid;
                return ClassID.OLE10_PACKAGE.Equals(clsId);
            }
            public override EmbeddedData Extract(DirectoryNode dn)
            {

                try
                {
                    // TODO: inspect the CompObj record for more details, i.e. the content type
                    Ole10Native ole10 = Ole10Native.CreateFromEmbeddedOleObject(dn);
                    return new EmbeddedData(ole10.FileName, ole10.DataBuffer, CONTENT_TYPE_BYTES);
                }
                catch (Ole10NativeException e)
                {
                    throw new IOException("", e);
                }
            }
        }

        class PdfExtractor : EmbeddedExtractor
        {
            static ClassID PdfClassID = new ClassID("{B801CA65-A1FC-11D0-85AD-444553540000}");
            public override bool CanExtract(DirectoryNode dn)
            {
                ClassID clsId = dn.StorageClsid;
                return (PdfClassID.Equals(clsId)
                || dn.HasEntry("CONTENTS"));
            }
            public override EmbeddedData Extract(DirectoryNode dn)
            {

                MemoryStream bos = new MemoryStream();
                InputStream is1 = dn.CreateDocumentInputStream("CONTENTS");
                IOUtils.Copy(is1, bos);
                is1.Close();
                return new EmbeddedData(dn.Name + ".pdf", bos.ToArray(), CONTENT_TYPE_PDF);
            }
            public override bool CanExtract(IPicture source)
            {
                IPictureData pd = source.PictureData;
                return (pd != null && pd.PictureType == PictureType.EMF);
            }

            /// <summary>
            /// <para>
            /// Mac Office encodes embedded objects inside the picture, e.g. PDF is part of an EMF.
            /// If an embedded stream is inside an EMF picture, this method extracts the payload.
            /// </para>
            /// <para>
            /// </para>
            /// </summary>
            /// <return>embedded data in an EMF picture or null if none is found/// </return>
            public override EmbeddedData Extract(IPicture source)
            {

                // check for emf+ embedded pdf (poor mans style :( )
                // Mac Excel 2011 embeds pdf files with this method.
                IPictureData pd = source.PictureData;
                if (pd == null || pd.PictureType != PictureType.EMF)
                {
                    return null;
                }

                // TODO: investigate if this is just an EMF-hack or if other formats are also embedded in EMF
                byte[] pictureBytes = pd.Data;
                int idxStart = IndexOf(pictureBytes, 0, Encoding.ASCII.GetBytes("%PDF-")); //.GetBytes(LocaleUtil.CHARSET_1252)
                if (idxStart == -1)
                {
                    return null;
                }

                int idxEnd = IndexOf(pictureBytes, idxStart, Encoding.ASCII.GetBytes("%%EOF")); // .GetBytes(LocaleUtil.CHARSET_1252)
                if (idxEnd == -1)
                {
                    return null;
                }

                int pictureBytesLen = idxEnd - idxStart + 6;
                byte[] pdfBytes = new byte[pictureBytesLen];
                System.Array.Copy(pictureBytes, idxStart, pdfBytes, 0, pictureBytesLen);
                String filename = source.ShapeName.Trim();
                if (!filename.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                {
                    filename += ".pdf";
                }
                return new EmbeddedData(filename, pdfBytes, CONTENT_TYPE_PDF);
            }


        }

        class OOXMLExtractor : EmbeddedExtractor
        {
            public override bool CanExtract(DirectoryNode dn)
            {
                return dn.HasEntry("package");
            }
            public override EmbeddedData Extract(DirectoryNode dn)
            {


                ClassID clsId = dn.StorageClsid;

                String contentType, ext;
                if (ClassID.WORD2007.Equals(clsId))
                {
                    ext = ".docx";
                    contentType = "application/vnd.Openxmlformats-officedocument.wordProcessingml.document";
                }
                else if (ClassID.WORD2007_MACRO.Equals(clsId))
                {
                    ext = ".docm";
                    contentType = "application/vnd.ms-word.document.macroEnabled.12";
                }
                else if (ClassID.EXCEL2007.Equals(clsId) || ClassID.EXCEL2003.Equals(clsId) || ClassID.EXCEL2010.Equals(clsId))
                {
                    ext = ".xlsx";
                    contentType = "application/vnd.Openxmlformats-officedocument.spreadsheetml.sheet";
                }
                else if (ClassID.EXCEL2007_MACRO.Equals(clsId))
                {
                    ext = ".xlsm";
                    contentType = "application/vnd.ms-excel.sheet.macroEnabled.12";
                }
                else if (ClassID.EXCEL2007_XLSB.Equals(clsId))
                {
                    ext = ".xlsb";
                    contentType = "application/vnd.ms-excel.sheet.binary.macroEnabled.12";
                }
                else if (ClassID.POWERPOINT2007.Equals(clsId))
                {
                    ext = ".pptx";
                    contentType = "application/vnd.Openxmlformats-officedocument.presentationml.presentation";
                }
                else if (ClassID.POWERPOINT2007_MACRO.Equals(clsId))
                {
                    ext = ".ppsm";
                    contentType = "application/vnd.ms-powerpoint.slideShow.macroEnabled.12";
                }
                else
                {
                    ext = ".zip";
                    contentType = "application/zip";
                }

                DocumentInputStream dis = dn.CreateDocumentInputStream("package");
                byte[] data = IOUtils.ToByteArray(dis);
                dis.Close();

                return new EmbeddedData(dn.Name + ext, data, contentType);
            }
        }

        class BiffExtractor : EmbeddedExtractor
        {
            public override bool CanExtract(DirectoryNode dn)
            {
                return CanExtractExcel(dn) || CanExtractWord(dn);
            }

            protected bool CanExtractExcel(DirectoryNode dn)
            {
                ClassID clsId = dn.StorageClsid;
                return (ClassID.EXCEL95.Equals(clsId)
                    || ClassID.EXCEL97.Equals(clsId)
                    || dn.HasEntry("Workbook") /*...*/);
            }

            protected bool CanExtractWord(DirectoryNode dn)
            {
                ClassID clsId = dn.StorageClsid;
                return (ClassID.WORD95.Equals(clsId)
                    || ClassID.WORD97.Equals(clsId)
                    || dn.HasEntry("WordDocument"));
            }
            public override EmbeddedData Extract(DirectoryNode dn)
            {

                EmbeddedData ed = base.Extract(dn);
                if (CanExtractExcel(dn))
                {
                    ed.Filename = (/*setter*/dn.Name + ".xls");
                    ed.ContentType = (/*setter*/CONTENT_TYPE_XLS);
                }
                else if (CanExtractWord(dn))
                {
                    ed.Filename = (/*setter*/dn.Name + ".doc");
                    ed.ContentType = (/*setter*/CONTENT_TYPE_DOC);
                }

                return ed;
            }
        }

        class FsExtractor : EmbeddedExtractor
        {
            public override bool CanExtract(DirectoryNode dn)
            {
                return true;
            }
            public override EmbeddedData Extract(DirectoryNode dn)
            {

                EmbeddedData ed = base.Extract(dn);
                ed.Filename = (/*setter*/dn.Name + ".ole");
                // TODO: read the content type from CombObj stream
                return ed;
            }
        }

        protected static void copyNodes(DirectoryNode src, DirectoryNode dest)
        {

            foreach (Entry e in src)
            {
                if (e is DirectoryNode)
                {
                    DirectoryNode srcDir = (DirectoryNode)e;
                    DirectoryNode destDir = (DirectoryNode)dest.CreateDirectory(srcDir.Name);
                    destDir.StorageClsid = (/*setter*/srcDir.StorageClsid);
                    copyNodes(srcDir, destDir);
                }
                else
                {
                    InputStream is1 = src.CreateDocumentInputStream(e);
                    try
                    {
                        dest.CreateDocument(e.Name, is1);
                    }
                    finally
                    {
                        is1.Close();
                    }
                }
            }
        }



        /// <summary>
        /// Knuth-Morris-Pratt Algorithm for Pattern Matching
        /// Finds the first occurrence of the pattern in the text.
        /// </summary>
        private static int IndexOf(byte[] data, int offset, byte[] pattern)
        {
            int[] failure = computeFailure(pattern);

            int j = 0;
            if (data.Length == 0)
            {
                return -1;
            }

            for (int i = offset; i < data.Length; i++)
            {
                while (j > 0 && pattern[j] != data[i])
                {
                    j = failure[j - 1];
                }
                if (pattern[j] == data[i]) { j++; }
                if (j == pattern.Length)
                {
                    return i - pattern.Length + 1;
                }
            }
            return -1;
        }

        /// <summary>
        /// Computes the failure function using a boot-strapping Process,
        /// where the pattern is matched against itself.
        /// </summary>
        private static int[] computeFailure(byte[] pattern)
        {
            int[] failure = new int[pattern.Length];

            int j = 0;
            for (int i = 1; i < pattern.Length; i++)
            {
                while (j > 0 && pattern[j] != pattern[i])
                {
                    j = failure[j - 1];
                }
                if (pattern[j] == pattern[i])
                {
                    j++;
                }
                failure[i] = j;
            }

            return failure;
        }

        
    }
}

