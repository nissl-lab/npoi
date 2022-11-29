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

using NPOI.POIFS.FileSystem;
using NPOI.SS.Formula.Functions;
using NPOI.Util;
using System;
using System.IO;

namespace NPOI.SL.UserModel
{
    public abstract class ObjectShape<S, P>: Shape<S, P>, PlaceableShape<S, P>
        where S : Shape<S, P>
        where P : TextParagraph<S, P, TextRun>
    {
		/**
         * Returns the picture data for this picture.
         *
         * @return the picture data for this picture.
         */
		public abstract PictureData getPictureData();

		/**
		 * Returns the ProgID that stores the OLE Programmatic Identifier.
		 * A ProgID is a string that uniquely identifies a given object, for example,
		 * "Word.Document.8" or "Excel.Sheet.8".
		 *
		 * @return the ProgID
		 */
		public abstract string getProgId();

		/**
		 * Returns the full name of the embedded object,
		 *  e.g. "Microsoft Word Document" or "Microsoft Office Excel Worksheet".
		 *
		 * @return the full name of the embedded object
		 */
		public abstract string getFullName();

		/**
		 * Updates the ole data. If there wasn't an object registered before, a new
		 * ole embedding is registered in the parent slideshow.<p>
		 *
		 * For HSLF this needs to be a {@link POIFSFileSystem} stream.
		 *
		 * @param application a preset application enum
		 * @param metaData or a custom metaData object, can be {@code null} if the application has been set
		 *
		 * @return an {@link OutputStream} which receives the new data, the data will be persisted on {@code close()}
		 *
		 * @throws IOException if the linked object data couldn't be found or a new object data couldn't be initialized
		 */
		public abstract OutputStream updateObjectData(Application application, ObjectMetaData metaData);

        /**
         * Reads the ole data as stream - the application specific stream is served
         * The {@link #readObjectDataRaw() raw data} serves the outer/wrapped object, which is usually a
         * {@link POIFSFileSystem} stream, whereas this method return the unwrapped entry
         *
         * @return an {@link InputStream} which serves the object data
         *
         * @throws IOException if the linked object data couldn't be found
         */
        public InputStream readObjectData()
		{
			string progId = getProgId();
            if (progId == null) 
            {
                throw new InvalidOperationException(
                    "Ole object hasn't been initialized or provided in the source xml. " +
                    "use updateObjectData() first or check the corresponding slideXXX.xml");
	        }

	        Application app = Application.lookup(progId);

            using (UnsynchronizedByteArrayOutputStream bos = new UnsynchronizedByteArrayOutputStream())
            {
                using (InputStream _is = FileMagic.prepareToCheckMagic(readObjectDataRaw()))
                {
                    FileMagic fm = FileMagic.valueOf(_is);
                    if (fm == FileMagic.OLE2)
                    {
                        using (POIFSFileSystem poifs = new POIFSFileSystem(_is))
                        {
                            string[] names = {
                                (app == null) ? null : app.getMetaData().getOleEntry(),
                                // fallback to the usual suspects
                                "Package",
                                "Contents",
                                "CONTENTS",
                                "CONTENTSV30",
                            };
                            DirectoryNode root = poifs.getRoot();
                            string entryName = null;
                            foreach (string n in names)
                            {
                                if (root.hasEntry(n))
                                {
                                    entryName = n;
                                    break;
                                }
                            }
                            if (entryName == null)
                            {
                                poifs.writeFilesystem(bos);
                            }
                            else
                            {
                                using (InputStream is2 = poifs.createDocumentInputStream(entryName))
                                {
                                    IOUtils.copy(is2, bos);
                                }
                            }
                        }
                    } else
                    {
                        IOUtils.copy(is, bos);
                    }
                    return bos.toInputStream();
                }
            }
        }
    
        /**
         * Convenience method to return the raw data as {@code InputStream}
         *
         * @return the raw data stream
         *
         * @throws IOException if the data couldn't be retrieved
         */
        InputStream readObjectDataRaw() throws IOException
        {
            return getObjectData().getInputStream();
        }

        /**
         * @return the data object
         */
        ObjectData getObjectData();
    }
}