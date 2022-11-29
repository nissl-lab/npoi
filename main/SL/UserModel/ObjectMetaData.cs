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

using NPOI.Util;
using System;
using System.Collections.Generic;

namespace NPOI.SL.UserModel
{
	//using NPOI.HPSF.ClassID;
	//using NPOI.HPSF.ClassIDPredefined;

	public interface ObjectMetaData
	{
		/**
		 * @return the name of the OLE shape
		 */
		string getObjectName();

		/**
		 * @return the program id assigned to the OLE container application
		 */
		string getProgId();

		/**
		 * @return the storage classid of the OLE entry
		 */
		ClassID getClassID();

		/**
		 * @return the name of the OLE entry inside the oleObject#.bin
		 */
		string getOleEntry();
	}

	public enum ApplicationEnum
	{
		EXCEL_V8,
		EXCEL_V12,
		WORD_V8,
		WORD_V12,
		PDF,
		CUSTOM
	}

	public class Application
	{
		public string objectName;
		public string progId;
		public string oleEntry;
		public ClassID classId;

		public static readonly Dictionary<ApplicationEnum, (string objectName, string progId, string oleEntry, ClassIDPredefined classId)> ApplicationLookup =
			new Dictionary<ApplicationEnum, (string objectName, string progId, string oleEntry, ClassIDPredefined classId)>
			{
				{ ApplicationEnum.EXCEL_V8,("Worksheet", "Excel.Sheet.8", "Package", ClassIDPredefined.EXCEL_V8) },
				{ ApplicationEnum.EXCEL_V12, ("Worksheet", "Excel.Sheet.12", "Package", ClassIDPredefined.EXCEL_V12) },
				{ ApplicationEnum.WORD_V8, ("Document", "Word.Document.8", "Package", ClassIDPredefined.WORD_V8) },
				{ ApplicationEnum.WORD_V12, ("Document", "Word.Document.12", "Package", ClassIDPredefined.WORD_V12) },
				{ ApplicationEnum.PDF, ("PDF", "AcroExch.Document", "Contents", ClassIDPredefined.PDF) },
				{ ApplicationEnum.CUSTOM, (null, null, null, null) }
			};

		public String getObjectName()
		{
			return objectName;
		}

		public String getProgId()
		{
			return progId;
		}

		public String getOleEntry()
		{
			return oleEntry;
		}

		public ClassID getClassID()
		{
			return classId;
		}

		public Application(string objectName, string progId, string oleEntry, ClassIDPredefined classId)
		{
			this.objectName = objectName;
			this.progId = progId;
			this.classId = (classId == null) ? null : classId.getClassID();
			this.oleEntry = oleEntry;
		}

		public static Application lookup(string progId)
		{
			foreach (var item in ApplicationLookup)
			{
				if (item.Value.progId != null && item.Value.progId.Equals(progId))
				{
					return new Application(item.Value.objectName, item.Value.progId, item.Value.oleEntry, item.Value.classId);
				}
			}
			return null;
		}

		public ObjectMetaData GetMetaData()
		{
			throw new NotImplementedException();
		}
	}
}