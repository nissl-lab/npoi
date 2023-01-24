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

using NPOI.Common.UserModel.Fonts;
using NPOI.Util;
using System;
using System.Collections.Generic;

namespace NPOI.HSLF.Record
{
	public class FontEmbeddedData : Record
	{
		public override Record[] GetChildRecords()
		{
			throw new NotImplementedException();
		}

		public FontHeader GetFontHeader()
		{
			throw new NotImplementedException();
		}

		public override IDictionary<string, Func<object>> GetGenericProperties()
		{
			throw new NotImplementedException();
		}

		public override long GetRecordType()
		{
			throw new NotImplementedException();
		}

		public override bool IsAnAtom()
		{
			throw new NotImplementedException();
		}

		public override void WriteOut(OutputStream o)
		{
			throw new NotImplementedException();
		}
	}
}