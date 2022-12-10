using NPOI.Common.UserModel;
using NPOI.HSLF.Record;
using NPOI.SL.UserModel;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HSLF.Model
{
	public class TextPFException9: GenericRecord
	{
		private static  AutoNumberingScheme DEFAULT_AUTONUMBER_SCHEME = AutoNumberingScheme.arabicPeriod;
    private static  short DEFAULT_START_NUMBER = 1;

    //private  byte mask1;
    //private  byte mask2;
    private  byte mask3;
		private  byte mask4;
		private  short bulletBlipRef;
    private  short fBulletHasAutoNumber;
    private  AutoNumberingScheme autoNumberScheme;
    private  short autoNumberStartNumber;
    private  int recordLength;
		public TextPFException9( byte[] source,  int startIndex)
		{ // NOSONAR
		  //this.mask1 = source[startIndex];
		  //this.mask2 = source[startIndex + 1];
			this.mask3 = source[startIndex + 2];
			this.mask4 = source[startIndex + 3];
			int length = 4;
			int index = startIndex + 4;
			if (0 == (mask3 & (byte)0x80))
			{
				this.bulletBlipRef = 0;
			}
			else
			{
				this.bulletBlipRef = LittleEndian.GetShort(source, index);
				index += 2;
				length = 6;
			}
			if (0 == (mask4 & 2))
			{
				this.fBulletHasAutoNumber = 0;
			}
			else
			{
				this.fBulletHasAutoNumber = LittleEndian.GetShort(source, index);
				index += 2;
				length += 2;
			}
			if (0 == (mask4 & 1))
			{
				this.autoNumberScheme = null;
				this.autoNumberStartNumber = 0;
			}
			else
			{
				this.autoNumberScheme = AutoNumberingScheme.ForNativeID(LittleEndian.GetShort(source, index));
				index += 2;
				this.autoNumberStartNumber = LittleEndian.GetShort(source, index);
				index += 2;
				length += 4;
			}
			this.recordLength = length;
		}
		public short GetBulletBlipRef()
		{
			return bulletBlipRef;
		}
		public short GetfBulletHasAutoNumber()
		{
			return fBulletHasAutoNumber;
		}
		public AutoNumberingScheme GetAutoNumberScheme()
		{
			if (autoNumberScheme != null)
			{
				return autoNumberScheme;
			}
			return HasBulletAutoNumber() ? DEFAULT_AUTONUMBER_SCHEME : null;
		}

		public short GetAutoNumberStartNumber()
		{
			if (autoNumberStartNumber != null)
			{
				return autoNumberStartNumber;
			}
			return (short)(HasBulletAutoNumber() ? DEFAULT_START_NUMBER : 0);
		}

		private bool HasBulletAutoNumber()
		{
			 short one = 1;
			return one.Equals(fBulletHasAutoNumber);
		}

		public int GetRecordLength()
		{
			return recordLength;
		}
		//@Override
	public override string ToString()
		{
			 StringBuilder sb = new StringBuilder("Record length: ").Append(this.recordLength).Append(" bytes\n");
			sb.Append("bulletBlipRef: ").Append(this.bulletBlipRef).Append("\n");
			sb.Append("fBulletHasAutoNumber: ").Append(this.fBulletHasAutoNumber).Append("\n");
			sb.Append("autoNumberScheme: ").Append(this.autoNumberScheme).Append("\n");
			sb.Append("autoNumberStartNumber: ").Append(this.autoNumberStartNumber).Append("\n");
			return sb.ToString();
		}

		//@Override
	public IDictionary<string, Func<T>> GetGenericProperties<T>()
		{
			return (IDictionary<string, Func<T>>)GenericRecordUtil.GetGenericProperties(
				"bulletBlipRef", ()=> GetBulletBlipRef(),
				"bulletHasAutoNumber", () => HasBulletAutoNumber(),
				"autoNumberScheme", GetAutoNumberScheme,
				"autoNumberStartNumber", ()=> GetAutoNumberStartNumber()
			);
		}

		public RecordTypes GetGenericRecordType()
		{
			throw new NotImplementedException();
		}

		public IList<GenericRecord> GetGenericChildren()
		{
			throw new NotImplementedException();
		}
	}
}
