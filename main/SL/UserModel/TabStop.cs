using System;

namespace NPOI.SL.UserModel
{
	public enum TabStopTypeEnum
	{
		LEFT,
		CENTER,
		RIGHT,
		DECIMAL
	}

	public interface TabStop
	{
		/**
     * Gets the position in points relative to the left side of the paragraph.
	     * 
	     * @return position in points
	     */
	    double getPositionInPoints();
	
	    /**
	     * Sets the position in points relative to the left side of the paragraph
	     *
	     * @param position position in points
	     */
	    void setPositionInPoints(double position);
	
	    TabStopType getType();
	
	    void setType(TabStopType type);
	}

	public class TabStopType
	{
		public int nativeId;
		public TabStopTypeEnum native;
		public int ooxmlId;

		public TabStopType(int nativeId, int ooxmlId)
		{
			this.nativeId = nativeId;
			this.ooxmlId = ooxmlId;
			this.native = (TabStopTypeEnum)nativeId;
		}
		public static TabStopType fromNativeId(int nativeId)
		{
			foreach (TabStopTypeEnum item in Enum.GetValues(typeof(TabStopTypeEnum)))
			{
				if (item == (TabStopTypeEnum)nativeId)
				{
					return new TabStopType(nativeId, nativeId +1);
				}
			}
			return null;
		}
		public static TabStopType fromOoxmlId(int ooxmlId)
		{
			foreach (TabStopTypeEnum item in Enum.GetValues(typeof(TabStopTypeEnum)))
			{
				if (item == (TabStopTypeEnum)ooxmlId-1)
				{
					return new TabStopType(ooxmlId - 1, ooxmlId);
				}
			}
			return null;
		}
	}
}