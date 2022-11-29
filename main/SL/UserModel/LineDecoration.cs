using System;

namespace NPOI.SL.UserModel
{
	/**
	*  Represents the shape decoration that appears at the ends of lines.
	*/
	public enum DecorationShapeEnum
	{
		NONE,
		TRIANGLE,
		STEALTH,
		DIAMOND,
		OVAL,
		ARROW
	}

	public class DecorationShape
	{
		public int nativeId;
		public int ooxmlId;
		public DecorationShapeEnum native;

		DecorationShape(int nativeId, int ooxmlId)
		{
			this.nativeId = nativeId;
			this.ooxmlId = ooxmlId;
			this.native = (DecorationShapeEnum)nativeId;
		}

		public static DecorationShape fromNativeId(int nativeId)
		{
			foreach (DecorationShapeEnum item in Enum.GetValues(typeof(DecorationShapeEnum)))
			{
				if ((int)item == nativeId) return new DecorationShape(nativeId, nativeId+1);
			}
			return null;
		}

		public static DecorationShape fromOoxmlId(int ooxmlId)
		{
			foreach (DecorationShapeEnum item in Enum.GetValues(typeof(DecorationShapeEnum)))
			{
				if ((int)item == ooxmlId -1) return new DecorationShape(ooxmlId-1, ooxmlId);
			}
			return null;
		}
	}

	public enum DecorationSizeEnum
	{
		SMALL,
		MEDIUM,
		LARGE
	}
	
	public class DecorationSize
	{ 
		public int nativeId;
		public int ooxmlId;
		public DecorationSizeEnum native;

		public DecorationSize(int nativeId, int ooxmlId)
		{
			this.nativeId = nativeId;
			this.ooxmlId = ooxmlId;
			this.native = (DecorationSizeEnum)nativeId;
		}
	
		public static DecorationSize fromNativeId(int nativeId)
		{
			foreach (DecorationSizeEnum item in Enum.GetValues(typeof(DecorationSizeEnum)))
			{
				if ((int)item == nativeId) return new DecorationSize(nativeId, nativeId+1);
			}
			return null;
		}
	
		public static DecorationSize fromOoxmlId(int ooxmlId)
		{
			foreach (DecorationSizeEnum item in Enum.GetValues(typeof(DecorationSizeEnum)))
			{
				if ((int)item == ooxmlId) return new DecorationSize(ooxmlId-1, ooxmlId);
			}
			return null;
		}
	}


	public interface LineDecoration
	{
		/**
	     * @return the line start shape
	     */
	    DecorationShape getHeadShape();
	
	    /**
	     * @return the width of the start shape
	     */
	    DecorationSize getHeadWidth();
	
	    /**
	     * @return the length of the start shape
	     */
	    DecorationSize getHeadLength();
	
	    /**
	     * @return the line end shape
	     */
	    DecorationShape getTailShape();
	
	    /**
	     * @return the width of the end shape
	     */
	    DecorationSize getTailWidth();
	
	    /**
	     * @return the length of the end shape
	     */
	    DecorationSize getTailLength();
	}
}
