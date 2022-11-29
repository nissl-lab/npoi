using NPOI.HPSF;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;

namespace NPOI.SL.UserModel
{
	public class Placeholder
	{
		public int nativeSlideId;
	    public int nativeSlideMasterId;
	    public int nativeNotesId;
	    public int nativeNotesMasterId;
	    public int ooxmlId;
		public readonly string nativeEnum;

		public static readonly Dictionary<string, int[]> PlaceholderEnum = new Dictionary<string, int[]>() {
			/**
			* No placeholder shape.
			*/
			{ "NONE", new int[]{0,0,0,0,0 } },
			/**
			* Title text placeholder shape.
			*/
			{ "TITLE", new int[]{ 13,1,1,1,1 } },
			/**
			* Body text placeholder shape.
			*/
			{ "BODY", new int[]{14,2,12,6,2 } },
			/**
			* Center title text placeholder shape.
			*/
			{ "CENTERED_TITLE", new int[]{15,3,3,3,3 } },
			/**
			* Sub-title text placeholder shape.
			*/
			{ "SUBTITLE", new int[]{16,4,4,4,4 } },
			/**
			* Date placeholder shape.
			*/
			{ "DATETIME", new int[]{7,7,7,7,5 } },
			/**
			* Slide number placeholder shape.
			*/
			{ "SLIDE_NUMBER", new int[]{8,8,8,8,6 } },
			/**
			* Footer placeholder shape.
			*/
			{ "FOOTER", new int[]{9,9,9,9,7 } },
			/**
			*  Header placeholder shape.
			*/
			{ "HEADER", new int[]{10,10,10,10,8 } },
			/**
			* Object placeholder shape.
			*/
			{ "CONTENT", new int[]{19,19,19,19,9 } },
			/**
			* Graph object placeholder shape.
			*/
			{ "CHART",new int[]{20,20,20,20,10 } },
			/**
			* Table object placeholder shape.
			*/
			{ "TABLE", new int[]{21,21,21,21,11 } },
			/**
			* Clipart object placeholder shape.
			*/
			{ "CLIP_ART",new int[]{22,22,22,22,12 } },
			/**
			* Organization chart object placeholder shape.
			*/
			{ "DGM", new int[]{23,23,23,23,13 } },
			/**
			* Media object placeholder shape.
			*/
			{ "MEDIA", new int[]{24,24,24,24,14 } },
			/**
			* Slide image placeholder shape.
			*/
			{ "SLIDE_IMAGE", new int[]{11,11,11,5,15 } },
			/**
			* Picture object placeholder shape.
			*/
			{ "PICTURE", new int[]{26,26,26,26,16 } },
			/**
			* Vertical object placeholder shape.
			*/
			{ "VERTICAL_OBJECT", new int[]{25,25,25,25,-2 } },
			/**
			* Vertical title text placeholder shape.
			*/
			{ "VERTICAL_TEXT_TITLE", new int[]{17,17,17,17,-2 } },
			/**
			* Vertical body text placeholder shape.
			*/
			{ "VERTICAL_TEXT_BODY", new int[]{18,18,18,18,-2 } }
		};
		
	    
	    public Placeholder(int nativeSlideId, int nativeSlideMasterId, int nativeNotesId, int nativeNotesMasterId, int ooxmlId)
		{
			this.nativeSlideId = nativeSlideId;
			this.nativeSlideMasterId = nativeSlideMasterId;
			this.nativeNotesId = nativeNotesId;
			this.nativeNotesMasterId = nativeNotesMasterId;
			this.ooxmlId = ooxmlId;
			this.nativeEnum = PlaceholderEnum.First( i => i.Value[0] == nativeSlideId).Key;
		}
	    
	    public static Placeholder lookupNativeSlide(int nativeId)
		{
			return lookupNative(nativeId, 0);
		}
	
	    public static Placeholder lookupNativeSlideMaster(int nativeId)
		{
			return lookupNative(nativeId, 1);
		}
	
	    public static Placeholder lookupNativeNotes(int nativeId)
		{
			return lookupNative(nativeId, 2);
		}
	
	    public static Placeholder lookupNativeNotesMaster(int nativeId)
		{
			return lookupNative(nativeId, 3);
		}
	
	    
	    private static Placeholder lookupNative(int nativeId, int type)
		{
			foreach (var item in PlaceholderEnum.Values)
			{
				if (
					type == 0 && item[0] == nativeId ||
					type == 1 && item[1] == nativeId ||
					type == 2 && item[2] == nativeId ||
					type == 3 && item[3] == nativeId
				) {
				return new Placeholder(item[0], item[1], item[2], item[3], item[4]);
				}
			}
	        return null;
	    }
	    
	    public static Placeholder lookupOoxml(int ooxmlId)
		{
			foreach (var item in PlaceholderEnum.Values)
			{
				if (item[4] == ooxmlId)
				{
					return new Placeholder(item[0], item[1], item[2], item[3], item[4]);
				}
			}
			return null;
		}
	}
}