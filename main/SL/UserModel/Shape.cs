using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SL.UserModel
{
	public interface Shape<S, P>
		where S : Shape<S, P>
		where P: TextParagraph<S, P, TextRun>
	{
		ShapeContainer<S, P> GetParent();

		/**
	    * @return the sheet this shape belongs to
	    */
	   Sheet<S, P> GetSheet();
	}
}
