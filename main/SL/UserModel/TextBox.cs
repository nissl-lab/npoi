using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SL.UserModel
{
	public interface TextBox<S,P>: AutoShape<S,P>
		where S : Shape<S,P>
		where P : TextParagraph<S,P, TextRun>
	{
	}
}
