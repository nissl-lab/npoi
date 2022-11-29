namespace NPOI.SL.UserModel
{
	public interface Background<S, P>: Shape<S,P>
		where S : Shape<S,P>
		where P : TextParagraph<S,P,TextRun>
	{
		FillStyle getFillStyle();
	}
}