namespace NPOI.SL.UserModel
{
	public interface AutoShape<S, P>: TextShape<S,P>
		where S : Shape<S, P>
		where P : TextParagraph<S, P, TextRun>
	{

	}
}