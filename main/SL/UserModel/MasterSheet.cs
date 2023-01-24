namespace NPOI.SL.UserModel
{
	public interface MasterSheet<S, P>: Sheet<S,P>
		where S : Shape<S,P>
		where P : TextParagraph<S,P,TextRun>
	{
		/**
	     * Return the placeholder shape for the specified type
	     * 
	     * @return the shape or {@code null} if it is not defined in this mastersheet
	     * 
	     * @since POI 4.0.0
	     */
	    SimpleShape<S, P> GetPlaceholder(Placeholder type);
	}
}