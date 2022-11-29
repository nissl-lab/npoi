using System;

namespace NPOI.SL.UserModel
{
	public interface PlaceableShape<S, P>
		where S : Shape<S, P>
		where P : TextParagraph<S, P, TextRun>
	{
		ShapeContainer<S, P> getParent();
	    
	    /**
	     * @return the sheet this shape belongs to
	     */
	    Sheet<S, P> getSheet();
	    
	    /**
	     * @return the position of this shape within the drawing canvas.
	     *         The coordinates are expressed in points
	     */
	    //Rectangle2D getAnchor();
	
	    /**
	     * @param anchor the position of this shape within the drawing canvas.
	     *               The coordinates are expressed in points
	     */
	    //void setAnchor(Rectangle2D anchor);
	
	    /**
	     * Rotation angle in degrees
	     * <p>
	     * Positive angles are clockwise (i.e., towards the positive y axis);
	     * negative angles are counter-clockwise (i.e., towards the negative y axis).
	     * </p>
	     *
	     * @return rotation angle in degrees
	     */
	    double getRotation();
	
	    /**
	     * Rotate this shape.
	     * <p>
	     * Positive angles are clockwise (i.e., towards the positive y axis);
	     * negative angles are counter-clockwise (i.e., towards the negative y axis).
	     * </p>
	     *
	     * @param theta the rotation angle in degrees.
	     */
	    void setRotation(double theta);
	
	    /**
	     * @param flip whether the shape is horizontally flipped
	     */
	    void setFlipHorizontal(bool flip);
	
	    /**
	     * Whether the shape is vertically flipped
	     *
	     * @param flip whether the shape is vertically flipped
	     */
	    void setFlipVertical(bool flip);
	
	    /**
	     * Whether the shape is horizontally flipped
	     *
	     * @return whether the shape is horizontally flipped
	     */
	    bool getFlipHorizontal();
	
	    /**
	     * Whether the shape is vertically flipped
	     *
	     * @return whether the shape is vertically flipped
	     */
	    bool getFlipVertical();
	}
}