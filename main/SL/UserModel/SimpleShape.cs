using System;
using System.Drawing;

namespace NPOI.SL.UserModel
{
	public interface SimpleShape<S, P>: Shape<S,P>, PlaceableShape<S, P>
		where S : Shape<S,P>
		where P : TextParagraph<S,P,TextRun>
	{
		FillStyle getFillStyle();
	    
	    LineDecoration getLineDecoration();
	    
	    StrokeStyle getStrokeStyle();
	    
	    /**
	     * Sets the line attributes.
	     * Possible attributes are Double (width), LineCap, LineDash, LineCompound, Color
	     * (implementations of PaintStyle aren't yet supported ...)
	     * 
	     * If no styles are given, the line will be hidden
	     *
	     * @param styles the line attributes
	     */
	    void setStrokeStyle(params object[] styles);
	
	    //CustomGeometry getGeometry();
	
	    ShapeType getShapeType();
	    
		void setShapeType(ShapeType type);
	
	    /**
	     * @return the placeholder or null if none is assigned
	     * @see #setPlaceholder(Placeholder)
	     */
	    Placeholder getPlaceholder();
	    
	    /**
	     * Specifies that the corresponding shape should be represented by the generating application
	     * as a placeholder. When a shape is considered a placeholder by the generating application
	     * it can have special properties to alert the user that they may enter content into the shape.
	     * 
	     * @param placeholder the placeholder or null to remove the reference to the placeholder
	     */
	    void setPlaceholder(Placeholder placeholder);
	
	    /**
	     * @return an accessor for placeholder details
	     *
	     * @since POI 4.0.0
	     */
	    PlaceholderDetails getPlaceholderDetails();
	
	    /**
	     * Checks if the shape is a placeholder.
	     * (placeholders aren't normal shapes, they are visible only in the Edit Master mode)
	     *
	     * @return {@code true} if the shape is a placeholder
	     * 
	     * @since POI 4.0.0
	     */
	    bool isPlaceholder();
	    
	    
	    //Shadow<S, P> getShadow();
	
	    /**
	     * Returns the solid color fill.
	     *
	     * @return solid fill color of null if not set or fill color
	     * is not solid (pattern or gradient)
	     */
	    Color getFillColor();
	
	    /**
	     * Specifies a solid color fill. The shape is filled entirely with the
	     * specified color.
	     *
	     * @param color the solid color fill. The value of <code>null</code> unsets
	     *              the solid fill attribute from the underlying implementation
	     */
	    void setFillColor(Color color);
	
	    /**
	     * Returns the hyperlink assigned to this shape
	     *
	     * @return the hyperlink assigned to this shape
	     * or <code>null</code> if not found.
	     * 
	     * @since POI 3.14-Beta1
	     */
	    Hyperlink<S, P> getHyperlink();
	    
	    /**
	     * Creates a hyperlink and asigns it to this shape.
	     * If the shape has already a hyperlink assigned, return it instead
	     *
	     * @return the hyperlink assigned to this shape
	     * 
	     * @since POI 3.14-Beta1
	     */
	    Hyperlink<S, P> createHyperlink();
	}
}