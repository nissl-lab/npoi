using System;
using System.Collections.Generic;

namespace NPOI.SL.UserModel
{
	public interface ShapeContainer<S, P>: IEnumerable<S>
		where S : Shape<S,P>
		where P : TextParagraph<S,P, TextRun>
	{
		/**
	     * Returns an list containing all of the elements in this container in proper
	     * sequence (from first to last element).
	     *
	     * @return an list containing all of the elements in this container in proper
	     *         sequence
	     */
	    List<S> getShapes();

		void addShape(S shape);

		/**
	     * Removes the specified shape from this sheet, if it is present
	     * (optional operation).  If this sheet does not contain the element,
	     * it is unchanged.
	     *
	     * @param shape the shape to be removed from this sheet, if present
	     * @return {@code true} if this sheet contained the specified element
	     * @throws IllegalArgumentException if the type of the specified shape
	     *         is incompatible with this sheet (optional)
	     */
	    bool removeShape(S shape);

		/**
	     * create a new shape with a predefined geometry and add it to this shape container
	     */
	    AutoShape<S, P> createAutoShape();
	
	    /**
	     * create a new shape with a custom geometry
	     */
	    //FreeformShape<S, P> createFreeform();
	
	    /**
	     * create a text box
	     */
	    TextBox<S, P> createTextBox();
	
	    /**
	     * create a connector
	     */
	    //ConnectorShape<S, P> createConnector();
	
	    /**
	     * create a group of shapes belonging to this container
	     */
	    //GroupShape<S, P> createGroup();
	
	    /**
	     * create a picture belonging to this container
	     */
	    //PictureShape<S, P> createPicture(PictureData pictureData);
	
	    /**
	     * Create a new Table of the given number of rows and columns
	     *
	     * @param numRows the number of rows
	     * @param numCols the number of columns
	     */
	    //TableShape<S, P> createTable(int numRows, int numCols);
	
	    /**
	     * Create a new OLE object shape with the given pictureData as preview image
	     *
	     * @param pictureData the preview image
	     */
	    //ObjectShape<?,?> createOleShape(PictureData pictureData);
	}
}