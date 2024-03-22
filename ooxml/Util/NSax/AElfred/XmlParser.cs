// AElfred XML Parser. This version of the AElfred parser is
// derived from the original Microstar distribution, with additional
// bug fixes by Michael Kay, and selected enhancements and further
// bug fixes from the version produced by David Brownell.
// ported to .Net by Rasmus John Pedersen

namespace NSAX.AElfred
{
  using System;
  using System.Collections;
  using System.Globalization;
  using System.IO;
  using System.Linq;
  using System.Net;
  using System.Text;

  using NSAX;

  /// <summary>
  ///   Parse XML documents and return parse events through call-backs.
  ///   Use the <c>SAXDriver</c> class as your entry point, as the
  ///   internal parser interfaces are subject to change.
  /// </summary>
  /// <seealso cref="SAXDriver" />
  internal sealed class XmlParser
  {
    // parse from buffer, avoiding slow per-character readCh()
    private const bool USE_CHEATS = true;

    // don't waste too much space in hashtables 
    private const int DEFAULT_ATTR_COUNT = 23;


    //////////////////////////////////////////////////////////////////////
    // Constructors.
    ////////////////////////////////////////////////////////////////////////


    /**
     * Construct a new parser with no associated handler.
     * @see #setHandler
     * @see #parse
     */
    // package private
    public XmlParser()
    {
      cleanupVariables();
    }


    /**
     * Set the handler that will receive parsing events.
     * @param handler The handler to receive callback events.
     * @see #parse
     */
    // package private
    public void setHandler(SAXDriver handler)
    {
      this._handler = handler;
    }


    /**
     * Parse an XML document from the character stream, byte stream, or URI
     * that you provide (in that order of preference).  Any URI that you
     * supply will become the base URI for resolving relative URI, and may
     * be used to acquire a reader or byte stream.
     *
     * <p>You may parse more than one document, but that must be done
     * sequentially.  Only one thread at a time may use this parser.
     *
     * @param systemId The URI of the document; should never be null,
     *	but may be so iff a reader <em>or</em> a stream is provided.
     * @param publicId The public identifier of the document, or null.
     * @param reader A character stream; must be null if stream isn't.
     * @param stream A byte input stream; must be null if reader isn't.
     * @param encoding The suggested encoding, or null if unknown.
     * @exception java.lang.Exception Basically SAXException or IOException
     */
    // package private 
    public void DoParse(string systemId, string publicId, TextReader reader, Stream stream, Encoding encoding)
    {
      if (this._handler == null) throw new InvalidOperationException("no callback handler");

      this._basePublicId = publicId;
      this._baseURI = systemId;
      this._baseReader = reader;
      this._baseInputStream = stream;

      initializeVariables();

      // predeclare the built-in entities here (replacement texts)
      // we don't need to intern(), since we're guaranteed literals
      // are always (globally) interned.
      setInternalEntity("amp", "&#38;");
      setInternalEntity("lt", "&#60;");
      setInternalEntity("gt", "&#62;");
      setInternalEntity("apos", "&#39;");
      setInternalEntity("quot", "&#34;");

      this._handler.StartDocument();

      pushURL("[document]", this._basePublicId, this._baseURI, this._baseReader, this._baseInputStream, encoding, false);

      try
      {
        parseDocument();
        this._handler.EndDocument();
      }
      finally
      {
        if (this._baseReader != null)
          try
          {
            this._baseReader.Close();
          }
          catch (IOException e)
          {
            /* ignore */
          }
        if (this._baseInputStream != null)
          try
          {
            this._baseInputStream.Close();
          }
          catch (IOException e)
          {
            /* ignore */
          }
        if (this._stream != null)
          try
          {
            this._stream.Close();
          }
          catch (IOException e)
          {
            /* ignore */
          }
        if (reader != null)
          try
          {
            reader.Close();
          }
          catch (IOException e)
          {
            /* ignore */
          }
        cleanupVariables();
      }
    }


    ////////////////////////////////////////////////////////////////////////
    // Constants.
    ////////////////////////////////////////////////////////////////////////

    //
    // Constants for element content type.
    //

    /**
     * Constant: an element has not been declared.
     * @see #getElementContentType
     */

    public const int CONTENT_UNDECLARED = 0;

    /**
     * Constant: the element has a content model of ANY.
     * @see #getElementContentType
     */

    public const int CONTENT_ANY = 1;

    /**
     * Constant: the element has declared content of EMPTY.
     * @see #getElementContentType
     */

    public const int CONTENT_EMPTY = 2;

    /**
     * Constant: the element has mixed content.
     * @see #getElementContentType
     */

    public const int CONTENT_MIXED = 3;

    /**
     * Constant: the element has element content.
     * @see #getElementContentType
     */

    public const int CONTENT_ELEMENTS = 4;


    //
    // Constants for the entity type.
    //

    /**
     * Constant: the entity has not been declared.
     * @see #getEntityType
     */

    public const int ENTITY_UNDECLARED = 0;

    /**
     * Constant: the entity is internal.
     * @see #getEntityType
     */

    public const int ENTITY_INTERNAL = 1;

    /**
     * Constant: the entity is external, non-parseable data.
     * @see #getEntityType
     */

    public const int ENTITY_NDATA = 2;

    /**
     * Constant: the entity is external XML data.
     * @see #getEntityType
     */

    public const int ENTITY_TEXT = 3;


    //
    // Constants for attribute type.
    //

    /**
     * Constant: the attribute has not been declared for this element type.
     * @see #getAttributeType
     */

    public const int ATTRIBUTE_UNDECLARED = 0;

    /**
     * Constant: the attribute value is a string value.
     * @see #getAttributeType
     */

    public const int ATTRIBUTE_CDATA = 1;

    /**
     * Constant: the attribute value is a unique identifier.
     * @see #getAttributeType
     */

    public const int ATTRIBUTE_ID = 2;

    /**
     * Constant: the attribute value is a reference to a unique identifier.
     * @see #getAttributeType
     */

    public const int ATTRIBUTE_IDREF = 3;

    /**
     * Constant: the attribute value is a list of ID references.
     * @see #getAttributeType
     */

    public const int ATTRIBUTE_IDREFS = 4;

    /**
     * Constant: the attribute value is the name of an entity.
     * @see #getAttributeType
     */

    public const int ATTRIBUTE_ENTITY = 5;

    /**
     * Constant: the attribute value is a list of entity names.
     * @see #getAttributeType
     */

    public const int ATTRIBUTE_ENTITIES = 6;

    /**
     * Constant: the attribute value is a name token.
     * @see #getAttributeType
     */

    public const int ATTRIBUTE_NMTOKEN = 7;

    /**
     * Constant: the attribute value is a list of name tokens.
     * @see #getAttributeType
     */

    public const int ATTRIBUTE_NMTOKENS = 8;

    /**
     * Constant: the attribute value is a token from an enumeration.
     * @see #getAttributeType
     */

    public const int ATTRIBUTE_ENUMERATED = 9;

    /**
     * Constant: the attribute is the name of a notation.
     * @see #getAttributeType
     */

    public const int ATTRIBUTE_NOTATION = 10;


    //
    // When the class is loaded, populate the hash table of
    // attribute types.
    //

    /**
     * Hash table of attribute types.
     */

    private static Hashtable attributeTypeHash;

    static XmlParser()
    {
      attributeTypeHash = new Hashtable(13);
      attributeTypeHash.Add("CDATA", ATTRIBUTE_CDATA);
      attributeTypeHash.Add("ID", ATTRIBUTE_ID);
      attributeTypeHash.Add("IDREF", ATTRIBUTE_IDREF);
      attributeTypeHash.Add("IDREFS", ATTRIBUTE_IDREFS);
      attributeTypeHash.Add("ENTITY", ATTRIBUTE_ENTITY);
      attributeTypeHash.Add("ENTITIES", ATTRIBUTE_ENTITIES);
      attributeTypeHash.Add("NMTOKEN", ATTRIBUTE_NMTOKEN);
      attributeTypeHash.Add("NMTOKENS", ATTRIBUTE_NMTOKENS);
      attributeTypeHash.Add("NOTATION", ATTRIBUTE_NOTATION);
    }


    //
    // Constants for supported encodings.  "external" is just a flag.
    //
    private const int ENCODING_EXTERNAL = 0;

    private const int ENCODING_UTF_8 = 1;

    private const int ENCODING_ISO_8859_1 = 2;

    private const int ENCODING_UCS_2_12 = 3;

    private const int ENCODING_UCS_2_21 = 4;

    private const int ENCODING_UCS_4_1234 = 5;

    private const int ENCODING_UCS_4_4321 = 6;

    private const int ENCODING_UCS_4_2143 = 7;

    private const int ENCODING_UCS_4_3412 = 8;

    private const int ENCODING_ASCII = 9;


    //
    // Constants for attribute default value.
    //

    /**
     * Constant: the attribute is not declared.
     * @see #getAttributeDefaultValueType
     */

    public const int ATTRIBUTE_DEFAULT_UNDECLARED = 30;

    /**
     * Constant: the attribute has a literal default value specified.
     * @see #getAttributeDefaultValueType
     * @see #getAttributeDefaultValue
     */

    public const int ATTRIBUTE_DEFAULT_SPECIFIED = 31;

    /**
     * Constant: the attribute was declared #IMPLIED.
     * @see #getAttributeDefaultValueType
     */

    public const int ATTRIBUTE_DEFAULT_IMPLIED = 32;

    /**
     * Constant: the attribute was declared #REQUIRED.
     * @see #getAttributeDefaultValueType
     */

    public const int ATTRIBUTE_DEFAULT_REQUIRED = 33;

    /**
     * Constant: the attribute was declared #FIXED.
     * @see #getAttributeDefaultValueType
     * @see #getAttributeDefaultValue
     */

    public const int ATTRIBUTE_DEFAULT_FIXED = 34;


    //
    // Constants for input.
    //
    private const int INPUT_NONE = 0;

    private const int INPUT_INTERNAL = 1;

    private const int INPUT_STREAM = 3;

    private const int INPUT_BUFFER = 4;

    private const int INPUT_READER = 5;


    //
    // Flags for reading literals.
    //
    // expand general entity refs (attribute values in dtd and content)
    private const int LIT_ENTITY_REF = 2;

    // normalize this value (space chars) (attributes, public ids)
    private const int LIT_NORMALIZE = 4;

    // literal is an attribute value 
    private const int LIT_ATTRIBUTE = 8;

    // don't expand parameter entities
    private const int LIT_DISABLE_PE = 16;

    // don't expand [or parse] character refs
    private const int LIT_DISABLE_CREF = 32;

    // don't parse general entity refs
    private const int LIT_DISABLE_EREF = 64;

    // don't expand general entities, but make sure we _could_
    private const int LIT_ENTITY_CHECK = 128;

    // literal is a public ID value 
    private const int LIT_PUBID = 256;

    //
    // Flags affecting PE handling in DTDs (if expandPE is true).
    // PEs expand with space padding, except inside literals.
    //
    private const int CONTEXT_NORMAL = 0;

    private const int CONTEXT_LITERAL = 1;


    //////////////////////////////////////////////////////////////////////
    // Error reporting.
    //////////////////////////////////////////////////////////////////////


    /**
     * Report an error.
     * @param message The error message.
     * @param textFound The text that caused the error (or null).
     * @see SAXDriver#error
     * @see #line
     */

    private void error(string message, string textFound, string textExpected)
    {
      if (textFound != null)
      {
        message = message + " (found \"" + textFound + "\")";
      }
      if (textExpected != null)
      {
        message = message + " (expected \"" + textExpected + "\")";
      }
      string uri = null;

      if (this._externalEntity != null)
      {
        uri = this._externalEntity.RequestUri.ToString();
      }
      this._handler.Error(message, uri, this._line, this._column);

      // "can't happen"
      throw new SAXException(message);
    }


    /**
     * Report a serious error.
     * @param message The error message.
     * @param textFound The text that caused the error (or null).
     */

    private void error(string message, char textFound, string textExpected)
    {
      error(message, textFound.ToString(CultureInfo.CurrentCulture), textExpected);
    }

    /** Report typical case fatal errors. */

    private void error(string message)
    {
      error(message, null, null);
    }


    //////////////////////////////////////////////////////////////////////
    // Major syntactic productions.
    //////////////////////////////////////////////////////////////////////


    /**
     * Parse an XML document.
     * <pre>
     * [1] document ::= prolog element Misc*
     * </pre>
     * <p>This is the top-level parsing function for a single XML
     * document.  As a minimum, a well-formed document must have
     * a document element, and a valid document must have a prolog
     * (one with doctype) as well.
     */

    private void parseDocument()
    {
      try
      {
        // added by MHK
        parseProlog();
        require('<', "document prolog");
        parseElement();
      }
      catch (EndOfStreamException ee)
      {
        // added by MHK
        error("premature end of file", "[EOF]", null);
      }

      try
      {
        parseMisc(); //skip all white, PIs, and comments
        char c = readCh(); //if this doesn't throw an exception...
        error("unexpected characters after document end", c, null);
      }
      catch (EndOfStreamException e)
      {
        return;
      }
    }


    /**
     * Skip a comment.
     * <pre>
     * [15] Comment ::= '&lt;!--' ((Char - '-') | ('-' (Char - '-')))* "-->"
     * </pre>
     * <p> (The <code>&lt;!--</code> has already been read.)
     */

    private void parseComment()
    {
      char c;
      bool saved = this._expandPE;

      this._expandPE = false;
      parseUntil("--");
      require('>', "-- in comment");
      this._expandPE = saved;
      this._handler.Comment(this._dataBuffer, 0, this._dataBufferPos);
      this._dataBufferPos = 0;
    }


    /**
     * Parse a processing instruction and do a call-back.
     * <pre>
     * [16] PI ::= '&lt;?' PITarget
     *		(S (Char* - (Char* '?&gt;' Char*)))?
     *		'?&gt;'
     * [17] PITarget ::= Name - ( ('X'|'x') ('M'|m') ('L'|l') )
     * </pre>
     * <p> (The <code>&lt;?</code> has already been read.)
     */

    private void parsePI()
    {
      string name;
      bool saved = this._expandPE;

      this._expandPE = false;
      name = readNmtoken(true);
      if ("xml".Equals(name, StringComparison.InvariantCultureIgnoreCase)) error("Illegal processing instruction target", name, null);
      if (!tryRead("?>"))
      {
        requireWhitespace();
        parseUntil("?>");
      }
      this._expandPE = saved;
      this._handler.ProcessingInstruction(name, dataBufferToString());
    }


    /**
     * Parse a CDATA section.
     * <pre>
     * [18] CDSect ::= CDStart CData CDEnd
     * [19] CDStart ::= '&lt;![CDATA['
     * [20] CData ::= (Char* - (Char* ']]&gt;' Char*))
     * [21] CDEnd ::= ']]&gt;'
     * </pre>
     * <p> (The '&lt;![CDATA[' has already been read.)
     */

    private void parseCDSect()
    {
      parseUntil("]]>");
      dataBufferFlush();
    }


    /**
     * Parse the prolog of an XML document.
     * <pre>
     * [22] prolog ::= XMLDecl? Misc* (Doctypedecl Misc*)?
     * </pre>
     * <p>There are a couple of tricks here.  First, it is necessary to
     * declare the XML default attributes after the DTD (if present)
     * has been read. [??]  Second, it is not possible to expand general
     * references in attribute value literals until after the entire
     * DTD (if present) has been parsed.
     * <p>We do not look for the XML declaration here, because it was
     * handled by pushURL ().
     * @see pushURL
     */

    private void parseProlog()
    {
      parseMisc();

      if (tryRead("<!DOCTYPE"))
      {
        parseDoctypedecl();
        parseMisc();
      }
    }


    /**
     * Parse the XML declaration.
     * <pre>
     * [23] XMLDecl ::= '&lt;?xml' VersionInfo EncodingDecl? SDDecl? S? '?&gt;'
     * [24] VersionInfo ::= S 'version' Eq
     *		("'" VersionNum "'" | '"' VersionNum '"' )
     * [26] VersionNum ::= ([a-zA-Z0-9_.:] | '-')*
     * [32] SDDecl ::= S 'standalone' Eq
     *		( "'"" ('yes' | 'no') "'"" | '"' ("yes" | "no") '"' )
     * [80] EncodingDecl ::= S 'encoding' Eq
     *		( "'" EncName "'" | "'" EncName "'" )
     * [81] EncName ::= [A-Za-z] ([A-Za-z0-9._] | '-')*
     * </pre>
     * <p> (The <code>&lt;?xml</code> and whitespace have already been read.)
     * @return the encoding in the declaration, uppercased; or null
     * @see #parseTextDecl
     * @see #setupDecoding
     */

    private string parseXMLDecl(bool ignoreEncoding)
    {
      string version;
      string encodingName = null;
      string standalone = null;
      int flags = LIT_DISABLE_CREF | LIT_DISABLE_PE | LIT_DISABLE_EREF;

      // Read the version.
      require("version", "XML declaration");
      parseEq();
      version = readLiteral(flags);
      if (!version.Equals("1.0"))
      {
        error("unsupported XML version", version, "1.0");
      }

      // Try reading an encoding declaration.
      bool white = tryWhitespace();
      if (tryRead("encoding"))
      {
        if (!white) error("whitespace required before 'encoding='");
        parseEq();
        encodingName = readLiteral(flags);
        if (!ignoreEncoding) setupDecoding(encodingName);
      }

      // Try reading a standalone declaration
      if (encodingName != null) white = tryWhitespace();
      if (tryRead("standalone"))
      {
        if (!white) error("whitespace required before 'standalone='");
        parseEq();
        standalone = readLiteral(flags);
        if (! ("yes".Equals(standalone) || "no".Equals(standalone))) error("standalone flag must be 'yes' or 'no'");
      }

      skipWhitespace();
      require("?>", "XML declaration");

      return encodingName;
    }


    /**
     * Parse a text declaration.
     * <pre>
     * [79] TextDecl ::= '&lt;?xml' VersionInfo? EncodingDecl S? '?&gt;'
     * [80] EncodingDecl ::= S 'encoding' Eq
     *		( '"' EncName '"' | "'" EncName "'" )
     * [81] EncName ::= [A-Za-z] ([A-Za-z0-9._] | '-')*
     * </pre>
     * <p> (The <code>&lt;?xml</code>' and whitespace have already been read.)
     * @return the encoding in the declaration, uppercased; or null
     * @see #parseXMLDecl
     * @see #setupDecoding
     */

    private string parseTextDecl(bool ignoreEncoding)
    {
      string encodingName = null;
      int flags = LIT_DISABLE_CREF | LIT_DISABLE_PE | LIT_DISABLE_EREF;

      // Read an optional version.
      if (tryRead("version"))
      {
        string version;
        parseEq();
        version = readLiteral(flags);
        if (!version.Equals("1.0"))
        {
          error("unsupported XML version", version, "1.0");
        }
        requireWhitespace();
      }


      // Read the encoding.
      require("encoding", "XML text declaration");
      parseEq();
      encodingName = readLiteral(flags);
      if (!ignoreEncoding) setupDecoding(encodingName);

      skipWhitespace();
      require("?>", "XML text declaration");

      return encodingName;
    }


    /**
     * Sets up internal state so that we can decode an entity using the
     * specified encoding.  This is used when we start to read an entity
     * and we have been given knowledge of its encoding before we start to
     * read any data (e.g. from a SAX input source or from a MIME type).
     *
     * <p> It is also used after autodetection, at which point only very
     * limited adjustments to the encoding may be used (switching between
     * related builtin decoders).
     *
     * @param encodingName The name of the encoding specified by the user.
     * @exception IOException if the encoding isn't supported either
     *	internally to this parser, or by the hosting JVM.
     * @see #parseXMLDecl
     * @see #parseTextDecl
     */

    private void setupDecoding(Encoding encoding) {
      setupDecoding(encoding.WebName);
    }

    private void setupDecoding(string encodingName)
    {
      encodingName = encodingName.ToUpper();

      // ENCODING_EXTERNAL indicates an encoding that wasn't
      // autodetected ... we can use builtin decoders, or
      // ones from the JVM (InputStreamReader).

      // Otherwise we can only tweak what was autodetected, and
      // only for single byte (ASCII derived) builtin encodings.

      // ASCII-derived encodings
      if (this._encoding == ENCODING_UTF_8 || this._encoding == ENCODING_EXTERNAL)
      {
        if (encodingName.Equals("ISO-8859-1") || encodingName.Equals("8859_1") || encodingName.Equals("ISO8859_1"))
        {
          this._encoding = ENCODING_ISO_8859_1;
          return;
        }
        else if (encodingName.Equals("US-ASCII") || encodingName.Equals("ASCII"))
        {
          this._encoding = ENCODING_ASCII;
          return;
        }
        else if (encodingName.Equals("UTF-8") || encodingName.Equals("UTF8"))
        {
          this._encoding = ENCODING_UTF_8;
          return;
        }
        else if (this._encoding != ENCODING_EXTERNAL)
        {
          // used to start with a new reader ...
          throw new EncodingException(encodingName);
        }
        // else fallthrough ...
        // it's ASCII-ish and something other than a builtin
      }

      // Unicode and such
      if (this._encoding == ENCODING_UCS_2_12 || this._encoding == ENCODING_UCS_2_21)
      {
        if (
          !(encodingName.Equals("ISO-10646-UCS-2") || encodingName.Equals("UTF-16") || encodingName.Equals("UTF-16BE")
            || encodingName.Equals("UTF-16LE"))) error("unsupported Unicode encoding", encodingName, "UTF-16");
        return;
      }

      // four byte encodings
      if (this._encoding == ENCODING_UCS_4_1234 || this._encoding == ENCODING_UCS_4_4321 || this._encoding == ENCODING_UCS_4_2143
          || this._encoding == ENCODING_UCS_4_3412)
      {
        if (!encodingName.Equals("ISO-10646-UCS-4")) error("unsupported 32-bit encoding", encodingName, "ISO-10646-UCS-4");
        return;
      }

      // assert encoding == ENCODING_EXTERNAL
      // if (encoding != ENCODING_EXTERNAL)
      //     throw new RuntimeException ("encoding = " + encoding);

      if (encodingName.Equals("UTF-16BE"))
      {
        this._encoding = ENCODING_UCS_2_12;
        return;
      }
      if (encodingName.Equals("UTF-16LE"))
      {
        this._encoding = ENCODING_UCS_2_21;
        return;
      }

      // We couldn't use the builtin decoders at all.  But we can try to
      // create a reader, since we haven't messed up buffering.  Tweak
      // the encoding name if necessary.

      if (encodingName.Equals("UTF-16") || encodingName.Equals("ISO-10646-UCS-2")) encodingName = "Unicode";
      // Ignoring all the EBCDIC aliases here

      
      this._reader = new StreamReader(this._stream, Encoding.GetEncoding(encodingName));
      this._sourceType = INPUT_READER;
    }


    /**
     * Parse miscellaneous markup outside the document element and DOCTYPE
     * declaration.
     * <pre>
     * [27] Misc ::= Comment | PI | S
     * </pre>
     */

    private void parseMisc()
    {
      while (true)
      {
        skipWhitespace();
        if (tryRead("<?"))
        {
          parsePI();
        }
        else if (tryRead("<!--"))
        {
          parseComment();
        }
        else
        {
          return;
        }
      }
    }


    /**
     * Parse a document type declaration.
     * <pre>
     * [28] doctypedecl ::= '&lt;!DOCTYPE' S Name (S ExternalID)? S?
     *		('[' (markupdecl | PEReference | S)* ']' S?)? '&gt;'
     * </pre>
     * <p> (The <code>&lt;!DOCTYPE</code> has already been read.)
     */

    private void parseDoctypedecl()
    {
      string doctypeName;
      string[] ids;

      // Read the document type name.
      requireWhitespace();
      doctypeName = readNmtoken(true);

      // Read the External subset's IDs
      skipWhitespace();
      ids = readExternalIds(false);

      // report (a) declaration of name, (b) lexical info (ids)
      this._handler.DoctypeDecl(doctypeName, ids[0], ids[1]);

      // Internal subset is parsed first, if present
      skipWhitespace();
      if (tryRead('['))
      {

        // loop until the subset ends
        while (true)
        {
          this._expandPE = true;
          skipWhitespace();
          this._expandPE = false;
          if (tryRead(']'))
          {
            break; // end of subset
          }
          else
          {
            // WFC, PEs in internal subset (only between decls)
            this._peIsError = this._expandPE = true;
            parseMarkupdecl();
            this._peIsError = this._expandPE = false;
          }
        }
      }

      // Read the external subset, if any
      if (ids[1] != null)
      {
        pushURL("[external subset]", ids[0], ids[1], null, null, null, false);

        // Loop until we end up back at '>'
        while (true)
        {
          this._expandPE = true;
          skipWhitespace();
          this._expandPE = false;
          if (tryRead('>'))
          {
            break;
          }
          else
          {
            this._expandPE = true;
            parseMarkupdecl();
            this._expandPE = false;
          }
        }
      }
      else
      {
        // No external subset.
        skipWhitespace();
        require('>', "internal DTD subset");
      }

      // done dtd
      this._handler.EndDoctype();
      this._expandPE = false;
    }


    /**
     * Parse a markup declaration in the internal or external DTD subset.
     * <pre>
     * [29] markupdecl ::= elementdecl | Attlistdecl | EntityDecl
     *		| NotationDecl | PI | Comment
     * [30] extSubsetDecl ::= (markupdecl | conditionalSect
     *		| PEReference | S) *
     * </pre>
     * <p> Reading toplevel PE references is handled as a lexical issue
     * by the caller, as is whitespace.
     */

    private void parseMarkupdecl()
    {
      if (tryRead("<!ELEMENT"))
      {
        parseElementdecl();
      }
      else if (tryRead("<!ATTLIST"))
      {
        parseAttlistDecl();
      }
      else if (tryRead("<!ENTITY"))
      {
        parseEntityDecl();
      }
      else if (tryRead("<!NOTATION"))
      {
        parseNotationDecl();
      }
      else if (tryRead("<?"))
      {
        parsePI();
      }
      else if (tryRead("<!--"))
      {
        parseComment();
      }
      else if (tryRead("<!["))
      {
        if (this._inputStack.Count > 0) parseConditionalSect();
        else error("conditional sections illegal in internal subset");
      }
      else
      {
        error("expected markup declaration");
      }
    }


    /**
     * Parse an element, with its tags.
     * <pre>
     * [39] element ::= EmptyElementTag | STag content ETag
     * [40] STag ::= '&lt;' Name (S Attribute)* S? '&gt;'
     * [44] EmptyElementTag ::= '&lt;' Name (S Attribute)* S? '/&gt;'
     * </pre>
     * <p> (The '&lt;' has already been read.)
     * <p>NOTE: this method actually chains onto parseContent (), if necessary,
     * and parseContent () will take care of calling parseETag ().
     */

    private void parseElement()
    {
      string gi;
      char c;
      int oldElementContent = this._currentElementContent;
      string oldElement = this._currentElement;
      object[] element;

      // This is the (global) counter for the
      // array of specified attributes.
      this._tagAttributePos = 0;

      // Read the element type name.
      gi = readNmtoken(true);

      // Determine the current content type.
      this._currentElement = gi;
      element = (object[])this._elementInfo[gi];
      this._currentElementContent = getContentType(element, CONTENT_ANY);

      // Read the attributes, if any.
      // After this loop, "c" is the closing delimiter.
      bool white = tryWhitespace();
      c = readCh();
      while (c != '/' && c != '>')
      {
        unread(c);
        if (!white) error("need whitespace between attributes");
        parseAttribute(gi);
        white = tryWhitespace();
        c = readCh();
      }

      // Supply any defaulted attributes.
      IEnumerable atts = declaredAttributes(element);
      if (atts != null)
      {

        foreach (string aname in atts.Cast<string>())
        {
          bool cont = false;
          // See if it was specified.
          for (int i = 0; i < this._tagAttributePos; i++)
          {
            if (this._tagAttributes[i] == aname)
            {
              cont = true;
            }
          }
          if (cont)
          {
            continue;
          }
          // I guess not...
          string defaultVal = this.getAttributeExpandedValue(gi, aname);
          if (defaultVal != null)
          {
            this._handler.Attribute(aname, defaultVal, false);
          }
        }
      }

      // Figure out if this is a start tag
      // or an empty element, and dispatch an
      // event accordingly.
      switch (c)
      {
        case '>':
          this._handler.StartElement(gi);
          parseContent();
          break;
        case '/':
          require('>', "empty element tag");
          this._handler.StartElement(gi);
          this._handler.EndElement(gi);
          break;
      }

      // Restore the previous state.
      this._currentElement = oldElement;
      this._currentElementContent = oldElementContent;
    }


    /**
     * Parse an attribute assignment.
     * <pre>
     * [41] Attribute ::= Name Eq AttValue
     * </pre>
     * @param name The name of the attribute's element.
     * @see SAXDriver#attribute
     */

    private void parseAttribute(string name)
    {
      string aname;
      int type;
      string value;
      int flags = LIT_ATTRIBUTE | LIT_ENTITY_REF;

      // Read the attribute name.
      aname = readNmtoken(true);
      type = getAttributeType(name, aname);

      // Parse '='
      parseEq();

      // Read the value, normalizing whitespace
      // unless it is CDATA.
      if (type == ATTRIBUTE_CDATA || type == ATTRIBUTE_UNDECLARED)
      {
        value = readLiteral(flags);
      }
      else
      {
        value = readLiteral(flags | LIT_NORMALIZE);
      }

      // WFC: no duplicate attributes
      for (int i = 0; i < this._tagAttributePos; i++) if (aname.Equals(this._tagAttributes[i])) error("duplicate attribute", aname, null);

      // Above check is almost redundant; the SAXDriver performs a more
      // rigorous check that the expanded-names of the attributes are distinct. However,
      // the check is needed here to spot duplicate xmlns:xx attributes. - MHK

      // Inform the handler about the
      // attribute.
      this._handler.Attribute(aname, value, true);
      this._dataBufferPos = 0;

      // Note that the attribute has been
      // specified.
      if (this._tagAttributePos == this._tagAttributes.Length)
      {
        string[] newAttrib = new string[this._tagAttributes.Length * 2];
        Array.Copy(this._tagAttributes, 0, newAttrib, 0, this._tagAttributePos);
        this._tagAttributes = newAttrib;
      }
      this._tagAttributes[this._tagAttributePos++] = aname;
    }


    /**
     * Parse an Equals sign surrounded by optional whitespace.
     * <pre>
     * [25] Eq ::= S? '=' S?
     * </pre>
     */

    private void parseEq()
    {
      skipWhitespace();
      require('=', "attribute name");
      skipWhitespace();
    }


    /**
     * Parse an end tag.
     * <pre>
     * [42] ETag ::= '</' Name S? '>'
     * </pre>
     * <p>NOTE: parseContent () chains to here, we already read the
     * "&lt;/".
     */

    private void parseETag()
    {
      require(this._currentElement, "element end tag");
      skipWhitespace();
      require('>', "name in end tag");
      this._handler.EndElement(this._currentElement);
      // not re-reporting any SAXException re bogus end tags,
      // even though that diagnostic might be clearer ...
    }


    /**
     * Parse the content of an element.
     * <pre>
     * [43] content ::= (element | CharData | Reference
     *		| CDSect | PI | Comment)*
     * [67] Reference ::= EntityRef | CharRef
     * </pre>
     * <p> NOTE: consumes ETtag.
     */

    private void parseContent()
    {
      char c;
      while (true)
      {
        //switch (currentElementContent) {
        //    case CONTENT_ANY:
        //    case CONTENT_MIXED:
        //    case CONTENT_UNDECLARED:    // this line added by MHK 24 May 2000
        //    case CONTENT_EMPTY:         // this line added by MHK 8 Sept 2000
        //        parseCharData ();
        //        break;
        //    case CONTENT_ELEMENTS:
        //        //parseWhitespace ();   // removed MHK 27 May 2001. The problem is that
        //                                // with element content, the text should be whitespace
        //                                // but if the document is invalid it might not be.
        //                                // Replaced with....
        //        parseCharData();        // This processes any char data, but still reports
        //                                // it as ignorable white space if within element content.
        //        break;
        //}

        parseCharData(); // parse it the same way regardless of content type
        // because it might not be valid anyway

        // Handle delimiters
        c = readCh();
        switch (c)
        {
          case '&': // Found "&"

            c = readCh();
            if (c == '#')
            {
              parseCharRef();
            }
            else
            {
              unread(c);
              parseEntityRef(true);
            }
            break;

          case '<': // Found "<"
            dataBufferFlush();
            c = readCh();
            switch (c)
            {
              case '!': // Found "<!"
                c = readCh();
                switch (c)
                {
                  case '-': // Found "<!-"
                    require('-', "start of comment");
                    parseComment();
                    break;
                  case '[': // Found "<!["
                    require("CDATA[", "CDATA section");
                    this._handler.StartCDATA();
                    this._inCDATA = true;
                    parseCDSect();
                    this._inCDATA = false;
                    this._handler.EndCDATA();
                    break;
                  default:
                    error("expected comment or CDATA section", c, null);
                    break;
                }
                break;

              case '?': // Found "<?"
                parsePI();
                break;

              case '/': // Found "</"
                parseETag();
                return;

              default: // Found "<" followed by something else
                unread(c);
                parseElement();
                break;
            }
            break;
        }
      }
    }


    /**
     * Parse an element type declaration.
     * <pre>
     * [45] elementdecl ::= '&lt;!ELEMENT' S Name S contentspec S? '&gt;'
     * </pre>
     * <p> NOTE: the '&lt;!ELEMENT' has already been read.
     */

    private void parseElementdecl()
    {
      string name;

      requireWhitespace();
      // Read the element type name.
      name = readNmtoken(true);

      requireWhitespace();
      // Read the content model.
      parseContentspec(name);

      skipWhitespace();
      require('>', "element declaration");
    }


    /**
     * Content specification.
     * <pre>
     * [46] contentspec ::= 'EMPTY' | 'ANY' | Mixed | elements
     * </pre>
     */

    private void parseContentspec(string name)
    {
      if (tryRead("EMPTY"))
      {
        setElement(name, CONTENT_EMPTY, null, null);
        return;
      }
      else if (tryRead("ANY"))
      {
        setElement(name, CONTENT_ANY, null, null);
        return;
      }
      else
      {
        require('(', "element name");
        dataBufferAppend('(');
        skipWhitespace();
        if (tryRead("#PCDATA"))
        {
          dataBufferAppend("#PCDATA");
          parseMixed();
          setElement(name, CONTENT_MIXED, dataBufferToString(), null);
        }
        else
        {
          parseElements();
          setElement(name, CONTENT_ELEMENTS, dataBufferToString(), null);
        }
      }
    }


    /**
     * Parse an element-content model.
     * <pre>
     * [47] elements ::= (choice | seq) ('?' | '*' | '+')?
     * [49] choice ::= '(' S? cp (S? '|' S? cp)+ S? ')'
     * [50] seq ::= '(' S? cp (S? ',' S? cp)* S? ')'
     * </pre>
     *
     * <p> NOTE: the opening '(' and S have already been read.
     */

    private void parseElements()
    {
      char c;
      char sep;

      // Parse the first content particle
      skipWhitespace();
      parseCp();

      // Check for end or for a separator.
      skipWhitespace();
      c = readCh();
      switch (c)
      {
        case ')':
          dataBufferAppend(')');
          c = readCh();
          switch (c)
          {
            case '*':
            case '+':
            case '?':
              dataBufferAppend(c);
              break;
            default:
              unread(c);
              break;
          }
          return;
        case ',': // Register the separator.
        case '|':
          sep = c;
          dataBufferAppend(c);
          break;
        default:
          error("bad separator in content model", c, null);
          return;
      }

      // Parse the rest of the content model.
      while (true)
      {
        skipWhitespace();
        parseCp();
        skipWhitespace();
        c = readCh();
        if (c == ')')
        {
          dataBufferAppend(')');
          break;
        }
        else if (c != sep)
        {
          error("bad separator in content model", c, null);
          return;
        }
        else
        {
          dataBufferAppend(c);
        }
      }

      // Check for the occurrence indicator.
      c = readCh();
      switch (c)
      {
        case '?':
        case '*':
        case '+':
          dataBufferAppend(c);
          return;
        default:
          unread(c);
          return;
      }
    }


    /**
     * Parse a content particle.
     * <pre>
     * [48] cp ::= (Name | choice | seq) ('?' | '*' | '+')?
     * </pre>
     */

    private void parseCp()
    {
      if (tryRead('('))
      {
        dataBufferAppend('(');
        parseElements();
      }
      else
      {
        dataBufferAppend(readNmtoken(true));
        char c = readCh();
        switch (c)
        {
          case '?':
          case '*':
          case '+':
            dataBufferAppend(c);
            break;
          default:
            unread(c);
            break;
        }
      }
    }


    /**
     * Parse mixed content.
     * <pre>
     * [51] Mixed ::= '(' S? ( '#PCDATA' (S? '|' S? Name)*) S? ')*'
     *	      | '(' S? ('#PCDATA') S? ')'
     * </pre>
     */

    private void parseMixed()
    {

      // Check for PCDATA alone.
      skipWhitespace();
      if (tryRead(')'))
      {
        dataBufferAppend(")*");
        tryRead('*');
        return;
      }

      // Parse mixed content.
      skipWhitespace();
      while (!tryRead(")*"))
      {
        require('|', "alternative");
        dataBufferAppend('|');
        skipWhitespace();
        dataBufferAppend(readNmtoken(true));
        skipWhitespace();
      }
      dataBufferAppend(")*");
    }


    /**
     * Parse an attribute list declaration.
     * <pre>
     * [52] AttlistDecl ::= '&lt;!ATTLIST' S Name AttDef* S? '&gt;'
     * </pre>
     * <p>NOTE: the '&lt;!ATTLIST' has already been read.
     */

    private void parseAttlistDecl()
    {
      string elementName;

      requireWhitespace();
      elementName = readNmtoken(true);
      bool white = tryWhitespace();
      while (!tryRead('>'))
      {
        if (!white) error("whitespace required before attribute definition");
        parseAttDef(elementName);
        white = tryWhitespace();
      }
    }


    /**
     * Parse a single attribute definition.
     * <pre>
     * [53] AttDef ::= S Name S AttType S DefaultDecl
     * </pre>
     */

    private void parseAttDef(string elementName)
    {
      string name;
      int type;
      string @enum = null;

      // Read the attribute name.
      name = readNmtoken(true);

      // Read the attribute type.
      requireWhitespace();
      type = readAttType();

      // Get the string of enumerated values
      // if necessary.
      if (type == ATTRIBUTE_ENUMERATED || type == ATTRIBUTE_NOTATION)
      {
        @enum = dataBufferToString();
      }

      // Read the default value.
      requireWhitespace();
      parseDefault(elementName, name, type, @enum);
    }


    /**
     * Parse the attribute type.
     * <pre>
     * [54] AttType ::= StringType | TokenizedType | EnumeratedType
     * [55] StringType ::= 'CDATA'
     * [56] TokenizedType ::= 'ID' | 'IDREF' | 'IDREFS' | 'ENTITY'
     *		| 'ENTITIES' | 'NMTOKEN' | 'NMTOKENS'
     * [57] EnumeratedType ::= NotationType | Enumeration
     * </pre>
     */

    private int readAttType()
    {
      if (tryRead('('))
      {
        parseEnumeration(false);
        return ATTRIBUTE_ENUMERATED;
      }
      else
      {
        string typeString = readNmtoken(true);
        if (typeString.Equals("NOTATION"))
        {
          parseNotationType();
        }
        if (attributeTypeHash.ContainsKey(typeString))
        {
          return (int)attributeTypeHash[typeString];
        }
        else
        {
          error("illegal attribute type", typeString, null);
          return ATTRIBUTE_UNDECLARED;
        }
      }
    }


    /**
     * Parse an enumeration.
     * <pre>
     * [59] Enumeration ::= '(' S? Nmtoken (S? '|' S? Nmtoken)* S? ')'
     * </pre>
     * <p>NOTE: the '(' has already been read.
     */

    private void parseEnumeration(bool isNames)
    {
      dataBufferAppend('(');

      // Read the first token.
      skipWhitespace();
      dataBufferAppend(readNmtoken(isNames));
      // Read the remaining tokens.
      skipWhitespace();
      while (!tryRead(')'))
      {
        require('|', "enumeration value");
        dataBufferAppend('|');
        skipWhitespace();
        dataBufferAppend(readNmtoken(isNames));
        skipWhitespace();
      }
      dataBufferAppend(')');
    }


    /**
     * Parse a notation type for an attribute.
     * <pre>
     * [58] NotationType ::= 'NOTATION' S '(' S? NameNtoks
     *		(S? '|' S? name)* S? ')'
     * </pre>
     * <p>NOTE: the 'NOTATION' has already been read
     */

    private void parseNotationType()
    {
      requireWhitespace();
      require('(', "NOTATION");

      parseEnumeration(true);
    }


    /**
     * Parse the default value for an attribute.
     * <pre>
     * [60] DefaultDecl ::= '#REQUIRED' | '#IMPLIED'
     *		| (('#FIXED' S)? AttValue)
     * </pre>
     */

    private void parseDefault(string elementName, string name, int type, string @enum)
    {
      int valueType = ATTRIBUTE_DEFAULT_SPECIFIED;
      string value = null;
      int flags = LIT_ATTRIBUTE | LIT_DISABLE_CREF | LIT_ENTITY_CHECK | LIT_DISABLE_PE;
      // ^^^^^^^^^^^^^^
      // added MHK 20 Mar 2002

      // Note: char refs not checked here, and input not normalized,
      // since it's done correctly later when we actually expand any
      // entity refs.  We ought to report char ref syntax errors now,
      // but don't.  Cost: unused defaults mean unreported WF errs.

      // LIT_ATTRIBUTE forces '<' checks now (ASAP) and turns whitespace
      // chars to spaces (doesn't matter when that's done if it doesn't
      // interfere with char refs expanding to whitespace).

      if (tryRead('#'))
      {
        if (tryRead("FIXED"))
        {
          valueType = ATTRIBUTE_DEFAULT_FIXED;
          requireWhitespace();
          value = readLiteral(flags);
        }
        else if (tryRead("REQUIRED"))
        {
          valueType = ATTRIBUTE_DEFAULT_REQUIRED;
        }
        else if (tryRead("IMPLIED"))
        {
          valueType = ATTRIBUTE_DEFAULT_IMPLIED;
        }
        else
        {
          error("illegal keyword for attribute default value");
        }
      }
      else value = readLiteral(flags);
      setAttribute(elementName, name, type, @enum, value, valueType);
    }


    /**
     * Parse a conditional section.
     * <pre>
     * [61] conditionalSect ::= includeSect || ignoreSect
     * [62] includeSect ::= '&lt;![' S? 'INCLUDE' S? '['
     *		extSubsetDecl ']]&gt;'
     * [63] ignoreSect ::= '&lt;![' S? 'IGNORE' S? '['
     *		ignoreSectContents* ']]&gt;'
     * [64] ignoreSectContents ::= Ignore
     *		('&lt;![' ignoreSectContents* ']]&gt;' Ignore )*
     * [65] Ignore ::= Char* - (Char* ( '&lt;![' | ']]&gt;') Char* )
     * </pre>
     * <p> NOTE: the '&gt;![' has already been read.
     */

    private void parseConditionalSect()
    {
      skipWhitespace();
      if (tryRead("INCLUDE"))
      {
        skipWhitespace();
        require('[', "INCLUDE");
        skipWhitespace();
        while (!tryRead("]]>"))
        {
          parseMarkupdecl();
          skipWhitespace();
        }
      }
      else if (tryRead("IGNORE"))
      {
        skipWhitespace();
        require('[', "IGNORE");
        int nesting = 1;
        char c;
        this._expandPE = false;
        for (int nest = 1; nest > 0;)
        {
          c = readCh();
          switch (c)
          {
            case '<':
              if (tryRead("!["))
              {
                nest++;
              }
              break;
            case ']':
              if (tryRead("]>"))
              {
                nest--;
              }
              break;
          }
        }
        this._expandPE = true;
      }
      else
      {
        error("conditional section must begin with INCLUDE or IGNORE");
      }
    }


    /**
     * Read and interpret a character reference.
     * <pre>
     * [66] CharRef ::= '&#' [0-9]+ ';' | '&#x' [0-9a-fA-F]+ ';'
     * </pre>
     * <p>NOTE: the '&#' has already been read.
     */

    private void parseCharRef()
    {
      int value = 0;
      char c;

      if (tryRead('x'))
      {
        bool done = false;
        while (!done)
        {
          c = readCh();
          switch (c)
          {
            case '0':
            case '1':
            case '2':
            case '3':
            case '4':
            case '5':
            case '6':
            case '7':
            case '8':
            case '9':
            case 'a':
            case 'A':
            case 'b':
            case 'B':
            case 'c':
            case 'C':
            case 'd':
            case 'D':
            case 'e':
            case 'E':
            case 'f':
            case 'F':
              value *= 16;
              value += Convert.ToInt16(c.ToString(CultureInfo.InvariantCulture), 16);
              break;
            case ';':
              done = true;
              break;
            default:
              error("illegal character in character reference", c, null);
              done = true;
              break;
          }
        }
      }
      else
      {

        bool done = false;
        while (!done)
        {
          c = readCh();
          switch (c)
          {
            case '0':
            case '1':
            case '2':
            case '3':
            case '4':
            case '5':
            case '6':
            case '7':
            case '8':
            case '9':
              value *= 10;
              value += Convert.ToInt16(c.ToString(CultureInfo.InvariantCulture), 10);
              break;
            case ';':
              done = true;
              break;
            default:
              error("illegal character in character reference", c, null);
              done = true;
              break;
          }
        }
      }

      // check for character refs being legal XML
      if ((value < 0x0020 && ! (value == '\n' || value == '\t' || value == '\r'))
          || (value >= 0xD800 && value <= 0xDFFF) || value == 0xFFFE || value == 0xFFFF || value > 0x0010ffff) error("illegal XML character reference U+" + value.ToString("X"));

      // Check for surrogates: 00000000 0000xxxx yyyyyyyy zzzzzzzz
      //  (1101|10xx|xxyy|yyyy + 1101|11yy|zzzz|zzzz:
      if (value <= 0x0000ffff)
      {
        // no surrogates needed
        dataBufferAppend((char)value);
      }
      else if (value <= 0x0010ffff)
      {
        value -= 0x10000;
        // > 16 bits, surrogate needed
        dataBufferAppend((char)(0xd800 | (value >> 10)));
        dataBufferAppend((char)(0xdc00 | (value & 0x0003ff)));
      }
      else
      {
        // too big for surrogate
        error(
          "character reference " + value + " is too large for UTF-16",
          value.ToString(CultureInfo.InvariantCulture),
          null);
      }
    }


    /**
     * Parse and expand an entity reference.
     * <pre>
     * [68] EntityRef ::= '&' Name ';'
     * </pre>
     * <p>NOTE: the '&amp;' has already been read.
     * @param externalAllowed External entities are allowed here.
     */

    private void parseEntityRef(bool externalAllowed)
    {
      string name;

      name = readNmtoken(true);
      require(';', "entity reference");
      switch (getEntityType(name))
      {
        case ENTITY_UNDECLARED:
          error("reference to undeclared entity", name, null);
          break;
        case ENTITY_INTERNAL:
          pushString(name, getEntityValue(name));
          break;
        case ENTITY_TEXT:
          if (externalAllowed)
          {
            pushURL(name, getEntityPublicId(name), getEntitySystemId(name), null, null, null, true);
          }
          else
          {
            error("reference to external entity in attribute value.", name, null);
          }
          break;
        case ENTITY_NDATA:
          if (externalAllowed)
          {
            error("unparsed entity reference in content", name, null);
          }
          else
          {
            error("reference to external entity in attribute value.", name, null);
          }
          break;
      }
    }


    /**
     * Parse and expand a parameter entity reference.
     * <pre>
     * [69] PEReference ::= '%' Name ';'
     * </pre>
     * <p>NOTE: the '%' has already been read.
     */

    private void parsePEReference()
    {
      string name;

      name = "%" + readNmtoken(true);
      require(';', "parameter entity reference");
      switch (getEntityType(name))
      {
        case ENTITY_UNDECLARED:
          // this is a validity problem, not a WFC violation ... but
          // we should disable handling of all subsequent declarations
          // unless this is a standalone document
          // warn ("reference to undeclared parameter entity", name, null);

          break;
        case ENTITY_INTERNAL:
          if (this._inLiteral) pushString(name, getEntityValue(name));
          else pushString(name, ' ' + getEntityValue(name) + ' ');
          break;
        case ENTITY_TEXT:
          if (!this._inLiteral) pushString(null, " ");
          pushURL(name, getEntityPublicId(name), getEntitySystemId(name), null, null, null, true);
          if (!this._inLiteral) pushString(null, " ");
          break;
      }
    }

    /**
     * Parse an entity declaration.
     * <pre>
     * [70] EntityDecl ::= GEDecl | PEDecl
     * [71] GEDecl ::= '&lt;!ENTITY' S Name S EntityDef S? '&gt;'
     * [72] PEDecl ::= '&lt;!ENTITY' S '%' S Name S PEDef S? '&gt;'
     * [73] EntityDef ::= EntityValue | (ExternalID NDataDecl?)
     * [74] PEDef ::= EntityValue | ExternalID
     * [75] ExternalID ::= 'SYSTEM' S SystemLiteral
     *		   | 'PUBLIC' S PubidLiteral S SystemLiteral
     * [76] NDataDecl ::= S 'NDATA' S Name
     * </pre>
     * <p>NOTE: the '&lt;!ENTITY' has already been read.
     */

    private void parseEntityDecl()
    {
      bool peFlag = false;

      // Check for a parameter entity.
      this._expandPE = false;
      requireWhitespace();
      if (tryRead('%'))
      {
        peFlag = true;
        requireWhitespace();
      }
      this._expandPE = true;

      // Read the entity name, and prepend
      // '%' if necessary.
      string name = readNmtoken(true);
      if (peFlag)
      {
        name = "%" + name;
      }

      // Read the entity value.
      requireWhitespace();
      char c = readCh();
      unread(c);
      if (c == '"' || c == '\'')
      {
        // Internal entity ... replacement text has expanded refs
        // to characters and PEs, but not to general entities
        string value = readLiteral(0);
        setInternalEntity(name, value);
      }
      else
      {
        // Read the external IDs
        string[] ids = readExternalIds(false);
        if (ids[1] == null)
        {
          error("system identifer missing", name, null);
        }

        // Check for NDATA declaration.
        bool white = tryWhitespace();
        if (!peFlag && tryRead("NDATA"))
        {
          if (!white) error("whitespace required before NDATA");
          requireWhitespace();
          string notationName = readNmtoken(true);
          setExternalDataEntity(name, ids[0], ids[1], notationName);
        }
        else
        {
          setExternalTextEntity(name, ids[0], ids[1]);
        }
      }

      // Finish the declaration.
      skipWhitespace();
      require('>', "NDATA");
    }


    /**
     * Parse a notation declaration.
     * <pre>
     * [82] NotationDecl ::= '&lt;!NOTATION' S Name S
     *		(ExternalID | PublicID) S? '&gt;'
     * [83] PublicID ::= 'PUBLIC' S PubidLiteral
     * </pre>
     * <P>NOTE: the '&lt;!NOTATION' has already been read.
     */

    private void parseNotationDecl()
    {
      string nname;
      string[] ids ;


      requireWhitespace();
      nname = readNmtoken(true);

      requireWhitespace();

      // Read the external identifiers.
      ids = readExternalIds(true);
      if (ids[0] == null && ids[1] == null)
      {
        error("external identifer missing", nname, null);
      }

      // Register the notation.
      setNotation(nname, ids[0], ids[1]);

      skipWhitespace();
      require('>', "notation declaration");
    }


    /**
     * Parse character data.
     * <pre>
     * [14] CharData ::= [^&lt;&amp;]* - ([^&lt;&amp;]* ']]&gt;' [^&lt;&amp;]*)
     * </pre>
     */

    private void parseCharData()
    {
      char c;

      // Start with a little cheat -- in most
      // cases, the entire sequence of
      // character data will already be in
      // the readBuffer; if not, fall through to
      // the normal approach.
      if (USE_CHEATS)
      {
        int lineAugment = 0;
        int columnAugment = 0;

        for (int i = this._readBufferPos; i < this._readBufferLength; i++)
        {

          switch (c = this._readBuffer[i])
          {
            case '\n':
              lineAugment++;
              columnAugment = 0;
              break;
            case '&':
            case '<':
              int start = this._readBufferPos;
              columnAugment++;
              this._readBufferPos = i;
              if (lineAugment > 0)
              {
                this._line += lineAugment;
                this._column = columnAugment;
              }
              else
              {
                this._column += columnAugment;
              }
              dataBufferAppend(this._readBuffer, start, i - start);
              return;
            case ']':
              // XXX missing two end-of-buffer cases
              if ((i + 2) < this._readBufferLength)
              {
                if (this._readBuffer[i + 1] == ']' && this._readBuffer[i + 2] == '>')
                {
                  error("character data may not contain ']]>'");
                }
              }
              columnAugment++;
              break;
            default:
              if (c < 0x0020 || c > 0xFFFD) error("illegal XML character U+" + ((int)c).ToString("X"));
              break;
              // FALLTHROUGH
            case '\r':
            case '\t':
              columnAugment++;
              break;
          }
        }
      }

      // OK, the cheat didn't work; start over
      // and do it by the book.

      int closeSquareBracketCount = 0;
      while (true)
      {
        c = readCh();
        switch (c)
        {
          case '<':
          case '&':
            unread(c);
            return;
          case ']':
            closeSquareBracketCount++;
            dataBufferAppend(c);
            break;
          case '>':
            if (closeSquareBracketCount >= 2)
            {
              // we've hit ']]>'
              error("']]>' is not allowed here");
              break;
            }
            break;
            // fall-through                
          default:
            closeSquareBracketCount = 0;
            dataBufferAppend(c);
            break;
        }
      }
    }


    //////////////////////////////////////////////////////////////////////
    // High-level reading and scanning methods.
    //////////////////////////////////////////////////////////////////////

    /**
     * Require whitespace characters.
     */

    private void requireWhitespace()
    {
      char c = readCh();
      if (isWhitespace(c))
      {
        skipWhitespace();
      }
      else
      {
        error("whitespace required", c, null);
      }
    }


    /**
     * Parse whitespace characters, and leave them in the data buffer.
     */

    private void parseWhitespace() // method no longer used - MHK
    {
      char c = readCh();
      while (isWhitespace(c))
      {
        dataBufferAppend(c);
        c = readCh();
      }
      unread(c);
    }


    /**
     * Skip whitespace characters.
     * <pre>
     * [3] S ::= (#x20 | #x9 | #xd | #xa)+
     * </pre>
     */

    private void skipWhitespace()
    {
      // Start with a little cheat.  Most of
      // the time, the white space will fall
      // within the current read buffer; if
      // not, then fall through.
      if (USE_CHEATS)
      {
        int lineAugment = 0;
        int columnAugment = 0;

        bool done = false;
        for (int i = this._readBufferPos; i < this._readBufferLength; i++)
        {
          if (done)
          {
            break;
          }
          switch (this._readBuffer[i])
          {
            case ' ':
            case '\t':
            case '\r':
              columnAugment++;
              break;
            case '\n':
              lineAugment++;
              columnAugment = 0;
              break;
            case '%':
              if (this._expandPE)
              {
                done = true;
              }
              break;
              // else fall through...
            default:
              this._readBufferPos = i;
              if (lineAugment > 0)
              {
                this._line += lineAugment;
                this._column = columnAugment;
              }
              else
              {
                this._column += columnAugment;
              }
              return;
          }
        }
      }

      // OK, do it by the book.
      char c = readCh();
      while (isWhitespace(c))
      {
        c = readCh();
      }
      unread(c);
    }


    /**
     * Read a name or (when parsing an enumeration) name token.
     * <pre>
     * [5] Name ::= (Letter | '_' | ':') (NameChar)*
     * [7] Nmtoken ::= (NameChar)+
     * </pre>
     */

    private string readNmtoken(bool isName)
    {
      char c;

      bool done = false;

      if (USE_CHEATS)
      {
        for (int i = this._readBufferPos; i < this._readBufferLength; i++)
        {
          if (done)
          {
            break;
          }
          c = this._readBuffer[i];
          switch (c)
          {
            case '%':
              if (this._expandPE)
              {
                done = true;
              }
              break;
              // else fall through...

              // What may legitimately come AFTER a name/nmtoken?
            case '<':
            case '>':
            case '&':
            case ',':
            case '|':
            case '*':
            case '+':
            case '?':
            case ')':
            case '=':
            case '\'':
            case '"':
            case '[':
            case ' ':
            case '\t':
            case '\r':
            case '\n':
            case ';':
            case '/':
              int start = this._readBufferPos;
              if (i == start) error("name expected", this._readBuffer[i], null);
              this._readBufferPos = i;
              return intern(this._readBuffer, start, i - start);

            default:
              // punt on exact tests from Appendix A; approximate
              // them using the Unicode ID start/part rules
              if (i == this._readBufferPos && isName)
              {
                if (!isUnicodeIdentifierStart(c) && c != ':' && c != '_') error("Not a name start character, U+" + ((int)c).ToString("X"));
              }
              else if (!isUnicodeIdentifierPart(c) && c != '-' && c != ':' && c != '_' && c != '.' && !isExtender(c)) error("Not a name character, U+" + ((int)c).ToString("X"));
              break;
          }
        }
      }

      this._nameBufferPos = 0;

      // Read the first character.
      loop:
      while (true)
      {
        c = readCh();
        switch (c)
        {
          case '%':
          case '<':
          case '>':
          case '&':
          case ',':
          case '|':
          case '*':
          case '+':
          case '?':
          case ')':
          case '=':
          case '\'':
          case '"':
          case '[':
          case ' ':
          case '\t':
          case '\n':
          case '\r':
          case ';':
          case '/':
            unread(c);
            if (this._nameBufferPos == 0)
            {
              error("name expected");
            }
            // punt on exact tests from Appendix A, but approximate them
            if (isName && !isUnicodeIdentifierStart(this._nameBuffer[0]) && ":_".IndexOf(this._nameBuffer[0]) == -1) error("Not a name start character, U+" + ((int)this._nameBuffer[0]).ToString("X"));
            string s = intern(this._nameBuffer, 0, this._nameBufferPos);
            this._nameBufferPos = 0;
            return s;
          default:
            // punt on exact tests from Appendix A, but approximate them

            if ((this._nameBufferPos != 0 || !isName) && !isUnicodeIdentifierPart(c) && ":-_.".IndexOf(c) == -1
                && !isExtender(c)) error("Not a name character, U+" + ((int)c).ToString("X"));
            if (this._nameBufferPos >= this._nameBuffer.Length) this._nameBuffer = (char[])extendArray(this._nameBuffer, this._nameBuffer.Length, this._nameBufferPos);
            this._nameBuffer[this._nameBufferPos++] = c;
            break;
        }
      }
    }

    private static bool isExtender(char c)
    {
      // [88] Extender ::= ...
      return c == 0x00b7 || c == 0x02d0 || c == 0x02d1 || c == 0x0387 || c == 0x0640 || c == 0x0e46 || c == 0x0ec6
             || c == 0x3005 || (c >= 0x3031 && c <= 0x3035) || (c >= 0x309d && c <= 0x309e)
             || (c >= 0x30fc && c <= 0x30fe);
    }


    /**
     * Read a literal.  With matching single or double quotes as
     * delimiters (and not embedded!) this is used to parse:
     * <pre>
     *	[9] EntityValue ::= ... ([^%&amp;] | PEReference | Reference)* ...
     *	[10] AttValue ::= ... ([^<&] | Reference)* ...
     *	[11] SystemLiteral ::= ... (URLchar - "'")* ...
     *	[12] PubidLiteral ::= ... (PubidChar - "'")* ...
     * </pre>
     * as well as the quoted strings in XML and text declarations
     * (for version, encoding, and standalone) which have their
     * own constraints.
     */

    private string readLiteral(int flags)
    {
      char delim, c;
      int startLine = this._line;
      bool saved = this._expandPE;

      // Find the first delimiter.
      delim = readCh();
      if (delim != '"' && delim != '\'' && delim != (char)0)
      {
        error("expected '\"' or \"'\"", delim, null);
        return null;
      }
      this._inLiteral = true;
      if ((flags & LIT_DISABLE_PE) != 0) this._expandPE = false;

      // Each level of input source has its own buffer; remember
      // ours, so we won't read the ending delimiter from any
      // other input source, regardless of entity processing.
      char[] ourBuf = this._readBuffer;

      // Read the literal.
      try
      {
        c = readCh();
        while (! (c == delim && this._readBuffer == ourBuf))
        {
          switch (c)
          {
              // attributes and public ids are normalized
              // in almost the same ways
            case '\n':
            case '\r':
              if ((flags & (LIT_ATTRIBUTE | LIT_PUBID)) != 0) c = ' ';
              break;
            case '\t':
              if ((flags & LIT_ATTRIBUTE) != 0) c = ' ';
              break;
            case '&':
              c = readCh();
              // Char refs are expanded immediately, except for
              // all the cases where it's deferred.
              if (c == '#')
              {
                if ((flags & LIT_DISABLE_CREF) != 0)
                {
                  dataBufferAppend('&');
                  continue;
                }
                parseCharRef();

                // It looks like an entity ref ...
              }
              else
              {
                unread(c);
                // Expand it?
                if ((flags & LIT_ENTITY_REF) > 0)
                {
                  parseEntityRef(false);

                  // Is it just data?
                }
                else if ((flags & LIT_DISABLE_EREF) != 0)
                {
                  dataBufferAppend('&');

                  // OK, it will be an entity ref -- expanded later.
                }
                else
                {
                  string name = readNmtoken(true);
                  require(';', "entity reference");
                  if ((flags & LIT_ENTITY_CHECK) != 0 && getEntityType(name) == ENTITY_UNDECLARED)
                  {
                    // Possibly a validity error, shouldn't report it?
                    error("General entity '" + name + "' must be declared before use");
                  }
                  dataBufferAppend('&');
                  dataBufferAppend(name);
                  dataBufferAppend(';');
                }
              }
              c = readCh();
              continue;

            case '<':
              // and why?  Perhaps so "&foo;" expands the same
              // inside and outside an attribute?
              if ((flags & LIT_ATTRIBUTE) != 0) error("attribute values may not contain '<'");
              break;

              // We don't worry about case '%' and PE refs, readCh does.

            default:
              break;
          }
          dataBufferAppend(c);
          c = readCh();
        }
      }
      catch (EndOfStreamException e)
      {
        error(
          "end of input while looking for delimiter (started on line " + startLine + ')',
          null,
          delim.ToString(CultureInfo.InvariantCulture));
      }
      this._inLiteral = false;
      this._expandPE = saved;

      // Normalise whitespace if necessary.
      if ((flags & LIT_NORMALIZE) > 0)
      {
        dataBufferNormalize();
      }

      // Return the value.
      return dataBufferToString();
    }


    /**
     * Try reading external identifiers.
     * A system identifier is not required for notations.
     * @param inNotation Are we in a notation?
     * @return A two-member string array containing the identifiers.
     */

    private string[] readExternalIds(bool inNotation)
    {
      char c;
      string[] ids = new string[2];
      int flags = LIT_DISABLE_CREF | LIT_DISABLE_PE | LIT_DISABLE_EREF;

      if (tryRead("PUBLIC"))
      {
        requireWhitespace();
        ids[0] = readLiteral(LIT_NORMALIZE | LIT_PUBID | flags);
        if (inNotation)
        {
          skipWhitespace();
          c = readCh();
          unread(c);
          if (c == '"' || c == '\'')
          {
            ids[1] = readLiteral(flags);
          }
        }
        else
        {
          requireWhitespace();
          ids[1] = readLiteral(flags);
        }

        for (int i = 0; i < ids[0].Length; i++)
        {
          c = ids[0][i];
          if (c >= 'a' && c <= 'z') continue;
          if (c >= 'A' && c <= 'Z') continue;
          if (" \r\n0123456789-' ()+,./:=?;!*#@$_%".IndexOf(c) != -1) continue;
          error("illegal PUBLIC id character U+" + ((int)c).ToString("X"));
        }
      }
      else if (tryRead("SYSTEM"))
      {
        requireWhitespace();
        ids[1] = readLiteral(flags);
      }

      // XXX should normalize system IDs as follows:
      // - Convert to UTF-8
      // - Map reserved and non-ASCII characters to %HH

      return ids;
    }


    /**
     * Test if a character is whitespace.
     * <pre>
     * [3] S ::= (#x20 | #x9 | #xd | #xa)+
     * </pre>
     * @param c The character to test.
     * @return true if the character is whitespace.
     */

    private bool isWhitespace(char c)
    {
      if (c > 0x20) return false;
      if (c == 0x20 || c == 0x0a || c == 0x09 || c == 0x0d) return true;
      return false; // illegal ...
    }


    //////////////////////////////////////////////////////////////////////
    // Utility routines.
    //////////////////////////////////////////////////////////////////////


    /**
     * Add a character to the data buffer.
     */

    private void dataBufferAppend(char c)
    {
      // Expand buffer if necessary.
      if (this._dataBufferPos >= this._dataBuffer.Length) this._dataBuffer = (char[])extendArray(this._dataBuffer, this._dataBuffer.Length, this._dataBufferPos);
      this._dataBuffer[this._dataBufferPos++] = c;
    }


    /**
     * Add a string to the data buffer.
     */

    private void dataBufferAppend(string s)
    {
      dataBufferAppend(s.ToCharArray(), 0, s.Length);
    }


    /**
     * Append (part of) a character array to the data buffer.
     */

    private void dataBufferAppend(char[] ch, int start, int length)
    {
      this._dataBuffer = (char[])extendArray(this._dataBuffer, this._dataBuffer.Length, this._dataBufferPos + length);

      Array.Copy(ch, start, this._dataBuffer, this._dataBufferPos, length);
      this._dataBufferPos += length;
    }


    /**
     * Normalise spaces in the data buffer.
     */

    private void dataBufferNormalize()
    {
      int i = 0;
      int j = 0;
      int end = this._dataBufferPos;

      // Skip spaces at the start.
      while (j < end && this._dataBuffer[j] == ' ')
      {
        j++;
      }

      // Skip whitespace at the end.
      while (end > j && this._dataBuffer[end - 1] == ' ')
      {
        end --;
      }

      // Start copying to the left.
      while (j < end)
      {

        char c = this._dataBuffer[j++];

        // Normalise all other whitespace to
        // a single space.
        if (c == ' ')
        {
          while (j < end && this._dataBuffer[j++] == ' ')
          {
          }

          this._dataBuffer[i++] = ' ';
          this._dataBuffer[i++] = this._dataBuffer[j - 1];
        }
        else
        {
          this._dataBuffer[i++] = c;
        }
      }

      // The new length is <= the old one.
      this._dataBufferPos = i;
    }


    /**
     * Convert the data buffer to a string.
     */

    private string dataBufferToString()
    {
      string s = new string(this._dataBuffer, 0, this._dataBufferPos);
      this._dataBufferPos = 0;
      return s;
    }


    /**
     * Flush the contents of the data buffer to the handler, as
     * appropriate, and reset the buffer for new input.
     */

    private void dataBufferFlush()
    {
      if (this._currentElementContent == CONTENT_ELEMENTS && this._dataBufferPos > 0 && !this._inCDATA)
      {
        // We can't just trust the buffer to be whitespace, there
        // are cases when it isn't
        for (int i = 0; i < this._dataBufferPos; i++)
        {
          if (!isWhitespace(this._dataBuffer[i]))
          {
            this._handler.CharData(this._dataBuffer, 0, this._dataBufferPos);
            this._dataBufferPos = 0;
          }
        }
        if (this._dataBufferPos > 0)
        {
          this._handler.IgnorableWhitespace(this._dataBuffer, 0, this._dataBufferPos);
          this._dataBufferPos = 0;
        }
      }
      else if (this._dataBufferPos > 0)
      {
        this._handler.CharData(this._dataBuffer, 0, this._dataBufferPos);
        this._dataBufferPos = 0;
      }
    }


    /**
     * Require a string to appear, or throw an exception.
     * <p><em>Precondition:</em> Entity expansion is not required.
     * <p><em>Precondition:</em> data buffer has no characters that
     * will get sent to the application.
     */

    private void require(string delim, string context)
    {
      int length = delim.Length;
      char[] ch;

      if (length < this._dataBuffer.Length)
      {
        ch = delim.ToCharArray(0, length);
      }
      else ch = delim.ToCharArray();

      if (USE_CHEATS && length <= (this._readBufferLength - this._readBufferPos))
      {
        int offset = this._readBufferPos;

        for (int i = 0; i < length; i++, offset++) if (ch[i] != this._readBuffer[offset]) error("unexpected characters in " + context, null, delim);
        this._readBufferPos = offset;

      }
      else
      {
        for (int i = 0; i < length; i++) require(ch[i], delim);
      }
    }


    /**
     * Require a character to appear, or throw an exception.
     */

    private void require(char delim, string after)
    {
      char c = readCh();

      if (c != delim)
      {
        error("unexpected character after " + after, c, delim + "");
      }
    }


    /**
     * Create an interned string from a character array.
     * &AElig;lfred uses this method to create an interned version
     * of all names and name tokens, so that it can test equality
     * with <code>==</code> instead of <code>string.Equals ()</code>.
     *
     * <p>This is much more efficient than constructing a non-interned
     * string first, and then interning it.
     *
     * @param ch an array of characters for building the string.
     * @param start the starting position in the array.
     * @param length the number of characters to place in the string.
     * @return an interned string.
     * @see #intern (string)
     * @see java.lang.string#intern
     */

    public string intern(char[] ch, int start, int length)
    {
      int index = 0;
      int hash = 0;
      object[] bucket;

      // Generate a hash code.
      for (int i = start; i < start + length; i++) hash = 31 * hash + ch[i];
      hash = (hash & 0x7fffffff) % SYMBOL_TABLE_LENGTH;

      // Get the bucket -- consists of {array,string} pairs
      if ((bucket = this._symbolTable[hash]) == null)
      {
        // first string in this bucket
        bucket = new object[8];

        // Search for a matching tuple, and
        // return the string if we find one.
      }
      else
      {
        while (index < bucket.Length)
        {
          char[] chFound = (char[])bucket[index];

          // Stop when we hit a null index.
          if (chFound == null) break;

          // If they're the same length, check for a match.
          if (chFound.Length == length)
          {
            for (int i = 0; i < chFound.Length; i++)
            {
              // continue search on failure
              if (ch[start + i] != chFound[i])
              {
                break;
              }
              else if (i == length - 1)
              {
                // That's it, we have a match!
                return (string)bucket[index + 1];
              }
            }
          }
          index += 2;
        }
        // Not found -- we'll have to add it.

        // Do we have to grow the bucket?
        bucket = (object[])extendArray(bucket, bucket.Length, index);
      }
      this._symbolTable[hash] = bucket;

      // OK, add it to the end of the bucket -- "local" interning.
      // Intern "globally" to let applications share interning benefits.
      string s = string.Intern(new string(ch, start, length));
      bucket[index] = s.ToCharArray();
      bucket[index + 1] = s;
      return s;
    }


    /**
     * Ensure the capacity of an array, allocating a new one if
     * necessary.  Usually called only a handful of times.
     */

    private object extendArray(object array, int currentSize, int requiredSize)
    {
      if (requiredSize < currentSize)
      {
        return array;
      }
      else
      {
        object newArray = null;
        int newSize = currentSize * 2;

        if (newSize <= requiredSize) newSize = requiredSize + 1;

        if (array is char[])
        {
          newArray = new char[newSize];
          Array.Copy((char[])array, 0, (char[])newArray, 0, currentSize);
        }
        else if (array is object[])
        {
          newArray = new object[newSize];
          Array.Copy((object[])array, 0, (object[])newArray, 0, currentSize);
        }
        else 
          throw new Exception();


        return newArray;
      }
    }


    //////////////////////////////////////////////////////////////////////
    // XML query routines.
    //////////////////////////////////////////////////////////////////////


    //
    // Elements
    //

    /**
     * Get the declared elements for an XML document.
     * <p>The results will be valid only after the DTD (if any) has been
     * parsed.
     * @return An enumeration of all element types declared for this
     *	 document (as Strings).
     * @see #getElementContentType
     * @see #getElementContentModel
     */

    public ICollection declaredElements()
    {
      return this._elementInfo.Keys;
    }


    /**
     * Look up the content type of an element.
     * @param element element info vector
     * @param defaultType value for null vector
     * @return An integer constant representing the content type.
     * @see #CONTENT_UNDECLARED
     * @see #CONTENT_ANY
     * @see #CONTENT_EMPTY
     * @see #CONTENT_MIXED
     * @see #CONTENT_ELEMENTS
     */

    private int getContentType(object[] element, int defaultType)
    {
      int retval;

      if (element == null) return defaultType;
      retval = ((int)element[0]);
      if (retval == CONTENT_UNDECLARED) retval = defaultType;
      return retval;
    }


    /**
     * Look up the content type of an element.
     * @param name The element type name.
     * @return An integer constant representing the content type.
     * @see #getElementContentModel
     * @see #CONTENT_UNDECLARED
     * @see #CONTENT_ANY
     * @see #CONTENT_EMPTY
     * @see #CONTENT_MIXED
     * @see #CONTENT_ELEMENTS
     */

    public int getElementContentType(string name)
    {
      object[] element = (object[])this._elementInfo[name];
      return getContentType(element, CONTENT_UNDECLARED);
    }


    /**
     * Look up the content model of an element.
     * <p>The result will always be null unless the content type is
     * CONTENT_ELEMENTS or CONTENT_MIXED.
     * @param name The element type name.
     * @return The normalised content model, as a string.
     * @see #getElementContentType
     */

    public string getElementContentModel(string name)
    {
      if (this._elementInfo.ContainsKey(name))
      {
        object[] element = (object[])this._elementInfo[name];
        return (string)element[1];
      }
      return null;
    }


    /**
     * Register an element.
     * Array format:
     *  [0] element type name
     *  [1] content model (mixed, elements only)
     *  [2] attribute hash table
     */

    private void setElement(string name, int contentType, string contentModel, Hashtable attributes)
    {
      object[] element = this._elementInfo.ContainsKey(name) ? (object[])this._elementInfo[name] : null;

      // first <!ELEMENT ...> or <!ATTLIST ...> for this type?
      if (element == null)
      {
        element = new object[3];
        element[0] = contentType;
        element[1] = contentModel;
        element[2] = attributes;
        this._elementInfo.Add(name, element);
        return;
      }

      // <!ELEMENT ...> declaration?	
      if (contentType != CONTENT_UNDECLARED)
      {
        // ... following an associated <!ATTLIST ...>
        if (((int)element[0]) == CONTENT_UNDECLARED)
        {
          element[0] = contentType;
          element[1] = contentModel;
        }
        else
        {
          // VC: Unique Element Type Declaration
          //verror ("multiple declarations for element type: " + name);
        }
      }

        // first <!ATTLIST ...>, before <!ELEMENT ...> ?
      else if (attributes != null)
      {
        element[2] = attributes;
      }

    }


    /**
     * Look up the attribute hash table for an element.
     * The hash table is the second item in the element array.
     */

    private Hashtable getElementAttributes(string name)
    {
      object[] element = this._elementInfo.ContainsKey(name) ? (object[])this._elementInfo[name] : null;
      if (element == null)
      {
        return null;
      }
      else
      {
        return (Hashtable)element[2];
      }
    }



    //
    // Attributes
    //

    /**
     * Get the declared attributes for an element type.
     * @param elname The name of the element type.
     * @return An Enumeration of all the attributes declared for
     *	 a specific element type.  The results will be valid only
     *	 after the DTD (if any) has been parsed.
     * @see #getAttributeType
     * @see #getAttributeEnumeration
     * @see #getAttributeDefaultValueType
     * @see #getAttributeDefaultValue
     * @see #getAttributeExpandedValue
     */

    private ICollection declaredAttributes(object[] element)
    {
      Hashtable attlist;

      if (element == null) return null;
      if ((attlist = (Hashtable)element[2]) == null) return null;
      return attlist.Keys;
    }

    /**
     * Get the declared attributes for an element type.
     * @param elname The name of the element type.
     * @return An Enumeration of all the attributes declared for
     *	 a specific element type.  The results will be valid only
     *	 after the DTD (if any) has been parsed.
     * @see #getAttributeType
     * @see #getAttributeEnumeration
     * @see #getAttributeDefaultValueType
     * @see #getAttributeDefaultValue
     * @see #getAttributeExpandedValue
     */

    public ICollection declaredAttributes(string elname)
    {
      return declaredAttributes((object[])this._elementInfo[elname]);
    }


    /**
     * Retrieve the declared type of an attribute.
     * @param name The name of the associated element.
     * @param aname The name of the attribute.
     * @return An integer constant representing the attribute type.
     * @see #ATTRIBUTE_UNDECLARED
     * @see #ATTRIBUTE_CDATA
     * @see #ATTRIBUTE_ID
     * @see #ATTRIBUTE_IDREF
     * @see #ATTRIBUTE_IDREFS
     * @see #ATTRIBUTE_ENTITY
     * @see #ATTRIBUTE_ENTITIES
     * @see #ATTRIBUTE_NMTOKEN
     * @see #ATTRIBUTE_NMTOKENS
     * @see #ATTRIBUTE_ENUMERATED
     * @see #ATTRIBUTE_NOTATION
     */

    public int getAttributeType(string name, string aname)
    {
      object[] attribute = getAttribute(name, aname);
      if (attribute == null)
      {
        return ATTRIBUTE_UNDECLARED;
      }
      else
      {
        return ((int)attribute[0]);
      }
    }


    /**
     * Retrieve the allowed values for an enumerated attribute type.
     * @param name The name of the associated element.
     * @param aname The name of the attribute.
     * @return A string containing the token list.
     * @see #ATTRIBUTE_ENUMERATED
     * @see #ATTRIBUTE_NOTATION
     */

    public string getAttributeEnumeration(string name, string aname)
    {
      object[] attribute = getAttribute(name, aname);
      if (attribute == null)
      {
        return null;
      }
      else
      {
        return (string)attribute[3];
      }
    }


    /**
     * Retrieve the default value of a declared attribute.
     * @param name The name of the associated element.
     * @param aname The name of the attribute.
     * @return The default value, or null if the attribute was
     *	 #IMPLIED or simply undeclared and unspecified.
     * @see #getAttributeExpandedValue
     */

    public string getAttributeDefaultValue(string name, string aname)
    {
      object[] attribute = getAttribute(name, aname);
      if (attribute == null)
      {
        return null;
      }
      else
      {
        return (string)attribute[1];
      }
    }


    /**
     * Retrieve the expanded value of a declared attribute.
     * <p>General entities (and char refs) will be expanded (once).
     * @param name The name of the associated element.
     * @param aname The name of the attribute.
     * @return The expanded default value, or null if the attribute was
     *	 #IMPLIED or simply undeclared
     * @see #getAttributeDefaultValue
     */

    public string getAttributeExpandedValue(string name, string aname)
    {
      object[] attribute = getAttribute(name, aname);

      if (attribute == null)
      {
        return null;
      }
      else if (attribute[4] == null && attribute[1] != null)
      {
        // we MUST use the same buf for both quotes else the literal
        // can't be properly terminated
        char[] buf = new char[1];
        int flags = LIT_ENTITY_REF | LIT_ATTRIBUTE;
        int type = getAttributeType(name, aname);

        if (type != ATTRIBUTE_CDATA && type != ATTRIBUTE_UNDECLARED) flags |= LIT_NORMALIZE;
        buf[0] = '"';
        pushCharArray(null, buf, 0, 1);
        pushString(null, (string)attribute[1]);
        pushCharArray(null, buf, 0, 1);
        attribute[4] = readLiteral(flags);
      }
      return (string)attribute[4];
    }


    /**
     * Retrieve the default value type of a declared attribute.
     * @see #ATTRIBUTE_DEFAULT_SPECIFIED
     * @see #ATTRIBUTE_DEFAULT_IMPLIED
     * @see #ATTRIBUTE_DEFAULT_REQUIRED
     * @see #ATTRIBUTE_DEFAULT_FIXED
     */

    public int getAttributeDefaultValueType(string name, string aname)
    {
      object[] attribute = getAttribute(name, aname);
      if (attribute == null)
      {
        return ATTRIBUTE_DEFAULT_UNDECLARED;
      }
      else
      {
        return ((int)attribute[2]);
      }
    }


    /**
     * Register an attribute declaration for later retrieval.
     * Format:
     * - String type
     * - String default value
     * - int value type
     */

    private void setAttribute(string elName, string name, int type, string enumeration, string value, int valueType)
    {
      Hashtable attlist;

      // Create a new hashtable if necessary.
      attlist = getElementAttributes(elName);
      if (attlist == null)
      {
        attlist = new Hashtable();
      }

      // ignore multiple attribute declarations!
      if (attlist.ContainsKey(name))
      {
        // warn ...
        return;
      }
      else
      {
        object[] attribute = new object[5];
        attribute[0] = type;
        attribute[1] = value;
        attribute[2] = valueType;
        attribute[3] = enumeration;
        attribute[4] = null;
        attlist.Add(name, attribute);

        // save; but don't overwrite any existing <!ELEMENT ...>
        setElement(elName, CONTENT_UNDECLARED, null, attlist);
      }
    }


    /**
     * Retrieve the five-member array representing an
     * attribute declaration.
     */

    private object[] getAttribute(string elName, string name)
    {
      Hashtable attlist = getElementAttributes(elName);
      if (attlist == null)
      {
        return null;
      }

      return (object[])attlist[name];
    }


    //
    // Entities
    //

    /**
     * Get declared entities.
     * @return An Enumeration of all the entities declared for
     *	 this XML document.  The results will be valid only
     *	 after the DTD (if any) has been parsed.
     * @see #getEntityType
     * @see #getEntityPublicId
     * @see #getEntitySystemId
     * @see #getEntityValue
     * @see #getEntityNotationName
     */

    public ICollection declaredEntities()
    {
      return this._entityInfo.Keys;
    }


    /**
     * Find the type of an entity.
     * @returns An integer constant representing the entity type.
     * @see #ENTITY_UNDECLARED
     * @see #ENTITY_INTERNAL
     * @see #ENTITY_NDATA
     * @see #ENTITY_TEXT
     */

    public int getEntityType(string ename)
    {
      object[] entity = this._entityInfo.ContainsKey(ename) ? (object[])this._entityInfo[ename] : null;
      if (entity == null)
      {
        return ENTITY_UNDECLARED;
      }
      else
      {
        return ((int)entity[0]);
      }
    }


    /**
     * Return an external entity's public identifier, if any.
     * @param ename The name of the external entity.
     * @return The entity's system identifier, or null if the
     *	 entity was not declared, if it is not an
     *	 external entity, or if no public identifier was
     *	 provided.
     * @see #getEntityType
     */

    public string getEntityPublicId(string ename)
    {
      object[] entity = this._entityInfo.ContainsKey(ename) ? (object[])this._entityInfo[ename] : null;
      if (entity == null)
      {
        return null;
      }
      else
      {
        return (string)entity[1];
      }
    }


    /**
     * Return an external entity's system identifier.
     * @param ename The name of the external entity.
     * @return The entity's system identifier, or null if the
     *	 entity was not declared, or if it is not an
     *	 external entity. Change made by MHK: The system identifier
     *   is returned as an absolute URL, resolved relative to the entity
     *   it was contained in.
     * @see #getEntityType
     */

    public string getEntitySystemId(string ename)
    {
      object[] entity = this._entityInfo.ContainsKey(ename) ? (object[])this._entityInfo[ename] : null;
      if (entity == null)
      {
        return null;
      }
      else
      {
        try
        {
          string relativeURI = (string)entity[2];
          Uri baseURI = (Uri)entity[5];
          if (baseURI == null) return relativeURI;
          Uri absoluteURI = new Uri(baseURI, relativeURI);
          return absoluteURI.ToString();
        }
        catch (IOException err)
        {
          // ignore the exception, a user entity resolver may be able
          // to do something; if not, the error will be caught later
          return (string)entity[2];
        }
      }
    }


    /**
     * Return the value of an internal entity.
     * @param ename The name of the internal entity.
     * @return The entity's value, or null if the entity was
     *	 not declared, or if it is not an internal entity.
     * @see #getEntityType
     */

    public string getEntityValue(string ename)
    {
      object[] entity = this._entityInfo.ContainsKey(ename) ? (object[])this._entityInfo[ename] : null;
      if (entity == null)
      {
        return null;
      }
      else
      {
        return (string)entity[3];
      }
    }


    /**
     * Get the notation name associated with an NDATA entity.
     * @param ename The NDATA entity name.
     * @return The associated notation name, or null if the
     *	 entity was not declared, or if it is not an
     *	 NDATA entity.
     * @see #getEntityType
     */

    public string getEntityNotationName(string eName)
    {
      object[] entity = this._entityInfo.ContainsKey(eName) ? (object[])this._entityInfo[eName] : null;
      if (entity == null)
      {
        return null;
      }
      else
      {
        return (string)entity[4];
      }
    }


    /**
     * Register an entity declaration for later retrieval.
     */

    private void setInternalEntity(string eName, string value)
    {
      setEntity(eName, ENTITY_INTERNAL, null, null, value, null);
    }


    /**
     * Register an external data entity.
     */

    private void setExternalDataEntity(string eName, string pubid, string sysid, string nName)
    {
      setEntity(eName, ENTITY_NDATA, pubid, sysid, null, nName);
    }


    /**
     * Register an external text entity.
     */

    private void setExternalTextEntity(string eName, string pubid, string sysid)
    {
      setEntity(eName, ENTITY_TEXT, pubid, sysid, null, null);
    }


    /**
     * Register an entity declaration for later retrieval.
     */

    private void setEntity(string eName, int eClass, string pubid, string sysid, string value, string nName)
    {
      object[] entity;

      if (!this._entityInfo.ContainsKey(eName))
      {
        entity = new object[6];
        entity[0] = eClass;
        entity[1] = pubid;
        entity[2] = sysid;
        entity[3] = value;
        entity[4] = nName;
        entity[5] = (this._externalEntity == null ? null : this._externalEntity.RequestUri);
        // added MHK: provides base URI for resolution

        this._entityInfo.Add(eName, entity);
      }
    }


    //
    // Notations.
    //

    /**
     * Get declared notations.
     * @return An Enumeration of all the notations declared for
     *	 this XML document.  The results will be valid only
     *	 after the DTD (if any) has been parsed.
     * @see #getNotationPublicId
     * @see #getNotationSystemId
     */

    public ICollection declaredNotations()
    {
      return this._notationInfo.Keys;
    }


    /**
     * Look up the public identifier for a notation.
     * You will normally use this method to look up a notation
     * that was provided as an attribute value or for an NDATA entity.
     * @param nname The name of the notation.
     * @return A string containing the public identifier, or null
     *	 if none was provided or if no such notation was
     *	 declared.
     * @see #getNotationSystemId
     */

    public string getNotationPublicId(string nname)
    {
      object[] notation = this._notationInfo.ContainsKey(nname) ? (object[])this._notationInfo[nname] : null;
      if (notation == null)
      {
        return null;
      }
      else
      {
        return (string)notation[0];
      }
    }


    /**
     * Look up the system identifier for a notation.
     * You will normally use this method to look up a notation
     * that was provided as an attribute value or for an NDATA entity.
     * @param nname The name of the notation.
     * @return A string containing the system identifier, or null
     *	 if no such notation was declared.
     * @see #getNotationPublicId
     */

    public string getNotationSystemId(string nname)
    {
      object[] notation = this._notationInfo.ContainsKey(nname) ? (object[])this._notationInfo[nname] : null;
      if (notation == null)
      {
        return null;
      }
      else
      {
        return (string)notation[1];
      }
    }


    /**
     * Register a notation declaration for later retrieval.
     * Format:
     * - public id
     * - system id
     */

    private void setNotation(string nname, string pubid, string sysid)
    {
      object[] notation;

      if (!this._notationInfo.ContainsKey(nname))
      {
        notation = new object[2];
        notation[0] = pubid;
        notation[1] = sysid;
        this._notationInfo.Add(nname, notation);
      }
      else
      {
        // VC: Unique Notation Name
        // (it's not fatal)
      }
    }


    //
    // Location.
    //


    /**
     * Return the current line number.
     */

    public int LineNumber
    {
      get
      {
        return this._line;
      }
    }


    /**
     * Return the current column number.
     */

    public int ColumnNumber
    {
      get
      {
        return this._column;
      }
    }


    //////////////////////////////////////////////////////////////////////
    // High-level I/O.
    //////////////////////////////////////////////////////////////////////


    /**
     * Read a single character from the readBuffer.
     * <p>The readDataChunk () method maintains the buffer.
     * <p>If we hit the end of an entity, try to pop the stack and
     * keep going.
     * <p> (This approach doesn't really enforce XML's rules about
     * entity boundaries, but this is not currently a validating
     * parser).
     * <p>This routine also attempts to keep track of the current
     * position in external entities, but it's not entirely accurate.
     * @return The next available input character.
     * @see #unread (char)
     * @see #unread (string)
     * @see #readDataChunk
     * @see #readBuffer
     * @see #line
     * @return The next character from the current input source.
     */

    private char readCh()
    {

      // As long as there's nothing in the
      // read buffer, try reading more data
      // (for an external entity) or popping
      // the entity stack (for either).
      while (this._readBufferPos >= this._readBufferLength)
      {
        switch (this._sourceType)
        {
          case INPUT_READER:
          case INPUT_STREAM:
            readDataChunk();
            while (this._readBufferLength < 1)
            {
              popInput();
              if (this._readBufferLength < 1)
              {
                readDataChunk();
              }
            }
            break;

          default:

            popInput();
            break;
        }
      }

      char c = this._readBuffer[this._readBufferPos++];

      if (c == '\n')
      {
        this._line++;
        this._column = 0;
      }
      else
      {
        if (c == '<')
        {
          /* the most common  return to parseContent () .. NOP */
          ;
        }
        else if ((c < 0x0020 && (c != '\t') && (c != '\r')) || c > 0xFFFD) error("illegal XML character U+" + ((int)c).ToString("X"));

          // If we're in the DTD and in a context where PEs get expanded,
          // do so ... 1/14/2000 errata identify those contexts.  There
          // are also spots in the internal subset where PE refs are fatal
          // errors, hence yet another flag.
        else if (c == '%' && this._expandPE)
        {
          if (this._peIsError && this._entityStack.Count == 1) // not an error if PE reference is in an external PE called from internal subset
            error("PE reference within declaration in internal subset.");
          parsePEReference();
          return readCh();
        }
        this._column++;
      }

      return c;
    }


    /**
     * Push a single character back onto the current input stream.
     * <p>This method usually pushes the character back onto
     * the readBuffer, while the unread (string) method treats the
     * string as a new internal entity.
     * <p>I don't think that this would ever be called with 
     * readBufferPos = 0, because the methods always reads a character
     * before unreading it, but just in case, I've added a boundary
     * condition.
     * @param c The character to push back.
     * @see #readCh
     * @see #unread (string)
     * @see #unread (char[])
     * @see #readBuffer
     */

    private void unread(char c)
    {
      // Normal condition.
      if (c == '\n')
      {
        this._line--;
        this._column = -1;
      }
      if (this._readBufferPos > 0)
      {
        this._readBuffer[--this._readBufferPos] = c;
      }
      else
      {
        pushString(null, c.ToString(CultureInfo.InvariantCulture));
      }
    }


    /**
     * Push a char array back onto the current input stream.
     * <p>NOTE: you must <em>never</em> push back characters that you
     * haven't actually read: use pushString () instead.
     * @see #readCh
     * @see #unread (char)
     * @see #unread (string)
     * @see #readBuffer
     * @see #pushString
     */

    private void unread(char[] ch, int length)
    {
      for (int i = 0; i < length; i++)
      {
        if (ch[i] == '\n')
        {
          this._line--;
          this._column = -1;
        }
      }
      if (length < this._readBufferPos)
      {
        this._readBufferPos -= length;
      }
      else
      {
        pushCharArray(null, ch, 0, length);
        this._sourceType = INPUT_BUFFER;
      }
    }


    /**
     * Push a new external input source.
     * The source will be some kind of parsed entity, such as a PE
     * (including the external DTD subset) or content for the body.
     * <p>TODO: Right now, this method always attempts to autodetect
     * the encoding; in the future, it should allow the caller to 
     * request an encoding explicitly, and it should also look at the
     * headers with an HTTP connection.
     * @param url The java.net.URL object for the entity.
     * @see SAXDriver#resolveEntity
     * @see #pushString
     * @see #sourceType
     * @see #pushInput
     * @see #detectEncoding
     * @see #sourceType
     * @see #readBuffer
     */

    private void pushURL(
      string ename,
      string publicId,
      string systemId,
      TextReader reader,
      Stream stream,
      Encoding encoding,
      bool isAbsolute) {
      string encodingName = encoding == null ? null : encoding.WebName;
      bool ignoreEncoding = false;

      // Push the existing status.
      pushInput(ename);

      // Create a new read buffer.
      // (Note the four-character margin)
      this._readBuffer = new char[READ_BUFFER_MAX + 4];
      this._readBufferPos = 0;
      this._readBufferLength = 0;
      this._readBufferOverflow = -1;
      this._stream = null;
      this._line = 1;
      this._column = 0;
      this._currentByteCount = 0;

      if (!isAbsolute)
      {

        // Make any system ID (URI/URL) absolute.  There's one case
        // where it may be null:  parser was invoked without providing
        // one, e.g. since the XML data came from a memory buffer.
        try
        {
          if (systemId != null && this._externalEntity != null)
          {
            systemId = new Uri(this._externalEntity.RequestUri, systemId).ToString();
          }
          else if (this._baseURI != null)
          {
            systemId = new Uri(new Uri(this._baseURI), systemId).ToString();
            // throws IOException if couldn't create new URL
          }
        }
        catch (IOException err)
        {
          popInput();
          error("Invalid URL " + systemId + " (" + err.Message + ")");
        }
      }

      // See if the application wants to
      // redirect the system ID and/or
      // supply its own character stream.
      if (reader == null && stream == null && systemId != null)
      {
        object input = null;
        try
        {
          input = this._handler.ResolveEntity(publicId, systemId);
        }
        catch (IOException err)
        {
          popInput();
          error("Failure resolving entity " + systemId + " (" + err.Message + ")");
        }
        if (input != null)
        {
          if (input is string)
          {
            systemId = (string)input;
            isAbsolute = true;
          }
          else if (input is Stream)
          {
            stream = (Stream)input;
          }
          else if (input is TextReader)
          {
            reader = (TextReader)input;
          }
        }
      }

      // Start the entity.
      if (systemId != null)
      {
        this._handler.StartExternalEntity(systemId);
      }
      else
      {
        this._handler.StartExternalEntity("[unidentified data stream]");
      }

      // If there's an explicit character stream, just
      // ignore encoding declarations.
      if (reader != null)
      {
        this._sourceType = INPUT_READER;
        this._reader = reader;
        tryEncodingDecl(true);
        return;
      }

      // Else we handle the conversion, and need to ensure
      // it's done right.
      this._sourceType = INPUT_STREAM;
      if (stream != null)
      {
        this._stream = stream;
      }
      else
      {
        // We have to open our own stream to the URL.
        WebRequest request = WebRequest.CreateDefault(new Uri(systemId));
        request.Method = "GET";
        try
        {
          this._externalEntity = request;
          this._stream = this._externalEntity.GetResponse().GetResponseStream();
        }
        catch (IOException err)
        {
          try
          {
            popInput();
          }
          catch (Exception err2)
          {
          }
          error("Cannot read input file " + err.Message);
        }
      }

      // If we get to here, there must be
      // an InputStream available.
      if (!this._stream.CanSeek)
      {
        // TODO: this could be bad?
        var memoryStream = new MemoryStream();
        _stream.CopyTo(memoryStream);
        _stream = memoryStream;
      }

      // Get any external encoding label.
      if (encodingName == null && this._externalEntity != null)
      {
        // External labels can be untrustworthy; filesystems in
        // particular often have the wrong default for content
        // that wasn't locally originated.  Those we autodetect.
        if (!"file".Equals(this._externalEntity.RequestUri.Scheme))
        {
          int temp;

          // application/xml;charset=something;otherAttr=...
          // ... with many variants on 'something'
          encodingName = this._externalEntity.ContentType;

          // MHK code (fix for Saxon 5.5.1/007): protect against encoding==null
          if (encodingName == null)
          {
            temp = -1;
          }
          else
          {
            temp = encodingName.IndexOf("charset");
          }

          // RFC 2376 sez MIME text defaults to ASCII, but since the
          // JDK will create a MIME type out of thin air, we always
          // autodetect when there's no explicit charset attribute.
          if (temp < 0) encodingName = null; // autodetect
          else
          {
            temp = encodingName.IndexOf('=', temp + 7);
            encodingName = encodingName.Substring(temp + 1); // +1 added by MHK 2 April 2001
            if ((temp = encodingName.IndexOf(';')) > 0) encodingName = encodingName.Substring(0, temp);

            // attributes can have comment fields (RFC 822)
            if ((temp = encodingName.IndexOf('(')) > 0) encodingName = encodingName.Substring(0, temp);
            // ... and values may be quoted
            if ((temp = encodingName.IndexOf('"')) > 0) encodingName = encodingName.Substring(temp + 1, encodingName.IndexOf('"', temp + 2));
            encodingName.Trim();
          }
        }
      }

      // if we got an external encoding label, use it ...
      if (encodingName != null)
      {
        this._encoding = ENCODING_EXTERNAL;
        setupDecoding(encodingName);
        ignoreEncoding = true;

        // ... else autodetect
      }
      else
      {
        detectEncoding();
        ignoreEncoding = false;
      }
      var position = 100;

      // Read any XML or text declaration.
      try
      {
        tryEncodingDecl(ignoreEncoding);
      }
      catch (EncodingException x)
      {
        encodingName = x.Message;

        // if we don't handle the declared encoding,
        // try letting a JVM InputStreamReader do it
        try
        {
          if (this._sourceType != INPUT_STREAM) throw;

          this._stream.Position = position;
          this._readBufferPos = 0;
          this._readBufferLength = 0;
          this._readBufferOverflow = -1;
          this._line = 1;
          this._currentByteCount = this._column = 0;

          this._sourceType = INPUT_READER;
          this._reader = new StreamReader(this._stream, Encoding.GetEncoding(encodingName));
          this._stream = null;

          tryEncodingDecl(true);

        }
        catch (IOException e)
        {
          error("unsupported text encoding", encodingName, null);
        }
      }
    }


    /**
     * Check for an encoding declaration.  This is the second part of the
     * XML encoding autodetection algorithm, relying on detectEncoding to
     * get to the point that this part can read any encoding declaration
     * in the document (using only US-ASCII characters).
     *
     * <p> Because this part starts to fill parser buffers with this data,
     * it's tricky to to a reader so that Java's built-in decoders can be
     * used for the character encodings that aren't built in to this parser
     * (such as EUC-JP, KOI8-R, Big5, etc).
     *
     * @return any encoding in the declaration, uppercased; or null
     * @see detectEncoding
     */

    private string tryEncodingDecl(bool ignoreEncoding)
    {
      // Read the XML/text declaration.
      if (tryRead("<?xml"))
      {
        dataBufferFlush();
        if (tryWhitespace())
        {
          if (this._inputStack.Count > 0)
          {
            return parseTextDecl(ignoreEncoding);
          }
          else
          {
            return parseXMLDecl(ignoreEncoding);
          }
        }
        else
        {
          unread("xml".ToCharArray(), 3);
          parsePI();
        }
      }
      return null;
    }


    /**
     * Attempt to detect the encoding of an entity.
     * <p>The trick here (as suggested in the XML standard) is that
     * any entity not in UTF-8, or in UCS-2 with a byte-order mark, 
     * <b>must</b> begin with an XML declaration or an encoding
     * declaration; we simply have to look for "&lt;?xml" in various
     * encodings.
     * <p>This method has no way to distinguish among 8-bit encodings.
     * Instead, it sets up for UTF-8, then (possibly) revises its assumption
     * later in setupDecoding ().  Any ASCII-derived 8-bit encoding
     * should work, but most will be rejected later by setupDecoding ().
     * <p>I don't currently detect EBCDIC, since I'm concerned that it
     * could also be a valid UTF-8 sequence; I'll have to do more checking
     * later.
     * <p>MHK Nov 2001: modified to handle a BOM on UTF-8 files, which is
     * allowed by XML 2nd edition, and generated when Windows Notepad does
     * "save as UTF-8".
     * @see #tryEncoding (byte[], byte, byte, byte, byte)
     * @see #tryEncoding (byte[], byte, byte)
     * @see #setupDecoding
     * @see #read8bitEncodingDeclaration
     */

    private void detectEncoding()
    {
      byte[] signature = new byte[4];

      // Read the first four bytes for
      // autodetection.
      this._stream.Read(signature, 0, 4);
      this._stream.Position = 0;

      //
      // FIRST:  four byte encodings (who uses these?)
      //
      if (tryEncoding(signature, (byte)0x00, (byte)0x00, (byte)0x00, (byte)0x3c))
      {
        // UCS-4 must begin with "<?xml"
        // 0x00 0x00 0x00 0x3c: UCS-4, big-endian (1234)
        this._encoding = ENCODING_UCS_4_1234;

      }
      else if (tryEncoding(signature, (byte)0x3c, (byte)0x00, (byte)0x00, (byte)0x00))
      {
        // 0x3c 0x00 0x00 0x00: UCS-4, little-endian (4321)
        this._encoding = ENCODING_UCS_4_4321;

      }
      else if (tryEncoding(signature, (byte)0x00, (byte)0x00, (byte)0x3c, (byte)0x00))
      {
        // 0x00 0x00 0x3c 0x00: UCS-4, unusual (2143)
        this._encoding = ENCODING_UCS_4_2143;

      }
      else if (tryEncoding(signature, (byte)0x00, (byte)0x3c, (byte)0x00, (byte)0x00))
      {
        // 0x00 0x3c 0x00 0x00: UCS-4, unusual (3421)
        this._encoding = ENCODING_UCS_4_3412;

        // 00 00 fe ff UCS_4_1234 (with BOM)
        // ff fe 00 00 UCS_4_4321 (with BOM)
      }
    
        // SECOND: three byte signature:
        // look for UTF-8 byte order mark 3C 3F 78, allowed by XML 1.0 2nd edition

      else if (tryEncoding(signature, (byte)0xef, (byte)0xbb, (byte)0xbf))
      {
        this._encoding = ENCODING_UTF_8;
        this._stream.ReadByte();
        this._stream.ReadByte();
        this._stream.ReadByte();
      }

        //
        // THIRD:  two byte encodings
        // note ... with 1/14/2000 errata the XML spec identifies some
        // more "broken UTF-16" autodetection cases, with no XML decl,
        // which we don't handle here (that's legal too).
        //
      else if (tryEncoding(signature, (byte)0xfe, (byte)0xff))
      {
        // UCS-2 with a byte-order marker. (UTF-16)
        // 0xfe 0xff: UCS-2, big-endian (12)
        this._encoding = ENCODING_UCS_2_12;
        this._stream.ReadByte();
        this._stream.ReadByte();

      }
      else if (tryEncoding(signature, (byte)0xff, (byte)0xfe))
      {
        // UCS-2 with a byte-order marker. (UTF-16)
        // 0xff 0xfe: UCS-2, little-endian (21)
        this._encoding = ENCODING_UCS_2_21;
        this._stream.ReadByte();
        this._stream.ReadByte();

      }
      else if (tryEncoding(signature, (byte)0x00, (byte)0x3c, (byte)0x00, (byte)0x3f))
      {
        // UTF-16-BE (otherwise, malformed UTF-16)
        // 0x00 0x3c 0x00 0x3f: UCS-2, big-endian, no byte-order mark
        this._encoding = ENCODING_UCS_2_12;
        error("no byte-order mark for UCS-2 entity");

      }
      else if (tryEncoding(signature, (byte)0x3c, (byte)0x00, (byte)0x3f, (byte)0x00))
      {
        // UTF-16-LE (otherwise, malformed UTF-16)
        // 0x3c 0x00 0x3f 0x00: UCS-2, little-endian, no byte-order mark
        this._encoding = ENCODING_UCS_2_21;
        error("no byte-order mark for UCS-2 entity");
      }

        //
        // THIRD:  ASCII-derived encodings, fixed and variable lengths
        //
      else if (tryEncoding(signature, (byte)0x3c, (byte)0x3f, (byte)0x78, (byte)0x6d))
      {
        // ASCII derived
        // 0x3c 0x3f 0x78 0x6d: UTF-8 or other 8-bit markup (read ENCODING)
        this._encoding = ENCODING_UTF_8;
        read8bitEncodingDeclaration();

      }
      else
      {
        // 4c 6f a7 94 ... we don't understand EBCDIC flavors
        // ... but we COULD at least kick in some fixed code page

        // (default) UTF-8 without encoding/XML declaration
        this._encoding = ENCODING_UTF_8;
      }
    }


    /**
     * Check for a four-byte signature.
     * <p>Utility routine for detectEncoding ().
     * <p>Always looks for some part of "<?XML" in a specific encoding.
     * @param sig The first four bytes read.
     * @param b1 The first byte of the signature
     * @param b2 The second byte of the signature
     * @param b3 The third byte of the signature
     * @param b4 The fourth byte of the signature
     * @see #detectEncoding
     */

    private static bool tryEncoding(byte[] sig, byte b1, byte b2, byte b3, byte b4)
    {
      return (sig[0] == b1 && sig[1] == b2 && sig[2] == b3 && sig[3] == b4);
    }


    /**
     * Check for a two-byte signature.
     * <p>Looks for a UCS-2 byte-order mark.
     * <p>Utility routine for detectEncoding ().
     * @param sig The first four bytes read.
     * @param b1 The first byte of the signature
     * @param b2 The second byte of the signature
     * @see #detectEncoding
     */

    private static bool tryEncoding(byte[] sig, byte b1, byte b2)
    {
      return ((sig[0] == b1) && (sig[1] == b2));
    }

    /**
     * Check for a three-byte signature.
     * <p>Looks for a UTF-8 byte-order mark.
     * <p>Utility routine for detectEncoding ().
     * @param sig The first four bytes read.
     * @param b1 The first byte of the signature
     * @param b2 The second byte of the signature
     * @param b3 The second byte of the signature
     * @see #detectEncoding
     */

    private static bool tryEncoding(byte[] sig, byte b1, byte b2, byte b3)
    {
      return ((sig[0] == b1) && (sig[1] == b2) && (sig[2] == b3));
    }

    /**
     * This method pushes a string back onto input.
     * <p>It is useful either as the expansion of an internal entity, 
     * or for backtracking during the parse.
     * <p>Call pushCharArray () to do the actual work.
     * @param s The string to push back onto input.
     * @see #pushCharArray
     */

    private void pushString(string ename, string s)
    {
      char[] ch = s.ToCharArray();
      pushCharArray(ename, ch, 0, ch.Length);
    }


    /**
     * Push a new internal input source.
     * <p>This method is useful for expanding an internal entity,
     * or for unreading a string of characters.  It creates a new
     * readBuffer containing the characters in the array, instead
     * of characters converted from an input byte stream.
     * @param ch The char array to push.
     * @see #pushString
     * @see #pushURL
     * @see #readBuffer
     * @see #sourceType
     * @see #pushInput
     */

    private void pushCharArray(string ename, char[] ch, int start, int length)
    {
      // Push the existing status
      pushInput(ename);
      this._sourceType = INPUT_INTERNAL;
      this._readBuffer = ch;
      this._readBufferPos = start;
      this._readBufferLength = length;
      this._readBufferOverflow = -1;
    }


    /**
     * Save the current input source onto the stack.
     * <p>This method saves all of the global variables associated with
     * the current input source, so that they can be restored when a new
     * input source has finished.  It also tests for entity recursion.
     * <p>The method saves the following global variables onto a stack
     * using a fixed-length array:
     * <ol>
     * <li>sourceType
     * <li>externalEntity
     * <li>readBuffer
     * <li>readBufferPos
     * <li>readBufferLength
     * <li>line
     * <li>encoding
     * </ol>
     * @param ename The name of the entity (if any) causing the new input.
     * @see #popInput
     * @see #sourceType
     * @see #externalEntity
     * @see #readBuffer
     * @see #readBufferPos
     * @see #readBufferLength
     * @see #line
     * @see #encoding
     */

    private void pushInput(string ename)
    {
      object[] input = new object[12];

      // Check for entity recursion.
      if (ename != null)
      {
        IEnumerator entities = this._entityStack.GetEnumerator();
        while (entities.MoveNext())
        {
          string e = (string)entities.Current;
          if (e == ename)
          {
            error("recursive reference to entity", ename, null);
          }
        }
      }
      this._entityStack.Push(ename);

      // Don't bother if there is no current input.
      if (this._sourceType == INPUT_NONE)
      {
        return;
      }

      // Set up a snapshot of the current
      // input source.
      input[0] = this._sourceType;
      input[1] = this._externalEntity;
      input[2] = this._readBuffer;
      input[3] = this._readBufferPos;
      input[4] = this._readBufferLength;
      input[5] = this._line;
      input[6] = this._encoding;
      input[7] = this._readBufferOverflow;
      input[8] = this._stream
      ;
      input[9] = this._currentByteCount;
      input[10] = this._column;
      input[11] = this._reader;

      // Push it onto the stack.
      this._inputStack.Push(input);
    }


    /**
     * Restore a previous input source.
     * <p>This method restores all of the global variables associated with
     * the current input source.
     * @exception java.io.EOFException
     *    If there are no more entries on the input stack.
     * @see #pushInput
     * @see #sourceType
     * @see #externalEntity
     * @see #readBuffer
     * @see #readBufferPos
     * @see #readBufferLength
     * @see #line
     * @see #encoding
     */

    private void popInput()
    {
      string uri;

      if (this._externalEntity != null) uri = this._externalEntity.RequestUri.ToString();
      else uri = this._baseURI;

      switch (this._sourceType)
      {
        case INPUT_STREAM:
          if (this._stream != null)
          {
            if (uri != null)
            {
              this._handler.EndExternalEntity(this._baseURI);
            }
            this._stream.Close();
          }
          break;
        case INPUT_READER:
          if (this._reader != null)
          {
            if (uri != null)
            {
              this._handler.EndExternalEntity(this._baseURI);
            }
            this._reader.Close();
          }
          break;
      }

      // Throw an EOFException if there
      // is nothing else to pop.
      if (this._inputStack.Count == 0)
      {
        throw new EndOfStreamException("no more input");
      }

      object[] input = (object[])this._inputStack.Pop();
      this._entityStack.Pop();

      this._sourceType = ((int)input[0]);
      this._externalEntity = (WebRequest)input[1];
      this._readBuffer = (char[])input[2];
      this._readBufferPos = ((int)input[3]);
      this._readBufferLength = ((int)input[4]);
      this._line = ((int)input[5]);
      this._encoding = ((int)input[6]);
      this._readBufferOverflow = ((int)input[7]);
      this._stream = (Stream)input[8];
      this._currentByteCount = ((int)input[9]);
      this._column = ((int)input[10]);
      this._reader = (TextReader)input[11];
    }


    /**
     * Return true if we can read the expected character.
     * <p>Note that the character will be removed from the input stream
     * on success, but will be put back on failure.  Do not attempt to
     * read the character again if the method succeeds.
     * @param delim The character that should appear next.  For a
     *	      insensitive match, you must supply this in upper-case.
     * @return true if the character was successfully read, or false if
     *	 it was not.
     * @see #tryRead (string)
     */

    private bool tryRead(char delim)
    {
      char c;

      // Read the character
      c = readCh();

      // Test for a match, and push the character
      // back if the match fails.
      if (c == delim)
      {
        return true;
      }
      else
      {
        unread(c);
        return false;
      }
    }


    /**
     * Return true if we can read the expected string.
     * <p>This is simply a convenience method.
     * <p>Note that the string will be removed from the input stream
     * on success, but will be put back on failure.  Do not attempt to
     * read the string again if the method succeeds.
     * <p>This method will push back a character rather than an
     * array whenever possible (probably the majority of cases).
     * <p><b>NOTE:</b> This method currently has a hard-coded limit
     * of 100 characters for the delimiter.
     * @param delim The string that should appear next.
     * @return true if the string was successfully read, or false if
     *	 it was not.
     * @see #tryRead (char)
     */

    private bool tryRead(string delim)
    {
      char[] ch = delim.ToCharArray();
      char c;

      // Compare the input, character-
      // by character.

      for (int i = 0; i < ch.Length; i++)
      {
        c = readCh();
        if (c != ch[i])
        {
          unread(c);
          if (i != 0)
          {
            unread(ch, i);
          }
          return false;
        }
      }
      return true;
    }



    /**
     * Return true if we can read some whitespace.
     * <p>This is simply a convenience method.
     * <p>This method will push back a character rather than an
     * array whenever possible (probably the majority of cases).
     * @return true if whitespace was found.
     */

    private bool tryWhitespace()
    {
      char c;
      c = readCh();
      if (isWhitespace(c))
      {
        skipWhitespace();
        return true;
      }
      else
      {
        unread(c);
        return false;
      }
    }


    /**
     * Read all data until we find the specified string.
     * This is useful for scanning CDATA sections and PIs.
     * <p>This is inefficient right now, since it calls tryRead ()
     * for every character.
     * @param delim The string delimiter
     * @see #tryRead (string, bool)
     * @see #readCh
     */

    private void parseUntil(string delim)
    {
      char c;
      int startLine = this._line;

      try
      {
        while (!tryRead(delim))
        {
          c = readCh();
          dataBufferAppend(c);
        }
      }
      catch (EndOfStreamException e)
      {
        error("end of input while looking for delimiter " + "(started on line " + startLine + ')', null, delim);
      }
    }


    /**
     * Read just the encoding declaration (or XML declaration) at the 
     * start of an external entity.
     * When this method is called, we know that the declaration is
     * present (or appears to be).  We also know that the entity is
     * in some sort of ASCII-derived 8-bit encoding.
     * The idea of this is to let us read what the 8-bit encoding is
     * before we've committed to converting any more of the file; the
     * XML or encoding declaration must be in 7-bit ASCII, so we're
     * safe as long as we don't go past it.
     */

    private void read8bitEncodingDeclaration()
    {
      int ch;
      this._readBufferPos = this._readBufferLength = 0;

      while (true)
      {
        ch = this._stream.ReadByte();
        this._readBuffer[this._readBufferLength++] = (char)ch;
        switch (ch)
        {
          case (int)'>':
            return;
          case -1:
            error("end of file before end of XML or encoding declaration.", null, "?>");
            break;
        }
        if (this._readBuffer.Length == this._readBufferLength) error("unfinished XML or encoding declaration");
      }
    }


    //////////////////////////////////////////////////////////////////////
    // Low-level I/O.
    //////////////////////////////////////////////////////////////////////


    /**
     * Read a chunk of data from an external input source.
     * <p>This is simply a front-end that fills the rawReadBuffer
     * with bytes, then calls the appropriate encoding handler.
     * @see #encoding
     * @see #rawReadBuffer
     * @see #readBuffer
     * @see #filterCR
     * @see #copyUtf8ReadBuffer
     * @see #copyIso8859_1ReadBuffer
     * @see #copyUcs_2ReadBuffer
     * @see #copyUcs_4ReadBuffer
     */

    private void readDataChunk()
    {
      int count;

      // See if we have any overflow (filterCR sets for CR at end)
      if (this._readBufferOverflow > -1)
      {
        this._readBuffer[0] = (char)this._readBufferOverflow;
        this._readBufferOverflow = -1;
        this._readBufferPos = 1;
        this._sawCR = true;
      }
      else
      {
        this._readBufferPos = 0;
        this._sawCR = false;
      }

      // input from a character stream.
      if (this._sourceType == INPUT_READER)
      {
        count = this._reader.Read(this._readBuffer, this._readBufferPos, READ_BUFFER_MAX - this._readBufferPos);
        if (count < 0) this._readBufferLength = this._readBufferPos;
        else this._readBufferLength = this._readBufferPos + count;
        if (this._readBufferLength > 0) filterCR(count >= 0);
        this._sawCR = false;
        return;
      }

      // Read as many bytes as possible into the raw buffer.
      count = this._stream.Read(this._rawReadBuffer, 0, READ_BUFFER_MAX);

      // Dispatch to an encoding-specific reader method to populate
      // the readBuffer.  In most parser speed profiles, these routines
      // show up at the top of the CPU usage chart.
      if (count > 0)
      {
        switch (this._encoding)
        {
            // one byte builtins
          case ENCODING_ASCII:
            copyIso8859_1ReadBuffer(count, (char)0x0080);
            break;
          case ENCODING_UTF_8:
            copyUtf8ReadBuffer(count);
            break;
          case ENCODING_ISO_8859_1:
            copyIso8859_1ReadBuffer(count, (char)0);
            break;

            // two byte builtins
          case ENCODING_UCS_2_12:
            copyUcs2ReadBuffer(count, 8, 0);
            break;
          case ENCODING_UCS_2_21:
            copyUcs2ReadBuffer(count, 0, 8);
            break;

            // four byte builtins
          case ENCODING_UCS_4_1234:
            copyUcs4ReadBuffer(count, 24, 16, 8, 0);
            break;
          case ENCODING_UCS_4_4321:
            copyUcs4ReadBuffer(count, 0, 8, 16, 24);
            break;
          case ENCODING_UCS_4_2143:
            copyUcs4ReadBuffer(count, 16, 24, 0, 8);
            break;
          case ENCODING_UCS_4_3412:
            copyUcs4ReadBuffer(count, 8, 0, 24, 16);
            break;
        }
      }
      else this._readBufferLength = this._readBufferPos;

      this._readBufferPos = 0;

      // Filter out all carriage returns if we've seen any
      // (including any saved from a previous read)
      if (this._sawCR)
      {
        filterCR(count >= 0);
        this._sawCR = false;

        // must actively report EOF, lest some CRs get lost.
        if (this._readBufferLength == 0 && count >= 0) readDataChunk();
      }

      if (count > 0) this._currentByteCount += count;
    }


    /**
     * Filter carriage returns in the read buffer.
     * CRLF becomes LF; CR becomes LF.
     * @param moreData true iff more data might come from the same source
     * @see #readDataChunk
     * @see #readBuffer
     * @see #readBufferOverflow
     */

    private void filterCR(bool moreData)
    {
      int i, j;

      this._readBufferOverflow = -1;

      for (i = j = this._readBufferPos; j < this._readBufferLength; i++, j++)
      {
        switch (this._readBuffer[j])
        {
          case '\r':
            if (j == this._readBufferLength - 1)
            {
              if (moreData)
              {
                this._readBufferOverflow = '\r';
                this._readBufferLength--;
              }
              else // CR at end of buffer
                this._readBuffer[i++] = '\n';
              break;
            }
            else if (this._readBuffer[j + 1] == '\n')
            {
              j++;
            }
            this._readBuffer[i] = '\n';
            break;

          case '\n':
          default:
            this._readBuffer[i] = this._readBuffer[j];
            break;
        }
      }
      this._readBufferLength = i;
    }

    /**
     * Convert a buffer of UTF-8-encoded bytes into UTF-16 characters.
     * <p>When readDataChunk () calls this method, the raw bytes are in 
     * rawReadBuffer, and the readonly characters will appear in 
     * readBuffer.
     * @param count The number of bytes to convert.
     * @see #readDataChunk
     * @see #rawReadBuffer
     * @see #readBuffer
     * @see #getNextUtf8Byte
     */

    private void copyUtf8ReadBuffer(int count)
    {
      int i = 0;
      int j = this._readBufferPos;
      byte b1;
      char c;

      /*
    // check once, so the runtime won't (if it's smart enough)
    if (count < 0 || count > rawReadBuffer.length)
        throw new ArrayIndexOutOfBoundsException (Integer.toString (count));
    */

      while (i < count)
      {
        b1 = this._rawReadBuffer[i++];

        // Determine whether we are dealing
        // with a one-, two-, three-, or four-
        // byte sequence.
        if ((b1 & 0x80) == 0x80)
        {
          if ((b1 & 0xe0) == 0xc0)
          {
            // 2-byte sequence: 00000yyyyyxxxxxx = 110yyyyy 10xxxxxx
            c = (char)(((b1 & 0x1f) << 6) | getNextUtf8Byte(i++, count));
          }
          else if ((b1 & 0xf0) == 0xe0)
          {
            // 3-byte sequence:
            // zzzzyyyyyyxxxxxx = 1110zzzz 10yyyyyy 10xxxxxx
            // most CJKV characters
            c = (char)(((b1 & 0x0f) << 12) | (getNextUtf8Byte(i++, count) << 6) | getNextUtf8Byte(i++, count));
          }
          else if ((b1 & 0xf8) == 0xf0)
          {
            // 4-byte sequence: 11101110wwwwzzzzyy + 110111yyyyxxxxxx
            //     = 11110uuu 10uuzzzz 10yyyyyy 10xxxxxx
            // (uuuuu = wwww + 1)
            // "Surrogate Pairs" ... from the "Astral Planes"
            int iso646 = b1 & 07;
            iso646 = (iso646 << 6) + getNextUtf8Byte(i++, count);
            iso646 = (iso646 << 6) + getNextUtf8Byte(i++, count);
            iso646 = (iso646 << 6) + getNextUtf8Byte(i++, count);

            if (iso646 <= 0xffff)
            {
              c = (char)iso646;
            }
            else
            {
              if (iso646 > 0x0010ffff) encodingError("UTF-8 value out of range for Unicode", iso646, 0);
              iso646 -= 0x010000;
              this._readBuffer[j++] = (char)(0xd800 | (iso646 >> 10));
              this._readBuffer[j++] = (char)(0xdc00 | (iso646 & 0x03ff));
              continue;
            }
          }
          else
          {
            // The five and six byte encodings aren't supported;
            // they exceed the Unicode (and XML) range.
            encodingError("invalid UTF-8 byte (check the XML declaration)", 0xff & b1, i);
            // NOTREACHED
            c = (char)0;
          }
        }
        else
        {
          // 1-byte sequence: 000000000xxxxxxx = 0xxxxxxx
          // (US-ASCII character, "common" case, one branch to here)
          c = (char)b1;
        }
        this._readBuffer[j++] = c;
        if (c == '\r') this._sawCR = true;
      }
      // How many characters have we read?
      this._readBufferLength = j;
    }


    /**
     * Return the next byte value in a UTF-8 sequence.
     * If it is not possible to get a byte from the current
     * entity, throw an exception.
     * @param pos The current position in the rawReadBuffer.
     * @param count The number of bytes in the rawReadBuffer
     * @return The significant six bits of a non-initial byte in
     *	 a UTF-8 sequence.
     * @exception EOFException If the sequence is incomplete.
     */

    private int getNextUtf8Byte(int pos, int count)
    {
      int val;

      // Take a character from the buffer
      // or from the actual input stream.
      if (pos < count)
      {
        val = this._rawReadBuffer[pos];
      }
      else
      {
        val = this._stream.ReadByte();
        if (val == -1)
        {
          encodingError("unfinished multi-byte UTF-8 sequence at EOF", -1, pos);
        }
      }

      // Check for the correct bits at the start.
      if ((val & 0xc0) != 0x80)
      {
        encodingError("bad continuation of multi-byte UTF-8 sequence", val, pos + 1);
      }

      // Return the significant bits.
      return (val & 0x3f);
    }


    /**
     * Convert a buffer of US-ASCII or ISO-8859-1-encoded bytes into
     * UTF-16 characters.
     *
     * <p>When readDataChunk () calls this method, the raw bytes are in 
     * rawReadBuffer, and the readonly characters will appear in 
     * readBuffer.
     *
     * @param count The number of bytes to convert.
     * @param mask For ASCII conversion, 0x7f; else, 0xff.
     * @see #readDataChunk
     * @see #rawReadBuffer
     * @see #readBuffer
     */

    private void copyIso8859_1ReadBuffer(int count, char mask)
    {
      int i, j;
      for (i = 0, j = this._readBufferPos; i < count; i++, j++)
      {
        char c = (char)(this._rawReadBuffer[i] & 0xff);
        if ((c & mask) != 0) throw new CharConversionException("non-ASCII character U+" + ((int)c).ToString("X"));
        this._readBuffer[j] = c;
        if (c == '\r')
        {
          this._sawCR = true;
        }
      }
      this._readBufferLength = j;
    }


    /**
     * Convert a buffer of UCS-2-encoded bytes into UTF-16 characters
     * (as used in Java string manipulation).
     *
     * <p>When readDataChunk () calls this method, the raw bytes are in 
     * rawReadBuffer, and the readonly characters will appear in 
     * readBuffer.
     * @param count The number of bytes to convert.
     * @param shift1 The number of bits to shift byte 1.
     * @param shift2 The number of bits to shift byte 2
     * @see #readDataChunk
     * @see #rawReadBuffer
     * @see #readBuffer
     */

    private void copyUcs2ReadBuffer(int count, int shift1, int shift2)
    {
      int j = this._readBufferPos;

      if (count > 0 && (count % 2) != 0)
      {
        encodingError("odd number of bytes in UCS-2 encoding", -1, count);
      }
      // The loops are faster with less internal brancing; hence two
      if (shift1 == 0)
      {
        // "UTF-16-LE"
        for (int i = 0; i < count; i += 2)
        {
          char c = (char)(this._rawReadBuffer[i + 1] << 8);
          c |= (char) (0xff & this._rawReadBuffer[i]);
          this._readBuffer[j++] = c;
          if (c == '\r') this._sawCR = true;
        }
      }
      else
      {
        // "UTF-16-BE"
        for (int i = 0; i < count; i += 2)
        {
          char c = (char)(this._rawReadBuffer[i] << 8);
          c |= (char)(0xff & this._rawReadBuffer[i + 1]);
          this._readBuffer[j++] = c;
          if (c == '\r') this._sawCR = true;
        }
      }
      this._readBufferLength = j;
    }


    /**
     * Convert a buffer of UCS-4-encoded bytes into UTF-16 characters.
     *
     * <p>When readDataChunk () calls this method, the raw bytes are in 
     * rawReadBuffer, and the readonly characters will appear in 
     * readBuffer.
     * <p>Java has Unicode chars, and this routine uses surrogate pairs
     * for ISO-10646 values between 0x00010000 and 0x000fffff.  An
     * exception is thrown if the ISO-10646 character has no Unicode
     * representation.
     *
     * @param count The number of bytes to convert.
     * @param shift1 The number of bits to shift byte 1.
     * @param shift2 The number of bits to shift byte 2
     * @param shift3 The number of bits to shift byte 2
     * @param shift4 The number of bits to shift byte 2
     * @see #readDataChunk
     * @see #rawReadBuffer
     * @see #readBuffer
     */

    private void copyUcs4ReadBuffer(int count, int shift1, int shift2, int shift3, int shift4)
    {
      int j = this._readBufferPos;

      if (count > 0 && (count % 4) != 0)
      {
        encodingError("number of bytes in UCS-4 encoding not divisible by 4", -1, count);
      }
      for (int i = 0; i < count; i += 4)
      {
        int value = (((this._rawReadBuffer[i] & 0xff) << shift1) | ((this._rawReadBuffer[i + 1] & 0xff) << shift2)
                     | ((this._rawReadBuffer[i + 2] & 0xff) << shift3) | ((this._rawReadBuffer[i + 3] & 0xff) << shift4));
        if (value < 0x0000ffff)
        {
          this._readBuffer[j++] = (char)value;
          if (value == (int)'\r')
          {
            this._sawCR = true;
          }
        }
        else if (value < 0x0010ffff)
        {
          value -= 0x010000;
          this._readBuffer[j++] = (char)(0xd8 | ((value >> 10) & 0x03ff));
          this._readBuffer[j++] = (char)(0xdc | (value & 0x03ff));
        }
        else
        {
          encodingError("UCS-4 value out of range for Unicode", value, i);
        }
      }
      this._readBufferLength = j;
    }


    /**
     * Report a character encoding error.
     */

    private void encodingError(string message, int value, int offset)
    {
      string uri;

      if (value != -1)
      {
        message = message + " (code: 0x" + value.ToString("X") + ')';
      }
      if (this._externalEntity != null)
      {
        uri = this._externalEntity.RequestUri.ToString();
      }
      else
      {
        uri = this._baseURI;
      }
      this._handler.Error(message, uri, -1, offset + this._currentByteCount);
    }


    //////////////////////////////////////////////////////////////////////
    // Local Variables.
    //////////////////////////////////////////////////////////////////////

    private bool isUnicodeIdentifierStart(char c)
    {
      //TODO: FIX THIS
      return c == '_' || char.IsLetter(c);
    }
    private bool isUnicodeIdentifierPart(char c)
    {
      //TODO: FIX THIS
      return char.IsPunctuation(c) || char.IsLetterOrDigit(c);
    }

    /**
     * Re-initialize the variables for each parse.
     */

    private void initializeVariables()
    {
      // First line
      this._line = 1;
      this._column = 0;

      // Set up the buffers for data and names
      this._dataBufferPos = 0;
      this._dataBuffer = new char[DATA_BUFFER_INITIAL];
      this._nameBufferPos = 0;
      this._nameBuffer = new char[NAME_BUFFER_INITIAL];

      // Set up the DTD hash tables
      this._elementInfo = new Hashtable();
      this._entityInfo = new Hashtable();
      this._notationInfo = new Hashtable();

      // Set up the variables for the current
      // element context.
      this._currentElement = null;
      this._currentElementContent = CONTENT_UNDECLARED;

      // Set up the input variables
      this._sourceType = INPUT_NONE;
      this._inputStack = new Stack();
      this._entityStack = new Stack();
      this._externalEntity = null;
      this._tagAttributePos = 0;
      this._tagAttributes = new string[100];
      this._rawReadBuffer = new byte[READ_BUFFER_MAX];
      this._readBufferOverflow = -1;

      this._inLiteral = false;
      this._expandPE = false;
      this._peIsError = false;

      this._inCDATA = false;

      this._symbolTable = new object[SYMBOL_TABLE_LENGTH][];
    }


    /**
     * Clean up after the parse to allow some garbage collection.
     */

    private void cleanupVariables()
    {
      this._dataBuffer = null;
      this._nameBuffer = null;

      this._elementInfo = null;
      this._entityInfo = null;
      this._notationInfo = null;

      this._currentElement = null;

      this._inputStack = null;
      this._entityStack = null;
      this._externalEntity = null;

      this._tagAttributes = null;
      this._rawReadBuffer = null;

      this._symbolTable = null;
    }

    /* used to restart reading with some InputStreamReader */

    private class EncodingException : IOException
    {
      public EncodingException(string encoding)
        : base(encoding)
      {
      }
    }

    //
    // The current XML handler interface.
    //
    private SAXDriver _handler;

    //
    // I/O information.
    //
    private TextReader _reader; // current reader

    private Stream _stream; // current input stream

    private int _line; // current line number

    private int _column; // current column number

    private int _sourceType; // type of input source

    private Stack _inputStack; // stack of input soruces

    private WebRequest _externalEntity; // current external entity

    private int _encoding; // current character encoding

    private int _currentByteCount; // bytes read from current source

    //
    // Buffers for decoded but unparsed character input.
    //
    private char[] _readBuffer;

    private int _readBufferPos;

    private int _readBufferLength;

    private int _readBufferOverflow; // overflow from last data chunk.


    //
    // Buffer for undecoded raw byte input.
    //
    private const int READ_BUFFER_MAX = 16384;

    private byte[] _rawReadBuffer;


    //
    // Buffer for parsed character data.
    //
    private const int DATA_BUFFER_INITIAL = 4096;

    private char[] _dataBuffer;

    private int _dataBufferPos;

    //
    // Buffer for parsed names.
    //
    private const int NAME_BUFFER_INITIAL = 1024;

    private char[] _nameBuffer;

    private int _nameBufferPos;


    //
    // Hashtables for DTD information on elements, entities, and notations.
    //
    private Hashtable _elementInfo;

    private Hashtable _entityInfo;

    private Hashtable _notationInfo;


    //
    // Element type currently in force.
    //
    private string _currentElement;

    private int _currentElementContent;

    //
    // Base external identifiers for resolution.
    //
    private string _basePublicId;

    private string _baseURI;

    private int _baseEncoding;

    private TextReader _baseReader;

    private Stream _baseInputStream;

    private char[] _baseInputBuffer ;

    private int _baseInputBufferStart;

    private int _baseInputBufferLength;

    //
    // Stack of entity names, to detect recursion.
    //
    private Stack _entityStack;

    //
    // PE expansion is enabled in most chunks of the DTD, not all.
    // When it's enabled, literals are treated differently.
    //
    private bool _inLiteral;

    private bool _expandPE;

    private bool _peIsError;

    //
    // Symbol table, for caching interned names.
    //
    private const int SYMBOL_TABLE_LENGTH = 1087;

    private object[][] _symbolTable;

    //
    // Hash table of attributes found in current start tag.
    //
    private string[] _tagAttributes;

    private int _tagAttributePos;

    //
    // Utility flag: have we noticed a CR while reading the last
    // data chunk?  If so, we will have to go back and normalise
    // CR or CR/LF line ends.
    //
    private bool _sawCR;

    //
    // Utility flag: are we in CDATA?  If so, whitespace isn't ignorable.
    // 
    private bool _inCDATA;
  }
}
