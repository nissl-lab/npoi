namespace NSAX.Ext {
  using System;

  using NSAX.Helpers;

  /// <summary>
  ///   SAX2 extension helper for additional Attributes information,
  ///   implementing the <see cref="Attributes2" /> interface.
  ///   <blockquote>
  ///     <em>
  ///       This module, both source code and documentation, is in the
  ///       Public Domain, and comes with <strong>NO WARRANTY</strong>.
  ///     </em>
  ///   </blockquote>
  ///   <para>This is not part of core-only SAX2 distributions.</para>
  ///   <para>
  ///     The <em>specified</em> flag for each attribute will always
  ///     be true, unless it has been set to false in the copy constructor
  ///     or using <see cref="SetSpecified" />.
  ///     Similarly, the <em>declared</em> flag for each attribute will
  ///     always be false, except for defaulted attributes (<em>specified</em>
  ///     is false), non-CDATA attributes, or when it is set to true using
  ///     <see cref="SetDeclared" />.
  ///     If you change an attribute's type by hand, you may need to modify
  ///     its <em>declared</em> flag to match.
  ///   </para>
  /// </summary>
  public class Attributes2 : Attributes, IAttributes2 {
    private bool[] _declared;
    private bool[] _specified;

    /// <summary>
    ///   Construct a new, empty Attributes2 object.
    /// </summary>
    public Attributes2() {
    }

    ////
    /// <summary>
    ///   Copy an existing Attributes or Attributes2 object.
    ///   If the object implements Attributes2, values of the
    ///   <em>specified</em> and <em>declared</em> flags for each
    ///   attribute are copied.
    ///   Otherwise the flag values are defaulted to assume no DTD was used,
    ///   unless there is evidence to the contrary (such as attributes with
    ///   type other than CDATA, which must have been <em>declared</em>).
    ///   <para>
    ///     This constructor is especially useful inside a
    ///     <see cref="IContentHandler.StartElement" /> event.
    ///   </para>
    /// </summary>
    /// <param name="atts">
    ///   The existing Attributes object.
    /// </param>
    public Attributes2(IAttributes atts) : base(atts) {
      SetAttributesInternal(atts);
    }

    public virtual bool IsDeclared(int index) {
      if (index < 0 || index >= Length) {
        throw new ArgumentOutOfRangeException("index", index, "no attribute at index");
      }
      return _declared[index];
    }

    public virtual bool IsDeclared(string uri, string localName) {
      int index = GetIndex(uri, localName);

      if (index < 0) {
        throw new ArgumentException("No such attribute: local=" + localName + ", namespace=" + uri);
      }
      return _declared[index];
    }

    public virtual bool IsDeclared(string qName) {
      int index = GetIndex(qName);

      if (index < 0) {
        throw new ArgumentException("No such attribute: " + qName);
      }
      return _declared[index];
    }

    public virtual bool IsSpecified(int index) {
      if (index < 0 || index >= Length) {
        throw new ArgumentOutOfRangeException("index", index, "No attribute at index");
      }
      return _specified[index];
    }

    public bool IsSpecified(string uri, string localName) {
      int index = GetIndex(uri, localName);

      if (index < 0) {
        throw new ArgumentException("No such attribute: local=" + localName + ", namespace=" + uri);
      }
      return _specified[index];
    }

    public virtual bool IsSpecified(string qName) {
      int index = GetIndex(qName);

      if (index < 0) {
        throw new ArgumentException("No such attribute: " + qName);
      }
      return _specified[index];
    }

    public override void SetAttributes(IAttributes atts) {
      SetAttributesInternal(atts);
    }

    public override void AddAttribute(string uri, string localName, string qName, string type, string value) {
      base.AddAttribute(uri, localName, qName, type, value);

      int length = Length;

      if (length < _specified.Length) {
        var newFlags = new bool[length];
        Array.Copy(_declared, 0, newFlags, 0, _declared.Length);
        _declared = newFlags;

        newFlags = new bool[length];
        Array.Copy(_specified, 0, newFlags, 0, _specified.Length);
        _specified = newFlags;
      }

      _specified[length - 1] = true;
      _declared[length - 1] = !"CDATA".Equals(type);
    }

    public override void RemoveAttribute(int index) {
      int origMax = Length - 1;

      base.RemoveAttribute(index);
      if (index != origMax) {
        Array.Copy(_declared, index + 1, _declared, index, origMax - index);
        Array.Copy(_specified, index + 1, _specified, index, origMax - index);
      }
    }

    public virtual void SetDeclared(int index, bool value) {
      if (index < 0 || index >= Length) {
        throw new ArgumentOutOfRangeException("index", index, "No attribute at index");
      }
      _declared[index] = value;
    }

    public virtual void SetSpecified(int index, bool value) {
      if (index < 0 || index >= Length) {
        throw new ArgumentOutOfRangeException("index", index, "No attribute at index");
      }
      _specified[index] = value;
    }

    private void SetAttributesInternal(IAttributes atts) {
      int length = atts.Length;

      base.SetAttributes(atts);
      _declared = new bool[length];
      _specified = new bool[length];

      var a2 = atts as IAttributes2;
      if (a2 != null) {
        for (int i = 0; i < length; i++) {
          _declared[i] = a2.IsDeclared(i);
          _specified[i] = a2.IsSpecified(i);
        }
      } else {
        for (int i = 0; i < length; i++) {
          _declared[i] = !"CDATA".Equals(atts.GetType(i));
          _specified[i] = true;
        }
      }
    }
  }
}
