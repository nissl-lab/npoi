namespace NSAX.Helpers {
  using System;

  /// <summary>
  ///   Default implementation of the Attributes interface.
  ///   <blockquote>
  ///     <em>
  ///       This module, both source code and documentation, is in the
  ///       Public Domain, and comes with <strong>NO WARRANTY</strong>.
  ///     </em>
  ///     See <a href='http://www.saxproject.org'>http://www.saxproject.org</a>
  ///     for further information.
  ///   </blockquote>
  ///   <para>
  ///     This class provides a default implementation of the SAX2
  ///     <see cref="IAttributes" /> interface, with the
  ///     addition of manipulators so that the list can be modified or
  ///     reused.
  ///   </para>
  ///   <para>There are two typical uses of this class:</para>
  ///   <ol>
  ///     <li>
  ///       to take a persistent snapshot of an Attributes object
  ///       in a <see cref="IContentHandler.StartElement" /> event; or
  ///     </li>
  ///     <li>to construct or modify an Attributes object in a SAX2 driver or filter.</li>
  ///   </ol>
  /// </summary>
  public class Attributes : IAttributes {
    private string[] _data;
    private int _length;

    /// <summary>
    ///   Construct a new, empty Attributes object.
    /// </summary>
    public Attributes() {
      _length = 0;
      _data = null;
    }

    /// <summary>
    ///   Copy an existing <see cref="IAttributes" /> object.
    ///   <para></para>
    ///   This constructor is especially useful inside a
    ///   <see cref="IContentHandler.StartElement" /> event.
    /// </summary>
    /// <param name="atts">
    ///   The existing IAttributes object.
    /// </param>
    public Attributes(IAttributes atts) {
      SetAttributesInternal(atts);
    }

    public int Length {
      get { return _length; }
    }

    public virtual string GetUri(int index) {
      if (index >= 0 && index < _length) {
        return _data[index * 5];
      }
      return null;
    }

    public virtual string GetLocalName(int index) {
      if (index >= 0 && index < _length) {
        return _data[index * 5 + 1];
      }
      return null;
    }

    public virtual string GetQName(int index) {
      if (index >= 0 && index < _length) {
        return _data[index * 5 + 2];
      }
      return null;
    }

    public virtual string GetType(int index) {
      if (index >= 0 && index < _length) {
        return _data[index * 5 + 3];
      }
      return null;
    }

    public virtual string GetValue(int index) {
      if (index >= 0 && index < _length) {
        return _data[index * 5 + 4];
      }
      return null;
    }

    public virtual int GetIndex(string uri, string localName) {
      int max = _length * 5;
      for (int i = 0; i < max; i += 5) {
        if (_data[i].Equals(uri) && _data[i + 1].Equals(localName)) {
          return i / 5;
        }
      }
      return -1;
    }

    public virtual int GetIndex(string qName) {
      int max = _length * 5;
      for (int i = 0; i < max; i += 5) {
        if (_data[i + 2].Equals(qName)) {
          return i / 5;
        }
      }
      return -1;
    }

    public virtual string GetType(string uri, string localName) {
      int max = _length * 5;
      for (int i = 0; i < max; i += 5) {
        if (_data[i].Equals(uri) && _data[i + 1].Equals(localName)) {
          return _data[i + 3];
        }
      }
      return null;
    }

    public virtual string GetType(string qName) {
      int max = _length * 5;
      for (int i = 0; i < max; i += 5) {
        if (_data[i + 2].Equals(qName)) {
          return _data[i + 3];
        }
      }
      return null;
    }

    public virtual string GetValue(string uri, string localName) {
      int max = _length * 5;
      for (int i = 0; i < max; i += 5) {
        if (_data[i].Equals(uri) && _data[i + 1].Equals(localName)) {
          return _data[i + 4];
        }
      }
      return null;
    }

    public virtual string GetValue(string qName) {
      int max = _length * 5;
      for (int i = 0; i < max; i += 5) {
        if (_data[i + 2].Equals(qName)) {
          return _data[i + 4];
        }
      }
      return null;
    }


    /// <summary>
    ///   Clear the attribute list for reuse.
    ///   <para>
    ///     Note that little memory is freed by this call:
    ///     the current array is kept so it can be
    ///     reused.
    ///   </para>
    /// </summary>
    public virtual void Clear() {
      ClearInternal();
    }

    /// <summary>
    ///   Copy an entire Attributes object.
    ///   <para>
    ///     It may be more efficient to reuse an existing object
    ///     rather than constantly allocating new ones.
    ///   </para>
    /// </summary>
    /// <param name="atts">The attributes to copy.</param>
    public virtual void SetAttributes(IAttributes atts) {
      SetAttributesInternal(atts);
    }

    /// <summary>
    ///   Add an attribute to the end of the list.
    ///   <para>
    ///     For the sake of speed, this method does no checking
    ///     to see if the attribute is already in the list: that is
    ///     the responsibility of the application.
    ///   </para>
    /// </summary>
    /// <param name="uri">
    ///   The Namespace URI, or the empty string if
    ///   none is available or Namespace processing is not
    ///   being performed.
    /// </param>
    /// <param name="localName">
    ///   The local name, or the empty string if
    ///   Namespace processing is not being performed.
    /// </param>
    /// <param name="qName">
    ///   The qualified (prefixed) name, or the empty string
    ///   if qualified names are not available.
    /// </param>
    /// <param name="type">The attribute type as a string.</param>
    /// <param name="value">The attribute value.</param>
    public virtual void AddAttribute(string uri, string localName, string qName, string type, string value) {
      EnsureCapacity(Length + 1);
      _data[Length * 5] = uri;
      _data[Length * 5 + 1] = localName;
      _data[Length * 5 + 2] = qName;
      _data[Length * 5 + 3] = type;
      _data[Length * 5 + 4] = value;
      _length++;
    }

    /// <summary>
    ///   Set an attribute in the list.
    ///   <para>
    ///     For the sake of speed, this method does no checking
    ///     for name conflicts or well-formedness: such checks are the
    ///     responsibility of the application.
    ///   </para>
    /// </summary>
    /// <param name="index">The index of the attribute (zero-based).</param>
    /// <param name="uri">
    ///   The Namespace URI, or the empty string if
    ///   none is available or Namespace processing is not
    ///   being performed.
    /// </param>
    /// <param name="localName">
    ///   The local name, or the empty string if
    ///   Namespace processing is not being performed.
    /// </param>
    /// <param name="qName">
    ///   The qualified (prefixed) name, or the empty string
    ///   if qualified names are not available.
    /// </param>
    /// <param name="type">The attribute type as a string.</param>
    /// <param name="value">The attribute value.</param>
    /// <exception cref="IndexOutOfRangeException">
    ///   When the supplied index does not point to an attribute in the list.
    /// </exception>
    public virtual void SetAttribute(int index, string uri, string localName, string qName, string type, string value) {
      if (index >= 0 && index < _length) {
        _data[index * 5] = uri;
        _data[index * 5 + 1] = localName;
        _data[index * 5 + 2] = qName;
        _data[index * 5 + 3] = type;
        _data[index * 5 + 4] = value;
      } else {
        BadIndex(index);
      }
    }

    /// <summary>
    ///   Remove an attribute from the list.
    /// </summary>
    /// <param name="index">The index of the attribute (zero-based).</param>
    /// <exception cref="IndexOutOfRangeException">
    ///   When the supplied index does not point to an attribute in the list.
    /// </exception>
    public virtual void RemoveAttribute(int index) {
      if (index >= 0 && index < _length) {
        if (index < _length - 1) {
          Array.Copy(_data, (index + 1) * 5, _data, index * 5, (_length - index - 1) * 5);
        }
        index = (_length - 1) * 5;
        _data[index++] = null;
        _data[index++] = null;
        _data[index++] = null;
        _data[index++] = null;
        _data[index] = null;
        _length--;
      } else {
        BadIndex(index);
      }
    }

    /// <summary>
    ///   Set the Namespace URI of a specific attribute.
    /// </summary>
    /// <param name="index">The index of the attribute (zero-based).</param>
    /// <param name="uri">
    ///   The attribute's Namespace URI, or the empty
    ///   string for none.
    /// </param>
    /// <exception cref="IndexOutOfRangeException">
    ///   When the supplied index does not point to an attribute in the list.
    /// </exception>
    public virtual void SetUri(int index, string uri) {
      if (index >= 0 && index < _length) {
        _data[index * 5] = uri;
      } else {
        BadIndex(index);
      }
    }

    /// <summary>
    ///   Set the local name of a specific attribute.
    /// </summary>
    /// <param name="index">The index of the attribute (zero-based).</param>
    /// <param name="localName">
    ///   The attribute's local name, or the empty
    ///   string for none.
    /// </param>
    /// <exception cref="IndexOutOfRangeException">
    ///   When the supplied index does not point to an attribute in the list.
    /// </exception>
    public virtual void SetLocalName(int index, string localName) {
      if (index >= 0 && index < _length) {
        _data[index * 5 + 1] = localName;
      } else {
        BadIndex(index);
      }
    }

    /// <summary>
    ///   Set the qualified name of a specific attribute.
    /// </summary>
    /// <param name="index">The index of the attribute (zero-based).</param>
    /// <param name="qName">
    ///   The attribute's qualified name, or the empty
    ///   string for none.
    /// </param>
    /// <exception cref="IndexOutOfRangeException">
    ///   When the supplied index does not point to an attribute in the list.
    /// </exception>
    public virtual void SetQName(int index, string qName) {
      if (index >= 0 && index < _length) {
        _data[index * 5 + 2] = qName;
      } else {
        BadIndex(index);
      }
    }

    /// <summary>
    ///   Set the type of a specific attribute.
    /// </summary>
    /// <param name="index">The index of the attribute (zero-based).</param>
    /// <param name="type">The attribute's type.</param>
    /// <exception cref="IndexOutOfRangeException">
    ///   When the supplied index does not point to an attribute in the list.
    /// </exception>
    public virtual void SetType(int index, string type) {
      if (index >= 0 && index < _length) {
        _data[index * 5 + 3] = type;
      } else {
        BadIndex(index);
      }
    }

    /// <summary>
    ///   Set the value of a specific attribute.
    /// </summary>
    /// <param name="index">The index of the attribute (zero-based).</param>
    /// <param name="value">The attribute's value.</param>
    /// <exception cref="IndexOutOfRangeException">
    ///   When the supplied index does not point to an attribute in the list.
    /// </exception>
    public virtual void SetValue(int index, string value) {
      if (index >= 0 && index < _length) {
        _data[index * 5 + 4] = value;
      } else {
        BadIndex(index);
      }
    }

    /// <summary>
    ///   Ensure the internal array's capacity.
    /// </summary>
    /// <param name="n">
    ///   The minimum number of attributes that the array must
    ///   be able to hold.
    /// </param>
    private void EnsureCapacity(int n) {
      if (n <= 0) {
        return;
      }
      int max;
      if (_data == null || _data.Length == 0) {
        max = 25;
      } else if (_data.Length >= n * 5) {
        return;
      } else {
        max = _data.Length;
      }
      while (max < n * 5) {
        max *= 2;
      }

      var newData = new string[max];
      if (_length > 0) {
        Array.Copy(_data, 0, newData, 0, _length * 5);
      }
      _data = newData;
    }

    /// <summary>
    ///   Report a bad array index in a manipulator.
    /// </summary>
    /// <param name="index">The index to report.</param>
    /// <exception cref="IndexOutOfRangeException">Always.</exception>
    private static void BadIndex(int index) {
      string msg = "Attempt to modify attribute at illegal index: " + index;
      throw new IndexOutOfRangeException(msg);
    }

    private void SetAttributesInternal(IAttributes atts) {
      ClearInternal();
      int length = atts.Length;
      if (length > 0) {
        _data = new string[length * 5];

        for (int i = 0; i < length; i++) {
          _data[i * 5] = atts.GetUri(i);
          _data[i * 5 + 1] = atts.GetLocalName(i);
          _data[i * 5 + 2] = atts.GetQName(i);
          _data[i * 5 + 3] = atts.GetType(i);
          _data[i * 5 + 4] = atts.GetValue(i);
        }
        _length = length;
      }
    }

    private void ClearInternal() {
      if (_data != null) {
        for (int i = 0; i < (_length * 5); i++) {
          _data[i] = null;
        }
      }
      _length = 0;
    }
  }

  // end of AttributesImpl.java
}
