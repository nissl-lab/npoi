/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */
using NPOI.SS.UserModel;
using NPOI.OpenXml4Net.OPC;
using System;
using NPOI.SS.Util;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.Util;

namespace NPOI.XSSF.UserModel
{
    /**
     * XSSF Implementation of a Hyperlink.
     * Note - unlike with HSSF, many kinds of hyperlink
     * are largely stored as relations of the sheet
     */
    public class XSSFHyperlink : IHyperlink
    {
        private HyperlinkType _type;
        private PackageRelationship _externalRel;
        private CT_Hyperlink _ctHyperlink; //contains a reference to the cell where the hyperlink is anchored, getRef()
        private String _location; //what the hyperlink refers to

        /**
         * Create a new XSSFHyperlink. This method is protected to be used only by XSSFCreationHelper
         *
         * @param type - the type of hyperlink to create
         */
        public XSSFHyperlink(HyperlinkType type)
        {
            _type = type;
            _ctHyperlink = new CT_Hyperlink();
            _externalRel = null;
        }

        /**
         * Create a XSSFHyperlink amd Initialize it from the supplied CTHyperlink bean and namespace relationship
         *
         * @param ctHyperlink the xml bean Containing xml properties
         * @param hyperlinkRel the relationship in the underlying OPC namespace which stores the actual link's Address
         */
        public XSSFHyperlink(CT_Hyperlink ctHyperlink, PackageRelationship hyperlinkRel)
        {
            _ctHyperlink = ctHyperlink;
            _externalRel = hyperlinkRel;

            // Figure out the Hyperlink type and destination

            if (_externalRel == null)
            {
                // If it has a location, it's internal
                if (ctHyperlink.location != null)
                {
                    _type = HyperlinkType.Document;
                    _location = ctHyperlink.location;
                }
                else if (ctHyperlink.id != null)
                {
                    throw new InvalidOperationException("The hyperlink for cell "
                            + ctHyperlink.@ref + " references relation "
                            + ctHyperlink.id + ", but that didn't exist!");
                }
                else
                {
                    // hyperlink is internal and is not related to other parts
                    _type = HyperlinkType.Document;
                }
            }
            else
            {
                Uri target = _externalRel.TargetUri;
                _location = target.ToString();
                if (ctHyperlink.location != null)
                {
                    // URI fragment
                    _location += "#" + ctHyperlink.location;
                }

                // Try to figure out the type
                if (_location.StartsWith("http://") || _location.StartsWith("https://")
                     || _location.StartsWith("ftp://"))
                {
                    _type = HyperlinkType.Url;
                }
                else if (_location.StartsWith("mailto:"))
                {
                    _type = HyperlinkType.Email;
                }
                else
                {
                    _type = HyperlinkType.File;
                }
            }
        }

        /**
         * Create a new XSSFHyperlink. This method is for Internal use only.
         * XSSFHyperlinks can be created by XSSFCreationHelper.
         *
         * @param type - the type of hyperlink to create, see {@link Hyperlink}
         */
        //FIXME: change to protected if/when SXSSFHyperlink class is created
        public XSSFHyperlink(IHyperlink other)
        {
            if (other is XSSFHyperlink)
            {
                XSSFHyperlink xlink = (XSSFHyperlink)other;
                _type = xlink.Type;
                _location = xlink._location;
                _externalRel = xlink._externalRel;
                _ctHyperlink = xlink._ctHyperlink.Copy();
            }
            else
            {
                _type = other.Type;
                _location = other.Address;
                _externalRel = null;
                _ctHyperlink = new CT_Hyperlink();
                SetCellReference(new CellReference(other.FirstRow, other.FirstColumn));
            }
        }

        /**
         * @return the underlying CTHyperlink object
         */
        public CT_Hyperlink GetCTHyperlink()
        {
            return _ctHyperlink;
        }

        /**
         * Do we need to a relation too, to represent
         * this hyperlink?
         */
        public bool NeedsRelationToo()
        {
            return (_type != HyperlinkType.Document);
        }

        /**
         * Generates the relation if required
         */
        internal void GenerateRelationIfNeeded(PackagePart sheetPart)
        {
            if (_externalRel == null && NeedsRelationToo())
            {
                // Generate the relation
                PackageRelationship rel =
                        sheetPart.AddExternalRelationship(_location, XSSFRelation.SHEET_HYPERLINKS.Relation);

                // Update the r:id
                _ctHyperlink.id = rel.Id;
            }
        }

        /**
         * Return the type of this hyperlink
         *
         * @return the type of this hyperlink
         */
        public HyperlinkType Type
        {
            get
            {
                return _type;
            }
        }
        [Obsolete("use property CellRef")]
        public string GetCellRef()
        {
            return _ctHyperlink.@ref;
        }
        /**
         * Get the reference of the cell this applies to,
         * es A55
         */
        public String CellRef
        {
            get
            {
                return _ctHyperlink.@ref;
            }
        }

        /**
         * Hyperlink Address. Depending on the hyperlink type it can be URL, e-mail, path to a file
         *
         * @return the Address of this hyperlink
         */
        public String Address
        {
            get
            {
                return _location;
            }
            set
            {
                Validate(value);
                _location = value;
                //we must Set location for internal hyperlinks
                if (_type == HyperlinkType.Document)
                {
                    this.Location = value;
                }
            }
        }

        private void Validate(String address)
        {
            switch (_type)
            {
                // email, path to file and url must be valid URIs
                case HyperlinkType.Email:
                case HyperlinkType.File:
                case HyperlinkType.Url:
                    Uri uri;
                    if (!Uri.TryCreate(address, UriKind.RelativeOrAbsolute, out uri))
                        throw new ArgumentException("Address of hyperlink must be a valid URI:" + address);
                    break;
                case HyperlinkType.Document:
                    // currently not evaluating anything.
                    break;
                default:
                    // this check wouldn't need to be done if _type was checked when object was set
                    // since _type is final, this check would only need to be done once
                    throw new InvalidOperationException("Invalid Hyperlink type: " + _type);
            }
        }

        /**
         * Return text label for this hyperlink
         *
         * @return text to display
         */
        public String Label
        {
            get
            {
                return _ctHyperlink.display;
            }
            set
            {
                _ctHyperlink.display = value;
            }
        }

        /**
         * Location within target. If target is a workbook (or this workbook) this shall refer to a
         * sheet and cell or a defined name. Can also be an HTML anchor if target is HTML file.
         *
         * @return location
         */
        public String Location
        {
            get
            {
                return _ctHyperlink.location;
            }

            set
            {
                _ctHyperlink.location = value;
            }
        }


        /**
         * Assigns this hyperlink to the given cell reference
         */
        public void SetCellReference(String ref1)
        {
            _ctHyperlink.@ref = ref1;
        }

        protected void SetCellReference(CellReference ref1)
        {
            SetCellReference(ref1.FormatAsString());
        }


        private CellReference buildCellReference()
        {
            String ref1 = _ctHyperlink.@ref;
            if (ref1 == null)
            {
                ref1 = "A1";
            }
            return new CellReference(ref1);
        }


        /**
         * Return the column of the first cell that Contains the hyperlink
         *
         * @return the 0-based column of the first cell that Contains the hyperlink
         */
        public int FirstColumn
        {
            get
            {
                return buildCellReference().Col;
            }
            set
            {
                SetCellReference(new CellReference(FirstRow, value));
            }
        }


        /**
         * Return the column of the last cell that Contains the hyperlink
         * For XSSF, a Hyperlink may only reference one cell
         * 
         * @return the 0-based column of the last cell that Contains the hyperlink
         */
        public int LastColumn
        {
            get
            {
                return buildCellReference().Col;
            }
            set
            {
                this.FirstColumn = value;
            }
        }

        /**
         * Return the row of the first cell that Contains the hyperlink
         *
         * @return the 0-based row of the cell that Contains the hyperlink
         */
        public int FirstRow
        {
            get
            {
                return buildCellReference().Row;
            }
            set
            {
                SetCellReference(new CellReference(value, FirstColumn));
            }
        }


        /**
         * Return the row of the last cell that Contains the hyperlink
         * For XSSF, a Hyperlink may only reference one cell
         *
         * @return the 0-based row of the last cell that Contains the hyperlink
         */
        public int LastRow
        {
            get
            {
                return buildCellReference().Row;
            }
            set
            {
                this.FirstRow = value;
            }
        }

        public string TextMark
        {
            get
            { throw new NotImplementedException(); }
            set
            { throw new NotImplementedException(); }
        }
        /// <summary>
        /// get or set additional text to help the user understand more about the hyperlink
        /// </summary>
        public String Tooltip
        {
            get
            {
                return _ctHyperlink.tooltip;
            }
            set
            {
                _ctHyperlink.tooltip = (value);
            }
        }

    }

}
