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

namespace NPOI.SS.UserModel
{
    using System;


    public class CellCopyPolicy
    {
        // cell-level policies
        public static bool DEFAULT_COPY_CELL_VALUE_POLICY = true;
        public static bool DEFAULT_COPY_CELL_STYLE_POLICY = true;
        public static bool DEFAULT_COPY_CELL_FORMULA_POLICY = true;
        public static bool DEFAULT_COPY_HYPERLINK_POLICY = true;
        public static bool DEFAULT_MERGE_HYPERLINK_POLICY = false;

        // row-level policies
        public static bool DEFAULT_COPY_ROW_HEIGHT_POLICY = true;
        public static bool DEFAULT_CONDENSE_ROWS_POLICY = false;

        // column-level policies
        public static bool DEFAULT_COPY_COLUMN_WIDTH_POLICY = true;

        // sheet-level policies
        public static bool DEFAULT_COPY_MERGED_REGIONS_POLICY = true;

        // cell-level policies
        private bool copyCellValue = DEFAULT_COPY_CELL_VALUE_POLICY;
        private bool copyCellStyle = DEFAULT_COPY_CELL_STYLE_POLICY;
        private bool copyCellFormula = DEFAULT_COPY_CELL_FORMULA_POLICY;
        private bool copyHyperlink = DEFAULT_COPY_HYPERLINK_POLICY;
        private bool mergeHyperlink = DEFAULT_MERGE_HYPERLINK_POLICY;

        // row-level policies
        private bool copyRowHeight = DEFAULT_COPY_ROW_HEIGHT_POLICY;
        private bool condenseRows = DEFAULT_CONDENSE_ROWS_POLICY;

        // column-level policies
        private bool copyColumnWidth = DEFAULT_COPY_COLUMN_WIDTH_POLICY;

        // sheet-level policies
        private bool copyMergedRegions = DEFAULT_COPY_MERGED_REGIONS_POLICY;

        /** 
         * Default CellCopyPolicy, uses default policy
         * For custom CellCopyPolicy, use {@link Builder} class
         */
        public CellCopyPolicy() { }

        /**
         * Copy constructor
         *
         * @param other policy to copy
         */
        public CellCopyPolicy(CellCopyPolicy other)
        {
            copyCellValue = other.IsCopyCellValue;
            copyCellStyle = other.IsCopyCellStyle;
            copyCellFormula = other.IsCopyCellFormula;
            copyHyperlink = other.IsCopyHyperlink;
            mergeHyperlink = other.IsMergeHyperlink;

            copyRowHeight = other.IsCopyRowHeight;
            condenseRows = other.IsCondenseRows;

            copyColumnWidth = other.copyColumnWidth;

            copyMergedRegions = other.IsCopyMergedRegions;
        }

        // should builder be Replaced with CellCopyPolicy Setters that return the object
        // to allow Setters to be chained together?
        // policy.CopyCellValue=(/*setter*/true).CopyCellStyle=(/*setter*/true)
        private CellCopyPolicy(Builder builder)
        {
            copyCellValue = builder.copyCellValue;
            copyCellStyle = builder.copyCellStyle;
            copyCellFormula = builder.copyCellFormula;
            copyHyperlink = builder.copyHyperlink;
            mergeHyperlink = builder.mergeHyperlink;

            copyRowHeight = builder.copyRowHeight;
            condenseRows = builder.condenseRows;

            copyColumnWidth = builder.copyColumnWidth;

            copyMergedRegions = builder.copyMergedRegions;
        }

        public class Builder
        {
            // cell-level policies
            internal bool copyCellValue = DEFAULT_COPY_CELL_VALUE_POLICY;
            internal bool copyCellStyle = DEFAULT_COPY_CELL_STYLE_POLICY;
            internal bool copyCellFormula = DEFAULT_COPY_CELL_FORMULA_POLICY;
            internal bool copyHyperlink = DEFAULT_COPY_HYPERLINK_POLICY;
            internal bool mergeHyperlink = DEFAULT_MERGE_HYPERLINK_POLICY;

            // row-level policies
            internal bool copyRowHeight = DEFAULT_COPY_ROW_HEIGHT_POLICY;
            internal bool condenseRows = DEFAULT_CONDENSE_ROWS_POLICY;

            // column-level policies
            internal bool copyColumnWidth = DEFAULT_COPY_COLUMN_WIDTH_POLICY;

            // sheet-level policies
            internal bool copyMergedRegions = DEFAULT_COPY_MERGED_REGIONS_POLICY;

            /**
             * Builder class for CellCopyPolicy
             */
            public Builder()
            {
            }

            // cell-level policies
            public Builder CellValue(bool copyCellValue)
            {
                this.copyCellValue = copyCellValue;
                return this;
            }
            public Builder CellStyle(bool copyCellStyle)
            {
                this.copyCellStyle = copyCellStyle;
                return this;
            }
            public Builder CellFormula(bool copyCellFormula)
            {
                this.copyCellFormula = copyCellFormula;
                return this;
            }
            public Builder CopyHyperlink(bool copyHyperlink)
            {
                this.copyHyperlink = copyHyperlink;
                return this;
            }
            public Builder MergeHyperlink(bool mergeHyperlink)
            {
                this.mergeHyperlink = mergeHyperlink;
                return this;
            }

            // row-level policies
            public Builder RowHeight(bool copyRowHeight)
            {
                this.copyRowHeight = copyRowHeight;
                return this;
            }

            public Builder CondenseRows(bool condenseRows)
            {
                this.condenseRows = condenseRows;
                return this;
            }

            // column-level policies
            public Builder ColumnWidth(bool copyColumnWidth)
            {
                this.copyColumnWidth = copyColumnWidth;
                return this;
            }

            // sheet-level policies
            public Builder MergedRegions(bool copyMergedRegions)
            {
                this.copyMergedRegions = copyMergedRegions;
                return this;
            }
            public CellCopyPolicy Build()
            {
                return new CellCopyPolicy(this);
            }
        }

        public Builder CreateBuilder()
        {
            Builder builder = new Builder()
                    .CellValue(copyCellValue)
                    .CellStyle(copyCellStyle)
                    .CellFormula(copyCellFormula)
                    .CopyHyperlink(copyHyperlink)
                    .MergeHyperlink(mergeHyperlink)
                    .RowHeight(copyRowHeight)
                    .CondenseRows(condenseRows)
                    .ColumnWidth(copyColumnWidth)
                    .MergedRegions(copyMergedRegions);
            return builder;
        }

        /*
         * Cell-level policies 
         */
        /**
         * @return the copyCellValue
         */
        public bool IsCopyCellValue
        {
            get
            {
                return copyCellValue;
            }
            set
            {
                this.copyCellValue = value;
            }
        }

        /**
         * @return the copyCellStyle
         */
        public bool IsCopyCellStyle
        {
            get
            {
                return copyCellStyle;
            }
            set
            {
                this.copyCellStyle = value;
            }
        }

        /**
         * @return the copyCellFormula
         */
        public bool IsCopyCellFormula
        {
            get
            {
                return copyCellFormula;
            }
            set
            {
                this.copyCellFormula = value;
            }
        }

        /**
         * @return the copyHyperlink
         */
        public bool IsCopyHyperlink
        {
            get
            {
                return copyHyperlink;
            }
            set
            {
                this.copyHyperlink = value;
            }
        }

        /**
         * @return the mergeHyperlink
         */
        public bool IsMergeHyperlink
        {
            get
            {
                return mergeHyperlink;
            }
            set
            {
                this.mergeHyperlink = value;
            }
        }

        /*
         * Row-level policies 
         */
        /**
         * @return the copyRowHeight
         */
        public bool IsCopyRowHeight
        {
            get
            {
                return copyRowHeight;
            }
            set
            {
                this.copyRowHeight = value;
            }
        }

        /**
         * If condenseRows is true, a discontinuities in srcRows will be Removed when copied to destination
         * For example:
         * Sheet.CopyRows({Row(1), Row(2), Row(5)}, 11, policy) results in rows 1, 2, and 5
         * being copied to rows 11, 12, and 13 if condenseRows is True, or rows 11, 11, 15 if condenseRows is false
         * @return the condenseRows
         */
        public bool IsCondenseRows
        {
            get
            {
                return condenseRows;
            }
            set
            {
                this.condenseRows = value;
            }
        }

        /*
         * Column-level policies 
         */
        /**
         * @return the copyColumnWidth
         */
        public bool IsCopyColumnWidth
        {
            get
            {
                return copyColumnWidth;
            }
            set
            {
                this.copyColumnWidth = value;
            }
        }

        /*
         * Sheet-level policies 
         */
        /**
         * @return the copyMergedRegions
         */
        public bool IsCopyMergedRegions
        {
            get
            {
                return copyMergedRegions;
            }
            set
            {
                this.copyMergedRegions = value;
            }
        }

        public CellCopyPolicy Clone()
        {
            return (CellCopyPolicy)this.MemberwiseClone();
        }
    }

}