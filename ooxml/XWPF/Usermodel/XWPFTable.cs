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
namespace NPOI.XWPF.UserModel
{
    using System;
    using System.Text;
    using NPOI.OpenXmlFormats.Wordprocessing;
    using System.Collections.Generic;

    /**
     * <p>Sketch of XWPFTable class. Only table's text is being hold.</p>
     * <p>Specifies the contents of a table present in the document. A table is a set
     * of paragraphs (and other block-level content) arranged in rows and columns.</p>
     */
    public class XWPFTable : IBodyElement, ISDTContents
    {

        protected StringBuilder text = new StringBuilder();
        private CT_Tbl ctTbl;
        protected internal List<XWPFTableRow> tableRows = new List<XWPFTableRow>();
        //protected List<String> styleIDs;

        // Create a map from this XWPF-level enum to the STBorder.Enum values
        public enum XWPFBorderType { NIL, NONE, SINGLE, THICK, DOUBLE, DOTTED, DASHED, DOT_DASH };
        internal static Dictionary<XWPFBorderType, ST_Border> xwpfBorderTypeMap;
        // Create a map from the STBorder.Enum values to the XWPF-level enums
        internal static Dictionary<ST_Border, XWPFBorderType> stBorderTypeMap;

        protected IBody part;
        static XWPFTable()
        {
            // populate enum maps
            xwpfBorderTypeMap = new Dictionary<XWPFBorderType, ST_Border>();
            xwpfBorderTypeMap.Add(XWPFBorderType.NIL, ST_Border.nil);
            xwpfBorderTypeMap.Add(XWPFBorderType.NONE, ST_Border.none);
            xwpfBorderTypeMap.Add(XWPFBorderType.SINGLE, ST_Border.single);
            xwpfBorderTypeMap.Add(XWPFBorderType.THICK, ST_Border.thick);
            xwpfBorderTypeMap.Add(XWPFBorderType.DOUBLE, ST_Border.@double);
            xwpfBorderTypeMap.Add(XWPFBorderType.DOTTED, ST_Border.dotted);
            xwpfBorderTypeMap.Add(XWPFBorderType.DASHED, ST_Border.dashed);
            xwpfBorderTypeMap.Add(XWPFBorderType.DOT_DASH, ST_Border.dotDash);

            stBorderTypeMap = new Dictionary<ST_Border, XWPFBorderType>();
            stBorderTypeMap.Add(ST_Border.nil, XWPFBorderType.NIL);
            stBorderTypeMap.Add(ST_Border.none, XWPFBorderType.NONE);
            stBorderTypeMap.Add(ST_Border.single, XWPFBorderType.SINGLE);
            stBorderTypeMap.Add(ST_Border.thick, XWPFBorderType.THICK);
            stBorderTypeMap.Add(ST_Border.@double, XWPFBorderType.DOUBLE);
            stBorderTypeMap.Add(ST_Border.dotted, XWPFBorderType.DOTTED);
            stBorderTypeMap.Add(ST_Border.dashed, XWPFBorderType.DASHED);
            stBorderTypeMap.Add(ST_Border.dotDash, XWPFBorderType.DOT_DASH);
        }
        public XWPFTable(CT_Tbl table, IBody part, int row, int col)
            : this(table, part)
        {

            CT_TblGrid ctTblGrid = table.AddNewTblGrid();
            for (int j = 0; j < col; j++)
            {
                CT_TblGridCol ctGridCol= ctTblGrid.AddNewGridCol();
                ctGridCol.w = 300;
            }
            for (int i = 0; i < row; i++)
            {
                XWPFTableRow tabRow = (GetRow(i) == null) ? CreateRow() : GetRow(i);
                for (int k = 0; k < col; k++)
                {
                    if (tabRow.GetCell(k) == null)
                    {
                        tabRow.CreateCell();
                    }
                }
            }
        }

        public void SetColumnWidth(int columnIndex, ulong width)
        {
            if (this.ctTbl.tblGrid == null)
                return;

            if (columnIndex > this.ctTbl.tblGrid.gridCol.Count)
            {
                throw new ArgumentOutOfRangeException(string.Format("Column index {0} doesn't exist.", columnIndex));
            }
            this.ctTbl.tblGrid.gridCol[columnIndex].w = width;

        }

        public XWPFTable(CT_Tbl table, IBody part)
        {
            this.part = part;
            this.ctTbl = table;

            // is an empty table: I add one row and one column as default
            if (table.SizeOfTrArray() == 0)
                CreateEmptyTable(table);

            foreach (CT_Row row in table.GetTrList()) {
                StringBuilder rowText = new StringBuilder();
                row.Table = table;
                XWPFTableRow tabRow = new XWPFTableRow(row, this);
                tableRows.Add(tabRow);
                foreach (CT_Tc cell in row.GetTcList()) {
                    foreach (CT_P ctp in cell.GetPList()) {
                        XWPFParagraph p = new XWPFParagraph(ctp, part);
                        if (rowText.Length > 0) {
                            rowText.Append('\t');
                        }
                        rowText.Append(p.Text);
                    }
                }
                if (rowText.Length > 0) {
                    this.text.Append(rowText);
                    this.text.Append('\n');
                }
            }
        }
        
        private static void CreateEmptyTable(CT_Tbl table)
        {
            // MINIMUM ELEMENTS FOR A TABLE
            table.AddNewTr().AddNewTc().AddNewP();

            CT_TblPr tblpro = table.AddNewTblPr();
            if (!tblpro.IsSetTblW())
                tblpro.AddNewTblW().w = "0";
            tblpro.tblW.type=(ST_TblWidth.auto);

            // layout
             tblpro.AddNewTblLayout().type =  ST_TblLayoutType.autofit;

            // borders
            CT_TblBorders borders = tblpro.AddNewTblBorders();
            borders.AddNewBottom().val=ST_Border.single;
            borders.AddNewInsideH().val = ST_Border.single;
            borders.AddNewInsideV().val = ST_Border.single;
            borders.AddNewLeft().val = ST_Border.single;
            borders.AddNewRight().val = ST_Border.single;
            borders.AddNewTop().val = ST_Border.single;

            
            //CT_TblGrid tblgrid=table.AddNewTblGrid();
            //tblgrid.AddNewGridCol().w= (ulong)2000;
           
        }

        /**
         * @return ctTbl object
         */

        public CT_Tbl GetCTTbl()
        {
            return ctTbl;
        }

        /**
         * Convenience method to extract text in cells.  This
         * does not extract text recursively in cells, and it does not
         * currently include text in SDT (form) components.
         * <p>
         * To get all text within a table, see XWPFWordExtractor's appendTableText
         * as an example. 
         *
         * @return text
         */
        public String Text
        {
            get
            {
                return text.ToString();
            }
        }

        /**
         * add a new column for each row in this table
         */
        public void AddNewCol()
        {
            if (ctTbl.SizeOfTrArray() == 0) {
                CreateRow();
            }
            for (int i = 0; i < ctTbl.SizeOfTrArray(); i++) {
                XWPFTableRow tabRow = new XWPFTableRow(ctTbl.GetTrArray(i), this);
                tabRow.CreateCell();
            }
        }

        /**
         * create a new XWPFTableRow object with as many cells as the number of columns defined in that moment
         *
         * @return tableRow
         */
        public XWPFTableRow CreateRow()
        {
            int sizeCol = ctTbl.SizeOfTrArray() > 0 ? ctTbl.GetTrArray(0)
                    .SizeOfTcArray() : 0;
            XWPFTableRow tabRow = new XWPFTableRow(ctTbl.AddNewTr(), this);
            AddColumn(tabRow, sizeCol);
            tableRows.Add(tabRow);
            return tabRow;
        }

        /**
         * @param pos - index of the row
         * @return the row at the position specified or null if no rows is defined or if the position is greather than the max size of rows array
         */
        public XWPFTableRow GetRow(int pos)
        {
            if (pos >= 0 && pos < ctTbl.SizeOfTrArray()) {
                //return new XWPFTableRow(ctTbl.GetTrArray(pos));
                return Rows[(pos)];
            }
            return null;
        }


        /**
         * @return width value
         */
        public int Width
        {
            get
            {
                CT_TblPr tblPr = GetTblPr();
                return tblPr.IsSetTblW() ? int.Parse(tblPr.tblW.w) : -1;
            }
            set 
            {

                CT_TblPr tblPr = GetTblPr();
                CT_TblWidth tblWidth = tblPr.IsSetTblW() ? tblPr.tblW : tblPr
                        .AddNewTblW();
                tblWidth.w = value.ToString();
                tblWidth.type = ST_TblWidth.pct;
            }
        }

        /**
         * @return number of rows in table
         */
        public int NumberOfRows
        {
            get
            {
                return ctTbl.SizeOfTrArray();
            }
        }

        /// <summary>
        /// Returns CTTblPr object for table. Creates it if it does not exist.
        /// </summary>
        private CT_TblPr GetTblPr()
        {
            return GetTblPr(true);
        }

        /// <summary>
        /// Returns CTTblPr object for table. If force parameter is true, will
        /// create the element if necessary. If force parameter is false, returns
        /// null when CTTblPr element is missing.
        /// </summary>
        /// <param name="force">- force creation of CTTblPr element if necessary</param>
        private CT_TblPr GetTblPr(bool force)
        {
            return (ctTbl.tblPr != null) ? ctTbl.tblPr
                    : (force ? ctTbl.AddNewTblPr() : null);
        }

        /// <summary>
        /// Return CTTblBorders object for table. If force parameter is true, will
        /// create the element if necessary. If force parameter is false, returns
        /// null when CTTblBorders element is missing.
        /// </summary>
        /// <param name="force">- force creation of CTTblBorders element if necessary</param>
        private CT_TblBorders GetTblBorders(bool force)
        {
            CT_TblPr tblPr = GetTblPr(force);
            return tblPr == null ? null
                   : tblPr.IsSetTblBorders() ? tblPr.tblBorders
                   : force ? tblPr.AddNewTblBorders()
                   : null;
        }

        /// <summary>
        /// Return CTBorder object for Inside Vertical border. If force parameter is true,
        /// will create the element if necessary. If force parameter is false, returns
        /// null when Inside Vertical border is missing.
        /// </summary>
        /// <param name="force">- force creation of Inside Vertical border if necessary.</param>
        private CT_Border GetTblInsideVBorder(bool force)
        {
            CT_TblBorders ctb = GetTblBorders(force);
            return ctb == null ? null
                   : ctb.IsSetInsideV() ? ctb.insideV
                   : force ? ctb.AddNewInsideV()
                   : null;
        }

        /// <summary>
        /// Return CTBorder object for Inside Horizontal border. If force parameter is true,
        /// will create the element if necessary. If force parameter is false, returns
        /// null when Inside Horizontal border is missing.
        /// </summary>
        /// <param name="force">- force creation of Inside Horizontal border if necessary.</param>
        private CT_Border GetTblInsideHBorder(bool force)
        {
            CT_TblBorders ctb = GetTblBorders(force);
            return ctb == null ? null
                   : ctb.IsSetInsideH() ? ctb.insideH
                   : force ? ctb.AddNewInsideH()
                   : null;
        }

        /// <summary>
        /// Return CTBorder object for Top border. If force parameter is true,
        /// will create the element if necessary. If force parameter is false, returns
        /// null when Top border is missing.
        /// </summary>
        /// <param name="force">- force creation of Top border if necessary.</param>
        private CT_Border GetTblTopBorder(bool force)
        {
            CT_TblBorders ctb = GetTblBorders(force);
            return ctb == null ? null
                   : ctb.IsSetTop() ? ctb.top
                   : force ? ctb.AddNewTop()
                   : null;
        }

        /// <summary>
        /// Return CTBorder object for Bottom border. If force parameter is true,
        /// will create the element if necessary. If force parameter is false, returns
        /// null when Bottom border is missing.
        /// </summary>
        /// <param name="force">- force creation of Bottom border if necessary.</param>
        private CT_Border GetTblBottomBorder(bool force)
        {
            CT_TblBorders ctb = GetTblBorders(force);
            return ctb == null ? null
                   : ctb.IsSetBottom() ? ctb.bottom
                   : force ? ctb.AddNewBottom()
                   : null;
        }

        /// <summary>
        /// Return CTBorder object for Left border. If force parameter is true,
        /// will create the element if necessary. If force parameter is false, returns
        /// null when Left border is missing.
        /// </summary>
        /// <param name="force">- force creation of Left border if necessary.</param>
        private CT_Border GetTblLeftBorder(bool force)
        {
            CT_TblBorders ctb = GetTblBorders(force);
            return ctb == null ? null
                   : ctb.IsSetLeft() ? ctb.left
                   : force ? ctb.AddNewLeft()
                   : null;
        }

        /// <summary>
        /// Return CTBorder object for Right border. If force parameter is true,
        /// will create the element if necessary. If force parameter is false, returns
        /// null when Right border is missing.
        /// </summary>
        /// <param name="force">- force creation of Right border if necessary.</param>
        private CT_Border GetTblRightBorder(bool force)
        {
            CT_TblBorders ctb = GetTblBorders(force);
            return ctb == null ? null
                   : ctb.IsSetRight() ? ctb.right
                   : force ? ctb.AddNewRight()
                   : null;
        }

        public int NumberOfColumns
        {
            get
            {
                return ctTbl.SizeOfTrArray() > 0 ? ctTbl.GetTrArray(0)
                        .SizeOfTcArray() : 0;
            }
        }

        /// <summary>
        /// Get or set the current table alignment or NULL
        /// </summary>
        /// <returns>Table Alignment as a <see cref="TableRowAlign"/> enum</returns>
        public TableRowAlign? TableAlignment
        {
            get
            {
                CT_TblPr tPr = GetTblPr(false);
                return tPr == null ? null
                        : tPr.IsSetJc() ? TableRowAlignExtension.ValueOf((int)tPr.jc.val)
                        : null;
            }
            set
            {
                if(value.HasValue) 
                {
                    CT_TblPr tPr = GetTblPr(true);
                    CT_Jc jc = tPr.IsSetJc() ? tPr.jc : tPr.AddNewJc();
                    jc.val = (ST_Jc)value.Value.GetValue();
                }
                else
                {
                    RemoveTableAlignment();
                }
            }
        }
    
        /// <summary>
        /// Removes the table alignment attribute from a table
        /// </summary>
        public void RemoveTableAlignment() {
            CT_TblPr tPr = GetTblPr(false);
            if (tPr != null && tPr.IsSetJc()) {
                tPr.UnsetJc();
            }
        }
        private static void AddColumn(XWPFTableRow tabRow, int sizeCol)
        {
            if (sizeCol > 0)
            {
                for (int i = 0; i < sizeCol; i++)
                {
                    tabRow.CreateCell();
                }
            }
        }

        /**
         * Get the StyleID of the table
         * @return	style-ID of the table
         */
        public String StyleID
        {
            get
            {
                String styleId = null;
                CT_TblPr tblPr = ctTbl.tblPr;
                if (tblPr != null)
                {
                    CT_String styleStr = tblPr.tblStyle;
                    if (styleStr != null)
                    {
                        styleId = styleStr.val;
                    }
                }
                return styleId;
            }
            set
            {
                CT_TblPr tblPr = GetTblPr();
                CT_String styleStr = tblPr.tblStyle;
                if (styleStr == null)
                {
                    styleStr = tblPr.AddNewTblStyle();
                }
                styleStr.val = value;
            }
        }
        public XWPFBorderType? InsideHBorderType
        {
            get
            {
                CT_Border b = GetTblInsideHBorder(false);
                return (b != null) ? stBorderTypeMap[b.val] : null;
            }
        }

        public int InsideHBorderSize
        {
            get
            {
                CT_Border b = GetTblInsideHBorder(false);
                return (b != null)
                        ? (b.sz.HasValue ? (int)b.sz.Value : -1)
                                : -1;
            }
        }

        public int InsideHBorderSpace
        {
            get
            {
                CT_Border b = GetTblInsideHBorder(false);
                return (b != null)
                        ? (b.space.HasValue ? (int)b.space.Value : -1)
                                : -1;
            }
        }

        public String InsideHBorderColor
        {
            get
            {
                CT_Border b = GetTblInsideHBorder(false);
                //return (b != null)
                //        ? (b.IsSetColor() ? b.color : null)
                //                : null;
                return b?.color;
            }
        }

        public XWPFBorderType? InsideVBorderType
        {
            get
            {
                CT_Border b = GetTblInsideVBorder(false);
                return (b != null) ? stBorderTypeMap[b.val] : null;
            }
        }

        public int InsideVBorderSize
        {
            get
            {
                CT_Border b = GetTblInsideVBorder(false);
                return (b != null)
                        ? (b.sz.HasValue ? (int)b.sz.Value : -1)
                                : -1;
            }
        }

        public int InsideVBorderSpace
        {
            get
            {
                CT_Border b = GetTblInsideVBorder(false);
                return (b != null)
                        ? (b.space.HasValue ? (int)b.space.Value : -1)
                                : -1;
            }
        }

        public String InsideVBorderColor
        {
            get
            {
                CT_Border b = GetTblInsideVBorder(false);
                //return (b != null)
                //        ? (b.color ? b.xgetColor().getStringValue() : null)
                //                : null;
                return b?.color;
            }
        }

        /**
         * Get top border type
         *
         * @return {@link XWPFBorderType} of the top borders or null if missing
         */
        public XWPFBorderType? TopBorderType
        {
            get
            {
                CT_Border b = GetTblTopBorder(false);
                return (b != null) ? stBorderTypeMap[b.val] : null;
            }
        }

        /**
         * Get top border size
         * 
         * @return The width of the top borders in 1/8th points,
         * -1 if missing.
         */
        public int TopBorderSize
        {
            get
            {
                CT_Border b = GetTblTopBorder(false);
                return (b != null)
                        ? (b.sz.HasValue ? (int) b.sz.Value : -1)
                                : -1;
            }
        }

        /**
         * Get top border spacing
         * 
         * @return The offset to the top borders in points,
         * -1 if missing.
         */
        public int TopBorderSpace
        {
            get
            {
                CT_Border b = GetTblTopBorder(false);
                return (b != null)
                        ? (b.space.HasValue ? (int) b.space.Value : -1)
                                : -1;
            }
        }

        /**
         * Get top border color
         * 
         * @return The color of the top borders, null if missing.
         */
        public String TopBorderColor
        {
            get
            {
                CT_Border b = GetTblTopBorder(false);
                //return (b != null)
                //        ? (b.isSetColor() ? b.xgetColor().getStringValue() : null)
                //                : null;
                return b?.color;
            }
        }

        /**
         * Get bottom border type
         *
         * @return {@link XWPFBorderType} of the bottom borders or null if missing
         */
        public XWPFBorderType? BottomBorderType
        {
            get
            {
                CT_Border b = GetTblBottomBorder(false);
                return (b != null) ? stBorderTypeMap[b.val] : null;
            }
        }

        /**
         * Get bottom border size
         * 
         * @return The width of the bottom borders in 1/8th points,
         * -1 if missing.
         */
        public int BottomBorderSize
        {
            get
            {
                CT_Border b = GetTblBottomBorder(false);
                return (b != null)
                        ? (b.sz.HasValue ? (int) b.sz.Value : -1)
                                : -1;
            }
        }

        /**
         * Get bottom border spacing
         * 
         * @return The offset to the bottom borders in points,
         * -1 if missing.
         */
        public int BottomBorderSpace
        {
            get
            {
                CT_Border b = GetTblBottomBorder(false);
                return (b != null)
                        ? (b.space.HasValue ? (int) b.space.Value : -1)
                                : -1;
            }
        }

        /**
         * Get bottom border color
         * 
         * @return The color of the bottom borders, null if missing.
         */
        public String BottomBorderColor
        {
            get
            {
                CT_Border b = GetTblBottomBorder(false);
                //return (b != null)
                //        ? (b.isSetColor() ? b.xgetColor().getStringValue() : null)
                //                : null;
                return b?.color;
            }
        }

        /**
         * Get Left border type
         *
         * @return {@link XWPFBorderType} of the Left borders or null if missing
         */
        public XWPFBorderType? LeftBorderType
        {
            get
            {
                CT_Border b = GetTblLeftBorder(false);
                return (b != null) ? stBorderTypeMap[b.val] : null;
            }
        }

        /**
         * Get Left border size
         * 
         * @return The width of the Left borders in 1/8th points,
         * -1 if missing.
         */
        public int LeftBorderSize
        {
            get
            {
                CT_Border b = GetTblLeftBorder(false);
                return (b != null)
                        ? (b.sz.HasValue ? (int) b.sz.Value : -1)
                                : -1;
            }
        }

        /**
         * Get Left border spacing
         * 
         * @return The offset to the Left borders in points,
         * -1 if missing.
         */
        public int LeftBorderSpace
        {
            get
            {
                CT_Border b = GetTblLeftBorder(false);
                return (b != null)
                        ? (b.space.HasValue ? (int) b.space.Value : -1)
                                : -1;
            }
        }

        /**
         * Get Left border color
         * 
         * @return The color of the Left borders, null if missing.
         */
        public String LeftBorderColor
        {
            get
            {
                CT_Border b = GetTblLeftBorder(false);
                //return (b != null)
                //        ? (b.isSetColor() ? b.xgetColor().getStringValue() : null)
                //                : null;
                return b?.color;
            }
        }

        /**
         * Get Right border type
         *
         * @return {@link XWPFBorderType} of the Right borders or null if missing
         */
        public XWPFBorderType? RightBorderType
        {
            get
            {
                CT_Border b = GetTblRightBorder(false);
                return (b != null) ? stBorderTypeMap[b.val] : null;
            }
        }

        /**
         * Get Right border size
         * 
         * @return The width of the Right borders in 1/8th points,
         * -1 if missing.
         */
        public int RightBorderSize
        {
            get
            {
                CT_Border b = GetTblRightBorder(false);
                return (b != null)
                        ? (b.sz.HasValue ? (int) b.sz.Value : -1)
                                : -1;
            }
        }

        /**
         * Get Right border spacing
         * 
         * @return The offset to the Right borders in points,
         * -1 if missing.
         */
        public int RightBorderSpace
        {
            get
            {
                CT_Border b = GetTblRightBorder(false);
                return (b != null)
                        ? (b.space.HasValue ? (int) b.space.Value : -1)
                                : -1;
            }
        }

        /**
         * Get Right border color
         * 
         * @return The color of the Right borders, null if missing.
         */
        public String RightBorderColor
        {
            get
            {
                CT_Border b = GetTblRightBorder(false);
                //return (b != null)
                //        ? (b.isSetColor() ? b.xgetColor().getStringValue() : null)
                //                : null;
                return b?.color;
            }
        }
        public int RowBandSize
        {
            get
            {
                int size = 0;
                CT_TblPr tblPr = GetTblPr();
                if (tblPr.IsSetTblStyleRowBandSize())
                {
                    CT_DecimalNumber rowSize = tblPr.tblStyleRowBandSize;
                    int.TryParse(rowSize.val, out size);
                }
                return size;
            }
            set 
            {
                CT_TblPr tblPr = GetTblPr();
                CT_DecimalNumber rowSize = tblPr.IsSetTblStyleRowBandSize() ? tblPr.tblStyleRowBandSize : tblPr.AddNewTblStyleRowBandSize();
                rowSize.val = value.ToString();			
            }
        }

        public int ColBandSize
        {
            get
            {
                int size = 0;
                CT_TblPr tblPr = GetTblPr();
                if (tblPr.IsSetTblStyleColBandSize())
                {
                    CT_DecimalNumber colSize = tblPr.tblStyleColBandSize;
                    int.TryParse(colSize.val, out size);
                }
                return size;
            }
            set 
            {
                CT_TblPr tblPr = GetTblPr();
                CT_DecimalNumber colSize = tblPr.IsSetTblStyleColBandSize() ? tblPr.tblStyleColBandSize : tblPr.AddNewTblStyleColBandSize();
                colSize.val = value.ToString();
            }
        }

        /// <summary>
        /// Set Inside horizontal borders for a table
        /// </summary>
        /// <param name="type">- <see cref="XWPFBorderType"/> e.g. single, double, thick</param>
        /// <param name="size">- Specifies the width of the current border. The width of this border is
        /// specified in measurements of eighths of a point, with a minimum value of two (onefourth
        /// of a point) and a maximum value of 96 (twelve points). Any values outside this
        /// range may be reassigned to a more appropriate value.
        /// </param>
        /// <param name="space">- Specifies the spacing offset that shall be used to place this border on the table</param>
        /// <param name="rgbColor">- This color may either be presented as a hex value (in RRGGBB format),
        /// or auto to allow a consumer to automatically determine the border color as appropriate.
        /// </param>
        public void SetInsideHBorder(XWPFBorderType type, int size, int space, String rgbColor)
        {
            CT_Border b = GetTblInsideHBorder(true);
            b.val = (xwpfBorderTypeMap[(type)]);
            b.sz = (ulong) size;
            b.space = (ulong) space;
            b.color = (rgbColor);
        }

        /// <summary>
        /// Set Inside Vertical borders for table
        /// </summary>
        /// <param name="type">- <see cref="XWPFBorderType"/> e.g. single, double, thick</param>
        /// <param name="size">- Specifies the width of the current border. The width of this border is
        /// specified in measurements of eighths of a point, with a minimum value of two (onefourth
        /// of a point) and a maximum value of 96 (twelve points). Any values outside this
        /// range may be reassigned to a more appropriate value.
        /// </param>
        /// <param name="space">- Specifies the spacing offset that shall be used to place this border on the table</param>
        /// <param name="rgbColor">- This color may either be presented as a hex value (in RRGGBB format),
        /// or auto to allow a consumer to automatically determine the border color as appropriate.
        /// </param>
        public void SetInsideVBorder(XWPFBorderType type, int size, int space, String rgbColor)
        {
            CT_Border b = GetTblInsideVBorder(true);
            b.val = (xwpfBorderTypeMap[type]);
            b.sz = (ulong) size;
            b.space = (ulong) space;
            b.color = (rgbColor);
        }

        /// <summary>
        /// Set Top borders for table
        /// </summary>
        /// <param name="type">- <see cref="XWPFBorderType"/> e.g. single, double, thick</param>
        /// <param name="size">- Specifies the width of the current border. The width of this border is
        /// specified in measurements of eighths of a point, with a minimum value of two (onefourth
        /// of a point) and a maximum value of 96 (twelve points). Any values outside this
        /// range may be reassigned to a more appropriate value.
        /// </param>
        /// <param name="space">- Specifies the spacing offset that shall be used to place this border on the table</param>
        /// <param name="rgbColor">- This color may either be presented as a hex value (in RRGGBB format),
        /// or auto to allow a consumer to automatically determine the border color as appropriate.
        /// </param>
        public void SetTopBorder(XWPFBorderType type, int size, int space, String rgbColor)
        {
            CT_Border b = GetTblTopBorder(true);
            b.val = xwpfBorderTypeMap[type];
            b.sz = (ulong)size;
            b.space = (ulong)space;
            b.color = (rgbColor);
        }

        /// <summary>
        /// Set Bottom borders for table
        /// </summary>
        /// <param name="type">- <see cref="XWPFBorderType"/> e.g. single, double, thick</param>
        /// <param name="size">- Specifies the width of the current border. The width of this border is
        /// specified in measurements of eighths of a point, with a minimum value of two (onefourth
        /// of a point) and a maximum value of 96 (twelve points). Any values outside this
        /// range may be reassigned to a more appropriate value.
        /// </param>
        /// <param name="space">- Specifies the spacing offset that shall be used to place this border on the table</param>
        /// <param name="rgbColor">- This color may either be presented as a hex value (in RRGGBB format),
        /// or auto to allow a consumer to automatically determine the border color as appropriate.
        /// </param>
        public void SetBottomBorder(XWPFBorderType type, int size, int space, String rgbColor)
        {
            CT_Border b = GetTblBottomBorder(true);
            b.val = xwpfBorderTypeMap[type];
            b.sz = (ulong)size;
            b.space = (ulong)space;
            b.color = (rgbColor);
        }

        /// <summary>
        /// Set Left borders for table
        /// </summary>
        /// <param name="type">- <see cref="XWPFBorderType"/> e.g. single, double, thick</param>
        /// <param name="size">- Specifies the width of the current border. The width of this border is
        /// specified in measurements of eighths of a point, with a minimum value of two (onefourth
        /// of a point) and a maximum value of 96 (twelve points). Any values outside this
        /// range may be reassigned to a more appropriate value.
        /// </param>
        /// <param name="space">- Specifies the spacing offset that shall be used to place this border on the table</param>
        /// <param name="rgbColor">- This color may either be presented as a hex value (in RRGGBB format),
        /// or auto to allow a consumer to automatically determine the border color as appropriate.
        /// </param>
        public void SetLeftBorder(XWPFBorderType type, int size, int space, String rgbColor)
        {
            CT_Border b = GetTblLeftBorder(true);
            b.val = xwpfBorderTypeMap[type];
            b.sz = (ulong)size;
            b.space = (ulong)space;
            b.color = (rgbColor);
        }

        /// <summary>
        /// Set Right borders for table
        /// </summary>
        /// <param name="type">- <see cref="XWPFBorderType"/> e.g. single, double, thick</param>
        /// <param name="size">- Specifies the width of the current border. The width of this border is
        /// specified in measurements of eighths of a point, with a minimum value of two (onefourth
        /// of a point) and a maximum value of 96 (twelve points). Any values outside this
        /// range may be reassigned to a more appropriate value.
        /// </param>
        /// <param name="space">- Specifies the spacing offset that shall be used to place this border on the table</param>
        /// <param name="rgbColor">- This color may either be presented as a hex value (in RRGGBB format),
        /// or auto to allow a consumer to automatically determine the border color as appropriate.
        /// </param>
        public void SetRightBorder(XWPFBorderType type, int size, int space, String rgbColor)
        {
            CT_Border b = GetTblRightBorder(true);
            b.val = xwpfBorderTypeMap[type];
            b.sz = (ulong)size;
            b.space = (ulong)space;
            b.color = (rgbColor);
        }

        /// <summary>
        /// Remove inside horizontal borders for table
        /// </summary>
        public void RemoveInsideHBorder()
        {
            CT_Border b = GetTblInsideHBorder(false);
            if(b != null)
            {
                GetTblBorders(false).UnsetInsideH();
                CleanupTblBorders();
            }
        }

        /// <summary>
        /// Remove inside vertical borders for table
        /// </summary>
        public void RemoveInsideVBorder()
        {
            CT_Border b = GetTblInsideVBorder(false);
            if(b != null)
            {
                GetTblBorders(false).UnsetInsideV();
                CleanupTblBorders();
            }
        }

        /// <summary>
        /// Remove top borders for table
        /// </summary>
        public void RemoveTopBorder()
        {
            CT_Border b = GetTblTopBorder(false);
            if(b != null)
            {
                GetTblBorders(false).UnsetTop();
                CleanupTblBorders();
            }
        }

        /// <summary>
        /// Remove bottom borders for table
        /// </summary>
        public void RemoveBottomBorder()
        {
            CT_Border b = GetTblBottomBorder(false);
            if(b != null)
            {
                GetTblBorders(false).UnsetBottom();
                CleanupTblBorders();
            }
        }

        /// <summary>
        /// Remove left borders for table
        /// </summary>
        public void RemoveLeftBorder()
        {
            CT_Border b = GetTblLeftBorder(false);
            if(b != null)
            {
                GetTblBorders(false).UnsetLeft();
                CleanupTblBorders();
            }
        }

        /// <summary>
        /// Remove right borders for table
        /// </summary>
        public void RemoveRightBorder()
        {
            CT_Border b = GetTblRightBorder(false);
            if(b != null)
            {
                GetTblBorders(false).UnsetRight();
                CleanupTblBorders();
            }
        }

        /// <summary>
        /// Remove all borders from table
        /// </summary>
        public void RemoveBorders()
        {
            CT_TblBorders b = GetTblBorders(false);
            if(b != null)
            {
                GetTblPr(false).UnsetTblBorders();
            }
        }

        /// <summary>
        /// removes the Borders node from Table properties if there are
        /// no border elements
        /// </summary>
        private void CleanupTblBorders()
        {
            CT_TblBorders b = GetTblBorders(false);
            if(b != null)
            {
                if(b.insideH == null &&
                        b.insideV == null &&
                        b.top == null &&
                        b.bottom == null &&
                        b.left == null &&
                        b.right == null)
                {
                    GetTblPr(false).UnsetTblBorders();
                }
            }
        }

        public int CellMarginTop
        {
            get
            {
                int margin = 0;
                CT_TblPr tblPr = GetTblPr();
                CT_TblCellMar tcm = tblPr.tblCellMar;
                if (tcm != null)
                {
                    CT_TblWidth tw = tcm.top;
                    if (tw != null)
                    {
                        int.TryParse(tw.w, out margin);
                    }
                }
                return margin;
            }
        }

        public int CellMarginLeft
        {
            get
            {
                int margin = 0;
                CT_TblPr tblPr = GetTblPr();
                CT_TblCellMar tcm = tblPr.tblCellMar;
                if (tcm != null)
                {
                    CT_TblWidth tw = tcm.left;
                    if (tw != null)
                    {
                        int.TryParse(tw.w, out margin);
                    }
                }
                return margin;
            }
        }

        public int CellMarginBottom
        {
            get
            {
                int margin = 0;
                CT_TblPr tblPr = GetTblPr();
                CT_TblCellMar tcm = tblPr.tblCellMar;
                if (tcm != null)
                {
                    CT_TblWidth tw = tcm.bottom;
                    if (tw != null)
                    {
                        int.TryParse(tw.w, out margin);
                    }
                }
                return margin;
            }
        }

        public int CellMarginRight
        {
            get
            {
                int margin = 0;
                CT_TblPr tblPr = GetTblPr();
                CT_TblCellMar tcm = tblPr.tblCellMar;
                if (tcm != null)
                {
                    CT_TblWidth tw = tcm.right;
                    if (tw != null)
                    {
                        int.TryParse(tw.w, out margin);
                    }
                }
                return margin;
            }
        }

        public string TableCaption
        {
            get
            {
                CT_TblPr tblPr = GetTblPr();
                if (tblPr.tblCaption != null)
                    return tblPr.tblCaption.val;
                else
                    return string.Empty;
            }
            set
            {
                CT_TblPr tblPr = GetTblPr();
                if (tblPr.tblCaption == null)
                {
                    CT_String caption = new CT_String();
                    caption.val = value;
                    tblPr.tblCaption = caption;
                }
                else
                {
                    tblPr.tblCaption.val = value;
                }
            }
        }

        public string TableDescription
        {
            get
            {
                CT_TblPr tblPr = GetTblPr();
                if (tblPr.tblDescription != null)
                    return tblPr.tblDescription.val;
                else
                    return string.Empty;
            }
            set
            {
                CT_TblPr tblPr = GetTblPr();
                if (tblPr.tblDescription == null)
                {
                    CT_String desc = new CT_String();
                    desc.val = value;
                    tblPr.tblDescription = desc;
                }
                else
                {
                    tblPr.tblDescription.val = value;
                }
            }
        }

        public void SetCellMargins(int top, int left, int bottom, int right)
        {
            CT_TblPr tblPr = GetTblPr();
            CT_TblCellMar tcm = tblPr.IsSetTblCellMar() ? tblPr.tblCellMar : tblPr.AddNewTblCellMar();

            CT_TblWidth tw = tcm.IsSetLeft() ? tcm.left : tcm.AddNewLeft();
            tw.type = (ST_TblWidth.dxa);
            tw.w = left.ToString();

            tw = tcm.IsSetTop() ? tcm.top : tcm.AddNewTop();
            tw.type = (ST_TblWidth.dxa);
            tw.w = top.ToString();

            tw = tcm.IsSetBottom() ? tcm.bottom : tcm.AddNewBottom();
            tw.type = (ST_TblWidth.dxa);
            tw.w = bottom.ToString();

            tw = tcm.IsSetRight() ? tcm.right : tcm.AddNewRight();
            tw.type = (ST_TblWidth.dxa);
            tw.w = right.ToString();
        }
    
        /**
         * add a new Row to the table
         * 
         * @param row	the row which should be Added
         */
        public void AddRow(XWPFTableRow row)
        {
            ctTbl.AddNewTr();
            ctTbl.SetTrArray(this.NumberOfRows-1, row.GetCTRow());
            tableRows.Add(row);
        }

        /**
         * add a new Row to the table
         * at position pos
         * @param row	the row which should be Added
         */
        public bool AddRow(XWPFTableRow row, int pos)
        {
            if (pos >= 0 && pos <= tableRows.Count)
            {
                ctTbl.InsertNewTr(pos);
                ctTbl.SetTrArray(pos, row.GetCTRow());
                tableRows.Insert(pos, row);
                return true;
            }
            return false;
        }

        /**
         * inserts a new tablerow 
         * @param pos
         * @return  the inserted row
         */
        public XWPFTableRow InsertNewTableRow(int pos)
        {
            if(pos >= 0 && pos <= tableRows.Count){
                CT_Row row = ctTbl.InsertNewTr(pos);
                XWPFTableRow tableRow = new XWPFTableRow(row, this);
                tableRows.Insert(pos, tableRow);
                return tableRow;
            }
            return null;
        }


        /**
         * Remove a row at position pos from the table
         * @param pos	position the Row in the Table
         */
        public bool RemoveRow(int pos)
        {
            if (pos >= 0 && pos < tableRows.Count) {
                if (ctTbl.SizeOfTrArray() > 0)
                {
                    ctTbl.RemoveTr(pos);
                }
                tableRows.RemoveAt(pos);
                return true;
            }
            return false;
        }

        public List<XWPFTableRow> Rows
        {
            get
            {
                return tableRows;
            }
        }


        /**
         * returns the type of the BodyElement Table
         * @see NPOI.XWPF.UserModel.IBodyElement#getElementType()
         */
        public BodyElementType ElementType
        {
            get
            {
                return BodyElementType.TABLE;
            }
        }

        public IBody Body
        {
            get
            {
                return part;
            }
        }

        /**
         * returns the part of the bodyElement
         * @see NPOI.XWPF.UserModel.IBody#getPart()
         */
        public POIXMLDocumentPart Part
        {
            get
            {
                if (part != null)
                {
                    return part.Part;
                }
                return null;
            }
        }

        /**
         * returns the partType of the bodyPart which owns the bodyElement
         * @see NPOI.XWPF.UserModel.IBody#getPartType()
         */
        public BodyType PartType
        {
            get
            {
                return part.PartType;
            }
        }

        /**
         * returns the XWPFRow which belongs to the CTRow row
         * if this row is not existing in the table null will be returned
         */
        public XWPFTableRow GetRow(CT_Row row)
        {
            for(int i=0; i<Rows.Count; i++){
                if(Rows[(i)].GetCTRow() == row) return GetRow(i); 
            }
            return null;
        }

    }// end class

}
