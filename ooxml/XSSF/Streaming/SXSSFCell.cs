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
using EnumsNET;
using NPOI.SS;
using NPOI.SS.Formula.Eval;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.Util;
using NPOI.XSSF.Streaming.Properties;
using NPOI.XSSF.Streaming.Values;
using NPOI.XSSF.UserModel;
using System;

namespace NPOI.XSSF.Streaming
{
    public class SXSSFCell : CellBase
    {
        private static readonly POILogger logger = POILogFactory.GetLogger(typeof(SXSSFCell));

        private readonly SXSSFRow _row;
        private Value _value;
        private ICellStyle _style;
        private Property _firstProperty;

        public SXSSFCell(SXSSFRow row, CellType cellType)
        {
            _row = row;
            _value = new BlankValue();
            SetType(cellType);
        }

        public override CellRangeAddress ArrayFormulaRange
        {
            get
            {
                return null;
            }
        }

        public override bool BooleanCellValue
        {
            get
            {
                CellType cellType = _value.GetType();
                switch (cellType)
                {
                    case CellType.Blank:
                        return false;
                    case CellType.Formula:
                        {
                            FormulaValue fv = (FormulaValue)_value;
                            if (fv.GetFormulaType() != CellType.Boolean)
                                throw typeMismatch(CellType.Boolean, CellType.Formula, false);
                            return ((BooleanFormulaValue)_value).PreEvaluatedValue;
                        }
                    case CellType.Boolean:
                        {
                            return ((BooleanValue)_value).Value;
                        }
                    default:
                        throw typeMismatch(CellType.Boolean, cellType, false);
                }
            }
        }

        //private string BuildTypeMismatchMessage(CellType expectedTypeCode, CellType actualTypeCode,
        //    bool isFormulaCell)
        //{
        //    return string.Format("Cannot get a {0} value from a {1} {2} cell", expectedTypeCode, actualTypeCode,(isFormulaCell ? "formula " : ""));
        //}

        public override CellType CachedFormulaResultType
        {
            get
            {
                if(!IsFormulaCell())
                {
                    throw new InvalidOperationException("Only formula cells have cached results");
                }

                return ((FormulaValue) _value).GetFormulaType();
            }
        }

        [Obsolete("Will be removed at NPOI 2.8, Use CachedFormulaResultType instead.")]
        [Removal(Version = "4.2")]
        public override CellType GetCachedFormulaResultTypeEnum()
        {
            return CachedFormulaResultType;
        }

        public override IComment CellComment
        {
            get
            {
                return (IComment)GetPropertyValue(Property.COMMENT);
            }

            set
            {
                SetProperty(Property.COMMENT, value);
            }
        }

        public override string CellFormula
        {
            get
            {
                if (_value.GetType() != CellType.Formula)
                    throw typeMismatch(CellType.Formula, _value.GetType(), false);
                return ((FormulaValue) _value).Value;
            }

            set
            {
                SetCellFormula(value);
            }
        }
        protected override void RemoveFormulaImpl()
        {
            if (CellType != CellType.Formula)
            {
                return;
            }
            switch (CachedFormulaResultType)
            {
                case CellType.Numeric:
                    double numericValue = ((NumericFormulaValue)_value).PreEvaluatedValue;
                    _value = new NumericValue();
                    ((NumericValue)_value).Value = numericValue;
                    break;
                case CellType.String:
                    String stringValue = ((StringFormulaValue)_value).PreEvaluatedValue;
                    _value = new PlainStringValue();
                    ((PlainStringValue)_value).Value = stringValue;
                    break;
                case CellType.Boolean:
                    bool booleanValue = ((BooleanFormulaValue)_value).PreEvaluatedValue;
                    _value = new BooleanValue();
                    ((BooleanValue)_value).Value = booleanValue;
                    break;
                case CellType.Error:
                    byte errorValue = (byte)((ErrorFormulaValue)_value).PreEvaluatedValue;
                    _value = new ErrorValue();
                    ((ErrorValue)_value).Value = errorValue;
                    break;
                default:
                    throw new InvalidOperationException();
            }
        }
        public override ICellStyle CellStyle
        {
            get
            {
                if (_style == null)
                {
                    SXSSFWorkbook wb = (SXSSFWorkbook)Row.Sheet.Workbook;
                    return wb.GetCellStyleAt(0);
                }
                else
                {
                    return _style;
                }
            }

            set
            {
                _style = value;
            }
        }
        private bool IsFormulaCell()
        {
            return _value is FormulaValue;
        }
        public override CellType CellType
        {
            get 
            { 
                if(IsFormulaCell())
                {
                    return CellType.Formula;
                } 
                return _value.GetType(); 
            }
        }


        public override int ColumnIndex
        {
            get
            {
                return _row.GetCellIndex(this);
            }
        }
        /// <summary>
        /// Get DateTime-type cell value
        /// </summary>
        public override DateTime? DateCellValue
        {
            get
            {
                if (CellType != CellType.Numeric && CellType != CellType.Formula)
                {
                    return null;
                }
                double value = NumericCellValue;
                bool date1904 = Sheet.Workbook.IsDate1904();
                return DateUtil.GetJavaDate(value,date1904);
            }
        }
#if NET6_0_OR_GREATER
        /// <summary>
        /// Get DateOnly-type cell value
        /// </summary>
        public override DateOnly? DateOnlyCellValue 
        { 
            get{
                if (CellType != CellType.Numeric && CellType != CellType.Formula)
                {
                    return null;
                }
                double value = NumericCellValue;
                bool date1904 = Sheet.Workbook.IsDate1904();
                return DateOnly.FromDateTime(DateUtil.GetJavaDate(value, date1904));
            }
        }
        public override TimeOnly? TimeOnlyCellValue 
        { 
            get{
                if (CellType != CellType.Numeric && CellType != CellType.Formula)
                {
                    return null;
                }
                double value = NumericCellValue;
                bool date1904 = Sheet.Workbook.IsDate1904();
                return TimeOnly.FromDateTime(DateUtil.GetJavaDate(value, date1904));
            }
        }
#endif
        public override byte ErrorCellValue
        {
            get
            {
                CellType cellType = _value.GetType();
                switch (cellType)
                {
                    case CellType.Blank:
                        return 0;
                    case CellType.Formula:
                        {
                            FormulaValue fv = (FormulaValue)_value;
                            if (fv.GetFormulaType() != CellType.Error)
                                throw typeMismatch(CellType.Error, CellType.Formula, false);
                            return ((ErrorFormulaValue)_value).PreEvaluatedValue;
                        }
                    case CellType.Error:
                        {
                            return ((ErrorValue)_value).Value;
                        }
                    default:
                        throw typeMismatch(CellType.Error, cellType, false);
                }
            }
        }

        public override IHyperlink Hyperlink
        {
            get
            {
                return (IHyperlink)GetPropertyValue(Property.HYPERLINK);
            }

            set
            {
                if (value == null)
                {
                    RemoveHyperlink();
                    return;
                }
                SetProperty(Property.HYPERLINK, value);

                XSSFHyperlink xssfobj = (XSSFHyperlink)value;
                // Assign to us
                CellReference ref1 = new CellReference(RowIndex, ColumnIndex);
                xssfobj.SetCellReference(ref1);

                // Add to the lists
                ((SXSSFSheet)Sheet)._sh.AddHyperlink(xssfobj);
            }
        }

        public override bool IsMergedCell
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public override bool IsPartOfArrayFormulaGroup
        {
            get
            {
                return false;
                throw new NotImplementedException();
            }
        }

        public override double NumericCellValue
        {
            get
            {
                CellType cellType = _value.GetType();
                switch (cellType)
                {
                    case CellType.Blank:
                        return 0.0;
                    case CellType.Formula:
                        {
                            FormulaValue fv = (FormulaValue)_value;
                            if (fv.GetFormulaType() != CellType.Numeric)
                                throw typeMismatch(CellType.Numeric, CellType.Formula, false);
                            return ((NumericFormulaValue)_value).PreEvaluatedValue;
                        }
                    case CellType.Numeric:
                        return ((NumericValue) _value).Value;
                    default:
                        throw typeMismatch(CellType.Numeric, cellType, false);
                }
            }
        }

        public override IRichTextString RichStringCellValue
        {
            get
            {
                CellType cellType = _value.GetType();
                if (cellType != CellType.String)
                    throw typeMismatch(CellType.String, cellType, false);

                StringValue sval = (StringValue)_value;
                if (sval.IsRichText())
                    return ((RichTextValue) _value).Value;
                else
                {
                    string plainText = StringCellValue;
                    return Sheet.Workbook.GetCreationHelper().CreateRichTextString(plainText);
                }
            }
        }

        public override IRow Row
        {
            get { return _row; }
        }

        public override int RowIndex
        {
            get
            {
                return _row.RowNum;
            }
        }

        public override ISheet Sheet
        {
            get
            {
                return _row.Sheet;
            }
        }

        public override string StringCellValue
        {
            get
            {
                CellType cellType = _value.GetType();
                switch (cellType)
                {
                    case CellType.Blank:
                        return "";
                    case CellType.Formula:
                    {
                        FormulaValue fv=(FormulaValue)_value;
                        if(fv.GetFormulaType() !=CellType.String)
                            throw typeMismatch(CellType.String, CellType.Formula, false);
                        if(_value is RichTextStringFormulaValue) {
                            return ((RichTextStringFormulaValue) _value).GetPreEvaluatedValue().String;
                        } else
                        {
                            return ((StringFormulaValue) _value).PreEvaluatedValue;
                        }
                    }
                    case CellType.String:
                    {
                        if (((StringValue) _value).IsRichText())
                            return ((RichTextValue) _value).Value.String;
                        else
                            return ((PlainStringValue) _value).Value;
                    }
                    default:
                        throw typeMismatch(CellType.String, cellType, false);
                }

            }
        }

        public override ICell CopyCellTo(int targetIndex)
        {
            throw new NotImplementedException();
        }

        public override void RemoveCellComment()
        {
            IComment comment = this.CellComment;
            if (comment != null)
            {
                CellAddress ref1 = new CellAddress(RowIndex, ColumnIndex);
                XSSFSheet sh = ((SXSSFSheet)Sheet)._sh;
                sh.GetCommentsTable(false).RemoveComment(ref1);
                sh.GetVMLDrawing(false).RemoveCommentShape(RowIndex, ColumnIndex);
            }

            RemoveProperty(Property.COMMENT);
        }

        //TODO: implement correctly
        public override void RemoveHyperlink()
        {
            RemoveProperty(Property.HYPERLINK);
            ((SXSSFSheet)Sheet)._sh.RemoveHyperlink(RowIndex, ColumnIndex);
        }

        public override void SetAsActiveCell()
        {
            Sheet.ActiveCell = Address;
        }

        [Obsolete]
        public override ICell SetCellErrorValue(byte value)
        {
            // for formulas, we want to keep the type and only have an ERROR as formula value
            if(_value.GetType()==CellType.Formula)
            {
                _value = new ErrorFormulaValue(CellFormula, value);
            }
            else
            {
                _value = new ErrorValue(value);
            }

            return this;
        }

        protected override ICell SetCellFormulaImpl(string formula)
        {
            if(formula == null)
                throw new ArgumentNullException("Argument formula can not be null");

            if(CellType == CellType.Formula)
            {
                ((FormulaValue) _value).Value = formula;
            }
            else
            {
                switch(CellType)
                {
                    case CellType.Numeric:
                        _value = new NumericFormulaValue(formula, NumericCellValue);
                        break;
                    case CellType.String:
                        if(_value is PlainStringValue) {
                            _value = new StringFormulaValue(formula, StringCellValue);
                        } else
                        {
                            if(!(_value is RichTextValue))
                                throw new AssertFailedException("Must be type RichTextValue");
                            _value = new RichTextStringFormulaValue(formula, ((RichTextValue) _value).Value);
                        }
                        break;
                    case CellType.Boolean:
                        _value = new BooleanFormulaValue(formula, BooleanCellValue);
                        break;
                    case CellType.Error:
                        _value = new ErrorFormulaValue(formula, ErrorCellValue);
                        break;
                    case CellType._None:
                    // fall-through
                    case CellType.Formula:
                    // fall-through
                    case CellType.Blank:
                        throw new AssertFailedException("Not support");
                }
            }
            return this;
        }

        protected override ICell SetCellTypeImpl(CellType cellType)
        {
            EnsureType(cellType);
            return this;
        }

        public override ICell SetCellValue(string value)
        {
            if (value != null)
            {
                if (value.Length > SpreadsheetVersion.EXCEL2007.MaxTextLength)
                {
                    throw new ArgumentException("The maximum length of cell contents (text) is 32,767 characters");
                }
                EnsureTypeOrFormulaType(CellType.String);

                if(_value.GetType() == CellType.Formula)
                {
                    ((StringFormulaValue) _value).PreEvaluatedValue = value;
                }
                else
                {
                    ((PlainStringValue) _value).Value = value;
                }
            }
            else
            {
                SetBlank();
            }

            return this;
        }

        public override ICell SetCellValue(bool value)
        {
            EnsureTypeOrFormulaType(CellType.Boolean);
            if (_value.GetType() == CellType.Formula)
                ((BooleanFormulaValue)_value).PreEvaluatedValue = value;
            else
                ((BooleanValue)_value).Value = value;
            return this;
        }

        public override ICell SetCellValue(IRichTextString value)
        {         
            if (value != null && value.String != null)
            {
                if(value.Length > SpreadsheetVersion.EXCEL2007.MaxTextLength)
                {
                    throw new ArgumentException("The maximum length of cell contents (text) is 32,767 characters");
                }
                EnsureRichTextStringType();

                if(_value is RichTextStringFormulaValue)
                {
                    ((RichTextStringFormulaValue) _value).SetPreEvaluatedValue(value);
                } else
                {
                    ((RichTextValue) _value).Value = value;
                }
            }
            else
            {
                SetBlank();
            }

            return this;
        }

        protected override ICell SetCellValueImpl(double value)
        {
            EnsureTypeOrFormulaType(CellType.Numeric);
            if(_value.GetType() == CellType.Formula)
                ((NumericFormulaValue) _value).PreEvaluatedValue = value;
            else
                ((NumericValue) _value).Value = value;
            return this;
        }

        protected override ICell SetCellValueImpl(DateTime value)
        {
            bool date1904 = ((SXSSFWorkbook)Sheet.Workbook).XssfWorkbook.IsDate1904();
            return SetCellValue(DateUtil.GetExcelDate(value, date1904));
        }

#if NET6_0_OR_GREATER
        protected override ICell SetCellValueImpl(DateOnly value)
        {
            bool date1904 = ((SXSSFWorkbook)Sheet.Workbook).XssfWorkbook.IsDate1904();
            return SetCellValue(DateUtil.GetExcelDate(value, date1904));
        }
#endif
        public override string ToString()
        {
            switch (CellType)
            {
                case CellType.Blank:
                    return "";
                case CellType.Boolean:
                    return BooleanCellValue ? "TRUE" : "FALSE";
                case CellType.Error:
                    return ErrorEval.GetText(ErrorCellValue);
                case CellType.Formula:
                    return CellFormula;
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(this))
                    {
                        SimpleDateFormat sdf = new SimpleDateFormat("dd-MMM-yyyy");
                        //sdf.setTimeZone(LocaleUtil.getUserTimeZone());
                        return sdf.Format(DateCellValue);
                    }
                    return NumericCellValue + "";
                case CellType.String:
                    return RichStringCellValue.ToString();
                default:
                    return "Unknown Cell Type: " + CellType;
            }
        }

        private void RemoveProperty(int type)
        {
            Property current = _firstProperty;
            Property previous = null;
            while (current != null && current.GetType() != type)
            {
                previous = current;
                current = current._next;
            }
            if (current != null)
            {
                if (previous != null)
                {
                    previous._next = current._next;
                }
                else
                {
                    _firstProperty = current._next;
                }
            }
        }

        private void SetProperty(int type, object value)
        {
            Property current = _firstProperty;
            Property previous = null;
            while (current != null && current.GetType() != type)
            {
                previous = current;
                current = current._next;
            }
            if (current != null)
            {
                current._value = value;
            }
            else
            {
                switch (type)
                {
                    case Property.COMMENT:
                        {
                            current = new CommentProperty(value);
                            break;
                        }
                    case Property.HYPERLINK:
                        {
                            current = new HyperlinkProperty(value);
                            break;
                        }
                    default:
                        {
                            throw new ArgumentException("Invalid type: " + type);
                        }
                }
                if (previous != null)
                {
                    previous._next = current;
                }
                else
                {
                    _firstProperty = current;
                }
            }
        }

        private object GetPropertyValue(int type)
        {
            return GetPropertyValue(type, null);
        }

        private object GetPropertyValue(int type, string defaultValue)
        {
            Property current = _firstProperty;
            while (current != null && current.GetType() != type) current = current._next;
            return current == null ? defaultValue : current._value;
        }

        private void EnsurePlainStringType()
        {
            if (_value.GetType() != CellType.String
               || ((StringValue)_value).IsRichText())
                _value = new PlainStringValue();
        }

        private void EnsureRichTextStringType()
        {
            // don't change cell type for formulas
            if(_value.GetType() == CellType.Formula)
            {
                String formula = ((FormulaValue)_value).Value;
                _value = new RichTextStringFormulaValue(formula, new XSSFRichTextString(""));
            }
            else if(_value.GetType()!=CellType.String ||
                    !((StringValue) _value).IsRichText())
            {
                _value = new RichTextValue();
            }
        }

        private void EnsureType(CellType type)
        {
            if (_value.GetType() != type)
                SetType(type);
        }

        /*
         * Sets the cell type to type if it is different
         */

        private void EnsureTypeOrFormulaType(CellType type)
        {
            if (_value.GetType() == type)
            {
                if (type == CellType.String && ((StringValue)_value).IsRichText())
                    SetType(CellType.String);
                return;
            }
            if (_value.GetType() == CellType.Formula)
            {
                if (((FormulaValue)_value).GetFormulaType() == type)
                    return;
                switch(type)
                {
                    case CellType.Boolean:
                        _value = new BooleanFormulaValue(CellFormula, false);
                        break;
                    case CellType.Numeric:
                        _value = new NumericFormulaValue(CellFormula, 0);
                        break;
                    case CellType.String:
                        _value = new StringFormulaValue(CellFormula, "");
                        break;
                    case CellType.Error:
                        _value = new ErrorFormulaValue(CellFormula, FormulaError._NO_ERROR.Code);
                        break;
                    default:
                        throw new AssertFailedException("Unknown celltype");
                }
                return;
            }
            SetType(type);
        }

        /*package*/
        private void SetType(CellType type)
        {
            switch (type)
            {
                case CellType.Numeric:
                    {
                        _value = new NumericValue();
                        break;
                    }
                case CellType.String:
                    {
                        PlainStringValue sval = new PlainStringValue();
                        if (_value != null)
                        {
                            // if a cell is not blank then convert the old value to string
                            String str = ConvertCellValueToString();
                            sval.Value = str;
                        }
                        _value = sval;
                        break;
                    }
                case CellType.Formula:
                    {
                        if(CellType == CellType.Blank)
                        {
                            _value = new NumericFormulaValue("", 0);
                        }
                        break;
                    }
                case CellType.Blank:
                    {
                        _value = new BlankValue();
                        break;
                    }
                case CellType.Boolean:
                    {
                        BooleanValue bval = new BooleanValue();
                        if (_value != null)
                        {
                            // if a cell is not blank then convert the old value to string
                            bool val = ConvertCellValueToBoolean();
                            bval.Value = val;
                        }
                        _value = bval;
                        break;
                    }
                case CellType.Error:
                    {
                        _value = new ErrorValue();
                        break;
                    }
                default:
                    {
                        throw new ArgumentException("Illegal type " + type);
                    }
            }
        }

        private static CellType ComputeTypeFromFormula(String formula)
        {
            return CellType.Numeric;
        }

        //COPIED FROM https://svn.apache.org/repos/asf/poi/trunk/src/ooxml/java/org/apache/poi/xssf/usermodel/XSSFCell.java since the functions are declared private there
        /**
         * Used to help format error messages
         */
        private static InvalidOperationException typeMismatch(CellType expectedTypeCode, CellType actualTypeCode, bool isFormulaCell)
        {
            String msg = "Cannot get a " + expectedTypeCode + " value from a " + actualTypeCode
                    + " " + (isFormulaCell ? "formula " : "") + "cell";
            return new InvalidOperationException(msg);
        }
        private bool ConvertCellValueToBoolean()
        {
            CellType cellType = _value.GetType();

            if (cellType == CellType.Formula)
            {
                cellType = CachedFormulaResultType;
            }

            switch (cellType)
            {
                case CellType.Boolean:
                    return BooleanCellValue;
                case CellType.String:

                    String text = StringCellValue;
                    return Boolean.Parse(text);
                case CellType.Numeric:
                    return NumericCellValue != 0;
                case CellType.Error:
                case CellType.Blank:
                    return false;
                default: throw new RuntimeException("Unexpected cell type (" + cellType + ")");
            }

        }
        private String ConvertCellValueToString()
        {
            CellType cellType = _value.GetType();
            return ConvertCellValueToString(cellType);
        }
        private String ConvertCellValueToString(CellType cellType)
        {
            switch (cellType)
            {
                case CellType.Blank:
                    return "";
                case CellType.Boolean:
                    return BooleanCellValue ? "TRUE" : "FALSE";
                case CellType.String:
                    return StringCellValue;
                case CellType.Numeric:
                    return NumericCellValue.ToString();
                case CellType.Error:
                    byte errVal = ErrorCellValue;
                    return FormulaError.ForInt(errVal).String;

                case CellType.Formula:
                    if (_value != null)
                    {
                        FormulaValue fv = (FormulaValue)_value;
                        if (fv.GetFormulaType() != CellType.Formula)
                        {
                            return ConvertCellValueToString(fv.GetFormulaType());
                        }
                    }
                    return "";
                default:
                    throw new InvalidOperationException("Unexpected cell type (" + cellType + ")");
            }
        }

        //END OF COPIED CODE
    }
}
