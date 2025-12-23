/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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


using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.SS.UserModel
{
    using NPOI.HPSF;
    using NPOI.SS.Formula.Functions;
    using NPOI.SS.Util;
    using NPOI.Util;
    /// <summary>
    /// Common implementation-independent logic shared by all implementations of <see cref="Cell"/>.
    /// </summary>
    /// <remarks>
    /// @author Vladislav "gallon" Galas gallon at apache dot org
    /// </remarks>
    public abstract class CellBase : ICell
    {
        public abstract int ColumnIndex { get; }
        public abstract int RowIndex { get; }
        public abstract ISheet Sheet { get; }
        public abstract IRow Row { get; }
        public abstract CellType CellType { get; }
        public abstract CellType CachedFormulaResultType { get; }
        public abstract CellRangeAddress ArrayFormulaRange { get; }
        public abstract bool IsPartOfArrayFormulaGroup { get; }
        public abstract string CellFormula { get; set; }

        public abstract double NumericCellValue { get; }

        public abstract DateTime? DateCellValue { get; }
#if NET6_0_OR_GREATER
        public abstract DateOnly? DateOnlyCellValue { get; }

        public abstract TimeOnly? TimeOnlyCellValue { get; }
#endif
        public abstract IRichTextString RichStringCellValue { get; }

        public abstract byte ErrorCellValue { get; }

        public abstract string StringCellValue { get; }

        public abstract bool BooleanCellValue { get; }

        public abstract ICellStyle CellStyle { get; set; }

        public CellAddress Address 
        {
            get
            {
                return new CellAddress(this);
            }
        }

        public abstract IComment CellComment { get; set; }
        public abstract IHyperlink Hyperlink { get; set; }

        public abstract bool IsMergedCell { get; }
        /// <summary>
        /// {@inheritDoc}
        /// </summary>
        public ICell SetCellType(CellType cellType)
        {
            if(cellType == CellType._None)
            {
                throw new ArgumentException("cellType shall not be null nor _NONE");
            }
            if(cellType == CellType.Formula)
            {
                if(CellType != CellType.Formula)
                {
                    throw new ArgumentException("Calling Cell.setCellType(CellType.FORMULA) is illegal. " +
                            "Use setCellFormula(String) directly.");
                }
                else
                {
                    return this;
                }
            }
            TryToDeleteArrayFormulaIfSet();

            return SetCellTypeImpl(cellType);
        }

        /// <summary>
        /// Implementation-specific logic
        /// </summary>
        /// <param name="cellType">new cell type. Guaranteed non-null, not _NONE.</param>
        protected abstract ICell SetCellTypeImpl(CellType cellType);

        /// <summary>
        /// <para>
        /// Called when this an array formula in this cell is deleted.
        /// </para>
        /// <para>
        /// The purpose of this method is to validate the cell state prior to modification.
        /// </para>
        /// </summary>
        /// <param name="message">a customized exception message for the case if deletion of the cell is impossible. If null, a
        /// default message will be generated
        /// </param>
        /// @see #setCellType(CellType)
        /// @see #setCellFormula(String)
        /// @see Row#removeCell(NPOI.SS.UserModel.Cell)
        /// @see NPOI.SS.UserModel.Sheet#removeRow(NPOI.SS.UserModel.Row)
        /// @see NPOI.SS.UserModel.Sheet#shiftRows(int, int, int)
        /// @see NPOI.SS.UserModel.Sheet#addMergedRegion(NPOI.SS.Util.CellRangeAddress)
        /// <exception cref="InvalidOperationException">if modification is not allowed</exception>
        /// 
        /// Note. Exposing this to public is ugly. Needed for methods like Sheet#shiftRows.
        public void TryToDeleteArrayFormula(string message)
        {
            if(!IsPartOfArrayFormulaGroup)
                return;

            CellRangeAddress arrayFormulaRange = ArrayFormulaRange;
            if(arrayFormulaRange.NumberOfCells > 1)
            {
                if(message == null)
                {
                    message = "Cell " + new CellReference(this).FormatAsString() + " is part of a multi-cell array formula. " +
                            "You cannot change part of an array.";
                }
                throw new InvalidOperationException(message);
            }
            //un-register the single-cell array formula from the parent sheet through public interface
            Row.Sheet.RemoveArrayFormula(this);
        }

        /// <summary>
        /// {@inheritDoc}
        /// </summary>
        public ICell SetCellFormula(string formula)
        {
            // todo validate formula here, before changing the cell?
            TryToDeleteArrayFormulaIfSet();

            if(formula == null)
            {
                RemoveFormula();
                return this;
            }

            // formula cells always have a value. If the cell is blank (either initially or after removing an
            // array formula), set value to 0
            if(GetValueType() == CellType.Blank)
            {
                SetCellValue(0);
            }

            return SetCellFormulaImpl(formula);
        }

        /// <summary>
        /// Implementation-specific Setting the formula.
        /// Shall not change the value.
        /// </summary>
        /// <param name="formula"></param>
        protected abstract ICell SetCellFormulaImpl(string formula);

        /**
         * Get value type of this cell. Can return BLANK, NUMERIC, STRING, BOOLEAN or ERROR.
         * For current implementations where type is strongly coupled with formula, is equivalent to
         * <code>getCellType() == CellType.FORMULA ? getCachedFormulaResultType() : getCellType()</code>
         *
         * <p>This is meant as a temporary helper method until the time when value type is decoupled from the formula.</p>
         * @return value type
         */
        protected CellType GetValueType()
        {
            CellType type = CellType;
            if(type != CellType.Formula)
            {
                return type;
            }
            return CachedFormulaResultType;
        }
        /// <summary>
        /// {@inheritDoc}
        /// </summary>
        public void RemoveFormula()
        {
            if(CellType == CellType.Blank)
            {
                return;
            }

            if(IsPartOfArrayFormulaGroup)
            {
                TryToDeleteArrayFormula(null);
                return;
            }

            RemoveFormulaImpl();
        }

        /// <summary>
        /// Implementation-specific removal of the formula.
        /// The cell is guaranteed to have a regular formula Set.
        /// Shall preserve the "cached" value.
        /// </summary>
        protected abstract void RemoveFormulaImpl();

        private void TryToDeleteArrayFormulaIfSet()
        {
            if(IsPartOfArrayFormulaGroup)
            {
                TryToDeleteArrayFormula(null);
            }
        }

        [Obsolete("deprecated 4.1. Use {@link #setCellErrorValue(FormulaError)} instead.")]
        [Removal(Version = "4.2")]
        public abstract ICell SetCellErrorValue(byte value);

        public abstract ICell CopyCellTo(int targetIndex);

        public abstract ICell SetCellValue(bool value);

        public abstract void SetAsActiveCell();

        public abstract void RemoveCellComment();

        public abstract void RemoveHyperlink();

        [Obsolete("Will be removed at NPOI 2.8, Use CachedFormulaResultType instead.")]
        [Removal(Version = "4.2")]
        public abstract CellType GetCachedFormulaResultTypeEnum();

        public virtual ICell SetBlank()
        {
            SetCellType(CellType.Blank);
            return this;
        }

        public ICell SetCellValue(double value)
        {
            if(Double.IsInfinity(value))
            {
                // Excel does not support positive/negative infinities,
                // rather, it gives a #DIV/0! error in these cases.
                SetCellErrorValue(FormulaError.DIV0.Code);
            }
            else if(Double.IsNaN(value))
            {
                SetCellErrorValue(FormulaError.NUM.Code);
            }
            else
            {
                SetCellValueImpl(value);
            }
            return this;
        }

        /**
         * Implementation-specific way to set a numeric value.
         * <code>value</code> is guaranteed to be a valid (non-NaN) double.
         * The implementation is expected to adjust the cell type accordingly, so that after this call
         * getCellType() or getCachedFormulaResultType() would return {@link CellType#NUMERIC}.
         * @param value the new value to set
         */
        protected abstract ICell SetCellValueImpl(double value);

        public ICell SetCellValue(DateTime value)
        {
            return SetCellValueImpl(value);
        }

        public ICell SetCellValue(DateTime? value)
        {
            if(value == null)
            {
                SetBlank();
                return this;
            }
            return SetCellValue(value.Value);
        }

        protected abstract ICell SetCellValueImpl(DateTime value);
#if NET6_0_OR_GREATER
        public ICell SetCellValue(DateOnly? value)
        {
            if (value == null)
            {
                SetBlank();
                return this;
            }
            
            return SetCellValue(value.Value);
        }
        public ICell SetCellValue(DateOnly value)
        {
            return SetCellValueImpl(value);
        }
        protected abstract ICell SetCellValueImpl(DateOnly value);
#endif
        public abstract ICell SetCellValue(IRichTextString value);

        public abstract ICell SetCellValue(string value);
    }
}


