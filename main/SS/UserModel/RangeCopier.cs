/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */


using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.SS.UserModel
{
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.Util;
    public abstract class RangeCopier
    {
        private ISheet sourceSheet;
        private ISheet destSheet;
        private FormulaShifter horizontalFormulaShifter;
        private FormulaShifter verticalFormulaShifter;

        public RangeCopier(ISheet sourceSheet, ISheet destSheet)
        {
            this.sourceSheet = sourceSheet;
            this.destSheet = destSheet;
        }
        public RangeCopier(ISheet sheet)
                : this(sheet, sheet)
        {

        }
        /// <summary>
        /// Uses input pattern to tile destination region, overwriting existing content. Works in following manner :
        /// 1.Start from top-left of destination.
        /// 2.Paste source but only inside of destination borders.
        /// 3.If there is space left on right or bottom side of copy, process it as in step 2.
        /// </summary>
        /// <param name="tilePatternRange">source range which should be copied in tiled manner</param>
        /// <param name="tileDestRange">    destination range, which should be overridden</param>
        public void CopyRange(CellRangeAddress tilePatternRange, CellRangeAddress tileDestRange)
        {
            ISheet sourceCopy = sourceSheet.Workbook.CloneSheet(sourceSheet.Workbook.GetSheetIndex(sourceSheet));
            int sourceWidthMinus1 = tilePatternRange.LastColumn - tilePatternRange.FirstColumn;
            int sourceHeightMinus1 = tilePatternRange.LastRow - tilePatternRange.FirstRow;
            int rightLimitToCopy;
            int bottomLimitToCopy;

            int nextRowIndexToCopy = tileDestRange.FirstRow;
            do
            {
                int nextCellIndexInRowToCopy = tileDestRange.FirstColumn;
                int heightToCopyMinus1 = Math.Min(sourceHeightMinus1, tileDestRange.LastRow - nextRowIndexToCopy);
                bottomLimitToCopy = tilePatternRange.FirstRow + heightToCopyMinus1;
                do
                {
                    int widthToCopyMinus1 = Math.Min(sourceWidthMinus1, tileDestRange.LastColumn - nextCellIndexInRowToCopy);
                    rightLimitToCopy = tilePatternRange.FirstColumn + widthToCopyMinus1;
                    CellRangeAddress rangeToCopy = new CellRangeAddress(
                        tilePatternRange.FirstRow,     bottomLimitToCopy,
                        tilePatternRange.FirstColumn,  rightLimitToCopy
                       );
                    CopyRange(rangeToCopy, nextCellIndexInRowToCopy - rangeToCopy.FirstColumn, nextRowIndexToCopy - rangeToCopy.FirstRow, sourceCopy);
                    nextCellIndexInRowToCopy += widthToCopyMinus1 + 1;
                } while(nextCellIndexInRowToCopy <= tileDestRange.LastColumn);
                nextRowIndexToCopy += heightToCopyMinus1 + 1;
            } while(nextRowIndexToCopy <= tileDestRange.LastRow);

            int tempCopyIndex = sourceSheet.Workbook.GetSheetIndex(sourceCopy);
            sourceSheet.Workbook.RemoveSheetAt(tempCopyIndex);
        }

        private void CopyRange(CellRangeAddress sourceRange, int deltaX, int deltaY, ISheet sourceClone)
        { //NOSONAR, it's a bit complex but monolith method, does not make much sense to divide it
            if(deltaX != 0)
                horizontalFormulaShifter = FormulaShifter.CreateForColumnCopy(sourceSheet.Workbook.GetSheetIndex(sourceSheet),
                        sourceSheet.SheetName, sourceRange.FirstColumn, sourceRange.LastColumn, deltaX, sourceSheet.Workbook.SpreadsheetVersion);
            if(deltaY != 0)
                verticalFormulaShifter = FormulaShifter.CreateForRowCopy(sourceSheet.Workbook.GetSheetIndex(sourceSheet),
                        sourceSheet.SheetName, sourceRange.FirstRow, sourceRange.LastRow, deltaY, sourceSheet.Workbook.SpreadsheetVersion);

            for(int rowNo = sourceRange.FirstRow; rowNo <= sourceRange.LastRow; rowNo++)
            {
                IRow sourceRow = sourceClone.GetRow(rowNo); // copy from source copy, original source might be overridden in process!
                for(int columnIndex = sourceRange.FirstColumn; columnIndex <= sourceRange.LastColumn; columnIndex++)
                {
                    ICell sourceCell = sourceRow.GetCell(columnIndex);
                    if(sourceCell == null)
                        continue;
                    IRow destRow = destSheet.GetRow(rowNo + deltaY);
                    if(destRow == null)
                        destRow = destSheet.CreateRow(rowNo + deltaY);

                    ICell newCell = destRow.GetCell(columnIndex + deltaX);
                    if(newCell == null)
                        newCell = destRow.CreateCell(columnIndex + deltaX, sourceCell.CellType);

                    cloneCellContent(sourceCell, newCell, null);
                    if(newCell.CellType == CellType.Formula)
                        AdjustCellReferencesInsideFormula(newCell, destSheet, deltaX, deltaY);
                }
            }
        }

        protected abstract void AdjustCellReferencesInsideFormula(ICell cell, ISheet destSheet, int deltaX, int deltaY); // this part is different for HSSF and XSSF

        protected bool adjustInBothDirections(Ptg[] ptgs, int sheetIndex, int deltaX, int deltaY)
        {
            bool adjustSucceeded = true;
            if(deltaY != 0)
                adjustSucceeded = verticalFormulaShifter.AdjustFormula(ptgs, sheetIndex);
            if(deltaX != 0)
                adjustSucceeded = adjustSucceeded && horizontalFormulaShifter.AdjustFormula(ptgs, sheetIndex);
            return adjustSucceeded;
        }

        // TODO clone some more properties ? 
        public static void cloneCellContent(ICell srcCell, ICell destCell, Dictionary<int, ICellStyle> styleMap)
        {
            if(styleMap != null)
            {
                if(srcCell.Sheet.Workbook == destCell.Sheet.Workbook)
                {
                    destCell.CellStyle = srcCell.CellStyle;
                }
                else
                {
                    int stHashCode = srcCell.CellStyle.GetHashCode();
                    ICellStyle newCellStyle = styleMap[stHashCode];
                    if(newCellStyle == null)
                    {
                        newCellStyle = destCell.Sheet.Workbook.CreateCellStyle();
                        newCellStyle.CloneStyleFrom(srcCell.CellStyle);
                        styleMap[stHashCode] = newCellStyle;
                    }
                    destCell.CellStyle = newCellStyle;
                }
            }
            switch(srcCell.CellType)
            {
                case CellType.String:
                    destCell.SetCellValue(srcCell.StringCellValue);
                    break;
                case CellType.Numeric:
                    destCell.SetCellValue(srcCell.NumericCellValue);
                    break;
                case CellType.Blank:
                    destCell.SetBlank();
                    break;
                case CellType.Boolean:
                    destCell.SetCellValue(srcCell.BooleanCellValue);
                    break;
                case CellType.Error:
                    destCell.SetCellErrorValue(srcCell.ErrorCellValue);
                    break;
                case CellType.Formula:
                    string oldFormula = srcCell.CellFormula;
                    destCell.SetCellFormula(oldFormula);
                    break;
                default:
                    break;
            }
        }
    }
}
