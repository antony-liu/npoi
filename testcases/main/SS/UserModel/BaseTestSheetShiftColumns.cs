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

namespace TestCases.SS.UserModel
{
    using NPOI.SS;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;

    public abstract class BaseTestSheetShiftColumns
    {
        protected ISheet sheet1;
        protected ISheet sheet2;
        protected IWorkbook workbook;

        protected ITestDataProvider _testDataProvider;

        public virtual void Init()
        {
            int rowIndex = 0;
            sheet1 = workbook.CreateSheet("sheet1");
            IRow row = sheet1.CreateRow(rowIndex++);
            row.CreateCell(0, CellType.Numeric).SetCellValue(0);
            row.CreateCell(1, CellType.Numeric).SetCellValue(1);
            row.CreateCell(2, CellType.Numeric).SetCellValue(2);

            row = sheet1.CreateRow(rowIndex++);
            row.CreateCell(0, CellType.Numeric).SetCellValue(0.1);
            row.CreateCell(1, CellType.Numeric).SetCellValue(1.1);
            row.CreateCell(2, CellType.Numeric).SetCellValue(2.1);
            row = sheet1.CreateRow(rowIndex++);
            row.CreateCell(0, CellType.Numeric).SetCellValue(0.2);
            row.CreateCell(1, CellType.Numeric).SetCellValue(1.2);
            row.CreateCell(2, CellType.Numeric).SetCellValue(2.2);
            row = sheet1.CreateRow(rowIndex++);
            row.CreateCell(0, CellType.Formula).SetCellFormula("A2*B3");
            row.CreateCell(1, CellType.Numeric).SetCellValue(1.3);
            row.CreateCell(2, CellType.Formula).SetCellFormula("B1-B3");
            row = sheet1.CreateRow(rowIndex++);
            row.CreateCell(0, CellType.Formula).SetCellFormula("SUM(C1:C4)");
            row.CreateCell(1, CellType.Formula).SetCellFormula("SUM(A3:C3)");
            row.CreateCell(2, CellType.Formula).SetCellFormula("$C1+C$2");
            row = sheet1.CreateRow(rowIndex++);
            row.CreateCell(1, CellType.Numeric).SetCellValue(1.5);
            row = sheet1.CreateRow(rowIndex);
            row.CreateCell(1, CellType.Boolean).SetCellValue(false);
            ICell textCell =  row.CreateCell(2, CellType.String);
            textCell.SetCellValue("TEXT");
            textCell.CellStyle = newCenterBottomStyle();

            sheet2 = workbook.CreateSheet("sheet2");
            row = sheet2.CreateRow(0);
            row.CreateCell(0, CellType.Numeric).SetCellValue(10);
            row.CreateCell(1, CellType.Numeric).SetCellValue(11);
            row.CreateCell(2, CellType.Formula).SetCellFormula("SUM(sheet1!B3:C3)");
            row = sheet2.CreateRow(1);
            row.CreateCell(0, CellType.Numeric).SetCellValue(21);
            row.CreateCell(1, CellType.Numeric).SetCellValue(22);
            row.CreateCell(2, CellType.Numeric).SetCellValue(23);
            row = sheet2.CreateRow(2);
            row.CreateCell(0, CellType.Formula).SetCellFormula("sheet1!A4+sheet1!C2+A2");
            row.CreateCell(1, CellType.Formula).SetCellFormula("SUM(sheet1!A3:$C3)");
            row = sheet2.CreateRow(3);
            row.CreateCell(0, CellType.String).SetCellValue("dummy");
        }

        private ICellStyle newCenterBottomStyle()
        {
            ICellStyle style = workbook.CreateCellStyle();
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Bottom;
            return style;
        }
        [Test]
        public virtual void TestShiftOneColumnRight()
        {
            sheet1.ShiftColumns(1, 2, 1);
            double c1Value = sheet1.GetRow(0).GetCell(2).NumericCellValue;
            ClassicAssert.AreEqual(1d, c1Value, 0.01);
            String formulaA4 = sheet1.GetRow(3).GetCell(0).CellFormula;
            ClassicAssert.AreEqual("A2*C3", formulaA4);
            String formulaC4 = sheet1.GetRow(3).GetCell(3).CellFormula;
            ClassicAssert.AreEqual("C1-C3", formulaC4);
            String formulaB5 = sheet1.GetRow(4).GetCell(2).CellFormula;
            ClassicAssert.AreEqual("SUM(A3:D3)", formulaB5);
            String formulaD5 = sheet1.GetRow(4).GetCell(3).CellFormula; // $C1+C$2
            ClassicAssert.AreEqual("$D1+D$2", formulaD5);

            ICell newb5Null = sheet1.GetRow(4).GetCell(1);
            ClassicAssert.AreEqual(newb5Null, null);
            bool logicalValue = sheet1.GetRow(6).GetCell(2).BooleanCellValue;
            ClassicAssert.AreEqual(logicalValue, false);
            ICell textCell = sheet1.GetRow(6).GetCell(3);
            ClassicAssert.AreEqual(textCell.StringCellValue, "TEXT");
            ClassicAssert.AreEqual(textCell.CellStyle.Alignment, HorizontalAlignment.Center);

            // other sheet
            String formulaC1 = sheet2.GetRow(0).GetCell(2).CellFormula; // SUM(sheet1!B3:C3)
            ClassicAssert.AreEqual("SUM(sheet1!C3:D3)", formulaC1);
            String formulaA3 = sheet2.GetRow(2).GetCell(0).CellFormula; // sheet1!A4+sheet1!C2+A2
            ClassicAssert.AreEqual("sheet1!A4+sheet1!D2+A2", formulaA3);
        }

        [Test]
        public virtual void TestShiftTwoColumnsRight()
        {
            sheet1.ShiftColumns(1, 2, 2);
            String formulaA4 = sheet1.GetRow(3).GetCell(0).CellFormula;
            ClassicAssert.AreEqual("A2*D3", formulaA4);
            String formulaD4 = sheet1.GetRow(3).GetCell(4).CellFormula;
            ClassicAssert.AreEqual("D1-D3", formulaD4);
            String formulaD5 = sheet1.GetRow(4).GetCell(3).CellFormula;
            ClassicAssert.AreEqual("SUM(A3:E3)", formulaD5);

            ICell b5Null = sheet1.GetRow(4).GetCell(1);
            ClassicAssert.AreEqual(b5Null, null);
            object c6Null = sheet1.GetRow(5).GetCell(2); // null cell A5 is shifted
                                                         // for 2 columns, so now
                                                         // c5 should be null
            ClassicAssert.AreEqual(c6Null, null);
        }

        [Test]
        public virtual void TestShiftOneColumnLeft()
        {
            sheet1.ShiftColumns(1, 2, -1);

            String formulaA5 = sheet1.GetRow(4).GetCell(0).CellFormula;
            ClassicAssert.AreEqual("SUM(A3:B3)", formulaA5);
            String formulaB4 = sheet1.GetRow(3).GetCell(1).CellFormula;
            ClassicAssert.AreEqual("A1-A3", formulaB4);
            String formulaB5 = sheet1.GetRow(4).GetCell(1).CellFormula;
            ClassicAssert.AreEqual("$B1+B$2", formulaB5);
            ICell newb6Null = sheet1.GetRow(5).GetCell(1);
            ClassicAssert.AreEqual(newb6Null, null);
        }

        [Test]
        [Ignore("XSSFSheet.ShiftColumns of POI failed")]
        public virtual void TestShiftTwoColumnsLeft()
        {
            ClassicAssert.Throws<InvalidOperationException>(() =>
            {
                sheet1.ShiftColumns(1, 2, -2);
            });
            
        }

        [Test]
        public virtual void TestShiftHyperlinks()
        {

            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet("test");
            IRow row = sheet.CreateRow(0);

            // How to create hyperlinks
            // https://poi.apache.org/spreadsheet/quick-guide.html#Hyperlinks
            ICreationHelper helper = wb.GetCreationHelper();
            ICellStyle hlinkStyle = wb.CreateCellStyle();
            IFont hlinkFont = wb.CreateFont();
            hlinkFont.Underline = FontUnderlineType.Single; // (Font.U_SINGLE);
            hlinkFont.Color = IndexedColors.Blue.Index;
            hlinkStyle.SetFont(hlinkFont);

            // 3D relative document link
            // CellAddress=A1, shifted to A4
            ICell cell = row.CreateCell(0);
            cell.CellStyle = hlinkStyle;
            createHyperlink(helper, cell, HyperlinkType.Document, "test!E1");

            // URL
            cell = row.CreateCell(1);
            // CellAddress=B1, shifted to B4
            cell.CellStyle = hlinkStyle;
            createHyperlink(helper, cell, HyperlinkType.Url, "http://poi.apache.org/");

            // row0 will be shifted on top of row1, so this URL should be removed
            // from the workbook
            IRow overwrittenRow = sheet.CreateRow(3);
            cell = overwrittenRow.CreateCell(2);
            // CellAddress=C4, will be overwritten (deleted)
            cell.CellStyle = hlinkStyle;
            createHyperlink(helper, cell, HyperlinkType.Email, "mailto:poi@apache.org");

            IRow unaffectedRow = sheet.CreateRow(20);
            cell = unaffectedRow.CreateCell(3);
            // CellAddress=D21, will be unaffected
            cell.CellStyle = hlinkStyle;
            createHyperlink(helper, cell, HyperlinkType.File, "54524.xlsx");

            cell = wb.CreateSheet("other").CreateRow(0).CreateCell(0);
            // CellAddress=Other!A1, will be unaffected
            cell.CellStyle = hlinkStyle;
            createHyperlink(helper, cell, HyperlinkType.Url, "http://apache.org/");

            int startRow = 0;
            int endRow = 4;
            int n = 3;
            sheet.ShiftColumns(startRow, endRow, n);

            IWorkbook read = _testDataProvider.WriteOutAndReadBack(wb);
            wb.Close();

            ISheet sh = read.GetSheet("test");

            IRow shiftedRow = sh.GetRow(0);

            // document link anchored on a shifted cell should be Moved
            // Note that hyperlinks do not track what they point to, so this
            // hyperlink should still refer to test!E1
            verifyHyperlink(shiftedRow.GetCell(3), HyperlinkType.Document, "test!E1");

            // URL, EMAIL, and FILE links anchored on a shifted cell should be Moved
            verifyHyperlink(shiftedRow.GetCell(4), HyperlinkType.Url, "http://poi.apache.org/");

            // Make sure hyperlinks were Moved and not copied
            ClassicAssert.IsNull(sh.GetHyperlink(0, 0), "Document hyperlink should be Moved, not copied");
            ClassicAssert.IsNull(sh.GetHyperlink(1, 0), "URL hyperlink should be Moved, not copied");

            ClassicAssert.AreEqual(4, sh.GetHyperlinkList().Count);
            read.Close();
        }

        private void createHyperlink(ICreationHelper helper, ICell cell, HyperlinkType linkType, String ref1)
        {
            cell.SetCellValue(ref1);
            IHyperlink link = helper.CreateHyperlink(linkType);
            link.Address = ref1;
            cell.Hyperlink = link;
        }

        private void verifyHyperlink(ICell cell, HyperlinkType linkType, String ref1)
        {
            ClassicAssert.IsTrue(cellHasHyperlink(cell));
            if(cell != null)
            {
                IHyperlink link = cell.Hyperlink;
                ClassicAssert.AreEqual(linkType, link.Type);
                ClassicAssert.AreEqual(ref1, link.Address);
            }
        }

        private bool cellHasHyperlink(ICell cell)
        {
            return (cell != null) && (cell.Hyperlink != null);
        }

        [Test]
        public virtual void ShiftMergedColumnsToMergedColumnsRight()
        {

            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet("test");

            // populate sheet cells
            populateSheetCells(sheet);
            CellRangeAddress A1_A5 = new CellRangeAddress(0, 4, 0, 0); // NOSONAR, it's more readable this way
            CellRangeAddress B1_B3 = new CellRangeAddress(0, 2, 1, 1); // NOSONAR, it's more readable this way

            sheet.AddMergedRegion(B1_B3);
            sheet.AddMergedRegion(A1_A5);

            // A1:A5 should be Moved to B1:B5
            // B1:B3 will be removed
            sheet.ShiftColumns(0, 0, 1);
            ClassicAssert.AreEqual(1, sheet.NumMergedRegions);
            ClassicAssert.AreEqual(CellRangeAddress.ValueOf("B1:B5"), sheet.GetMergedRegion(0));

            wb.Close();
        }

        [Test]
        [Ignore("XSSFSheet.ShiftColumns of POI failed")]
        public virtual void ShiftMergedColumnsToMergedColumnsLeft()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet("test");
            populateSheetCells(sheet);

            CellRangeAddress A1_A5 = new CellRangeAddress(0, 4, 0, 0);  // NOSONAR, it's more readable this way
            CellRangeAddress B1_B3 = new CellRangeAddress(0, 2, 1, 1);  // NOSONAR, it's more readable this way

            sheet.AddMergedRegion(A1_A5);
            sheet.AddMergedRegion(B1_B3);

            // A1:E1 should be removed
            // B1:B3 will be A1:A3
            sheet.ShiftColumns(1, 5, -1);

            ClassicAssert.AreEqual(1, sheet.NumMergedRegions);
            ClassicAssert.AreEqual(CellRangeAddress.ValueOf("A1:A3"), sheet.GetMergedRegion(0));

            wb.Close();
        }

        private void populateSheetCells(ISheet sheet)
        {
            // populate sheet cells
            for(int i = 0; i < 2; i++)
            {
                IRow row = sheet.CreateRow(i);
                for(int j = 0; j < 5; j++)
                {
                    ICell cell = row.CreateCell(j);
                    cell.SetCellValue(i + "x" + j);
                }
            }
        }

        [Test]
        public virtual void TestShiftWithMergedRegions()
        {

            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);
            row.CreateCell(0).SetCellValue(1.1);
            row = sheet.CreateRow(1);
            row.CreateCell(0).SetCellValue(2.2);
            CellRangeAddress region = new CellRangeAddress(0, 2, 0, 0);
            ClassicAssert.AreEqual("A1:A3", region.FormatAsString());

            sheet.AddMergedRegion(region);

            sheet.ShiftColumns(0, 1, 2);
            region = sheet.GetMergedRegion(0);
            ClassicAssert.AreEqual("C1:C3", region.FormatAsString());
            wb.Close();
        }

        protected abstract IWorkbook openWorkbook(String spreadsheetFileName);

        protected abstract IWorkbook GetReadBackWorkbook(IWorkbook wb);
    


        protected static  String AMDOCS = "Amdocs";
        protected static  String AMDOCS_TEST = "Amdocs:\ntest\n";

        [Test]
        public virtual void TestCommentsShifting()
        {

            IWorkbook inputWb = openWorkbook("56017.xlsx");

            ISheet sheet = inputWb.GetSheetAt(0);
            IComment comment = sheet.GetCellComment(new CellAddress(0, 0));
            ClassicAssert.IsNotNull(comment);
            ClassicAssert.AreEqual(AMDOCS, comment.Author);
            ClassicAssert.AreEqual(AMDOCS_TEST, comment.String.String);

            sheet.ShiftColumns(0, 1, 1);

            // comment in column 0 is gone
            comment = sheet.GetCellComment(new CellAddress(0, 0));
            ClassicAssert.IsNull(comment);

            // comment is column in column 1
            comment = sheet.GetCellComment(new CellAddress(0, 1));
            ClassicAssert.IsNotNull(comment);
            ClassicAssert.AreEqual(AMDOCS, comment.Author);
            ClassicAssert.AreEqual(AMDOCS_TEST, comment.String.String);

            IWorkbook wbBack = GetReadBackWorkbook(inputWb);
            inputWb.Close();
            ClassicAssert.IsNotNull(wbBack);

            ISheet sheetBack = wbBack.GetSheetAt(0);

            // comment in column 0 is gone
            comment = sheetBack.GetCellComment(new CellAddress(0, 0));
            ClassicAssert.IsNull(comment);

            // comment is now in column 1
            comment = sheetBack.GetCellComment(new CellAddress(0, 1));
            ClassicAssert.IsNotNull(comment);
            ClassicAssert.AreEqual(AMDOCS, comment.Author);
            ClassicAssert.AreEqual(AMDOCS_TEST, comment.String.String);
            wbBack.Close();
        }

        // transposed version of TestXSSFSheetShiftRows.testBug54524()
        [Test]
        public virtual void TestBug54524()
        {

            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet();
            IRow firstRow = sheet.CreateRow(0);
            firstRow.CreateCell(0).SetCellValue("");
            firstRow.CreateCell(1).SetCellValue(1);
            firstRow.CreateCell(2).SetCellValue(2);
            firstRow.CreateCell(3).SetCellFormula("SUM(B1:C1)");
            firstRow.CreateCell(4).SetCellValue("X");

            sheet.ShiftColumns(3, 5, -1);

            ICell cell = CellUtil.GetCell(sheet.GetRow(0), 1);
            ClassicAssert.AreEqual(1.0, cell.NumericCellValue, 0);
            cell = CellUtil.GetCell(sheet.GetRow(0), 2);
            ClassicAssert.AreEqual("SUM(B1:B1)", cell.CellFormula);
            cell = CellUtil.GetCell(sheet.GetRow(0), 3);
            ClassicAssert.AreEqual("X", cell.StringCellValue);
            wb.Close();
        }
    }
}

