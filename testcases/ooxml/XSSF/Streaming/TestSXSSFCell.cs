/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for Additional information regarding copyright ownership.
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

namespace TestCases.XSSF.Streaming
{
    using MathNet.Numerics;
    using NPOI.OpenXmlFormats.Spreadsheet;
    using NPOI.SS;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.XSSF;
    using NPOI.XSSF.Streaming;
    using NPOI.XSSF.UserModel;
    using NSubstitute;
    using NUnit.Framework;using NUnit.Framework.Legacy;
    using System;
    using System.Text;
    using TestCases.SS.UserModel;

    /**
     * Tests various functionality having to do with {@link SXSSFCell}.  For instance support for
     * particular datatypes, etc.
     */
    //[Ignore("This may cause file access denied")]
    public class TestSXSSFCell : BaseTestXCell
    {

        public TestSXSSFCell()
            : base(SXSSFITestDataProvider.instance)
        {

        }

        [TearDown]
        public static void TearDown() {
            //SXSSFITestDataProvider.instance.Cleanup();
        }

        [Test]
        public void TestPreserveSpaces()
        {
            String[] samplesWithSpaces = {
                " POI",
                "POI ",
                " POI ",
                "\nPOI",
                "\n\nPOI \n",
            };
            foreach (String str in samplesWithSpaces) 
            {
                using(var swb = _testDataProvider.CreateWorkbook())
                {
                    ICell sCell = swb.CreateSheet().CreateRow(0).CreateCell(0);
                    sCell.SetCellValue(str);
                    ClassicAssert.AreEqual(sCell.StringCellValue, str);

                    // read back as XSSF and check that xml:spaces="preserve" is Set
                    using(var xwb = (XSSFWorkbook) _testDataProvider.WriteOutAndReadBack(swb))
                    {
                        XSSFCell xCell = xwb.GetSheetAt(0).GetRow(0).GetCell(0) as XSSFCell;

                        CT_Rst is1 = xCell.GetCTCell().@is;
                        ClassicAssert.IsNotNull(is1);
                        //XmlCursor c = is1.NewCursor();
                        //c.ToNextToken();
                        //String t = c.GetAttributeText(new QName("http://www.w3.org/XML/1998/namespace", "space"));
                        //c.Dispose();

                        ClassicAssert.IsTrue(is1.XmlText.Contains("xml:space=\"preserve\""));
                        //write is1 to xml stream writer ,get the xml text and parse it and get space attr.
                        //ClassicAssert.AreEqual("preserve", t, "expected xml:spaces=\"preserve\" \"" + str + "\"");
                    }
                }
            }
        }

        [Test]
        public void GetCellTypeEnumDelegatesToGetCellType()
        {
            var instance = Substitute.ForPartsOf<SXSSFCell>(null, CellType.Blank);
            //SXSSFCell instance = spy(new SXSSFCell(null, CellType.Blank));
            CellType result = instance.CellType;
            _ = instance.Received().CellType;
            ClassicAssert.AreEqual(CellType.Blank, result);
        }

        [Test]
        public void GetCachedFormulaResultTypeEnum_delegatesTo_getCachedFormulaResultType()
        {
            var instance = Substitute.ForPartsOf<SXSSFCell>(null, CellType.Blank);
            instance.SetCellFormula("");
            _ = instance.CachedFormulaResultType;
            _ = instance.Received().CachedFormulaResultType;
        }

        [Test]
        public void GetCachedFormulaResultType_throwsISE_whenNotAFormulaCell()
        {
            Assert.Throws<InvalidOperationException>(() => {
                SXSSFCell instance = new SXSSFCell(null, CellType.Blank);
                _ = instance.CachedFormulaResultType;
            });
            
        }


        [Test]
        public void SetCellValue_withTooLongRichTextString_throwsIAE()
        {
            Assert.Throws<ArgumentException>(() =>
            {
                ICell cell = Substitute.ForPartsOf<SXSSFCell>();
                cell.CellType.Returns(CellType.Blank);
                int length = SpreadsheetVersion.EXCEL2007.MaxTextLength + 1;
                string string1 = Encoding.UTF8.GetString(new byte[length]).Replace("\0", "x");
                IRichTextString richTextString = new XSSFRichTextString(string1);
                cell.SetCellValue(richTextString);
            });
        }

        [Test]
        public void GetArrayFormulaRange_returnsNull()
        {
            ICell cell = new SXSSFCell(null, CellType.Blank);
            CellRangeAddress result = cell.ArrayFormulaRange;
            ClassicAssert.IsNull(result);
        }

        [Test]
        public void IsPartOfArrayFormulaGroup_returnsFalse()
        {
            ICell cell = new SXSSFCell(null, CellType.Blank);
            bool result = cell.IsPartOfArrayFormulaGroup;
            ClassicAssert.IsFalse(result);
        }

        [Test]
        public void GetErrorCellValue_returns0_onABlankCell()
        {
            ICell cell = new SXSSFCell(null, CellType.Blank);
            ClassicAssert.AreEqual(CellType.Blank, cell.CellType);
            byte result = cell.ErrorCellValue;
            ClassicAssert.AreEqual(0, result);
        }

        [Test]
        [Ignore("Stub")]
        public override void SetCellType_BLANK_removesArrayFormula_ifCellIsPartOfAnArrayFormulaGroupContainingOnlyThisCell()
        {

        }
        [Test]
        [Ignore("Stub")]
        public override void SetCellType_BLANK_throwsISE_ifCellIsPartOfAnArrayFormulaGroupContainingOtherCells()
        {

        }
        [Test]
        [Ignore("Stub")]
        public override void SetCellFormula_throwsISE_ifCellIsPartOfAnArrayFormulaGroupContainingOtherCells()
        {

        }
        [Test]
        [Ignore("Stub")]
        public override void RemoveFormula_turnsCellToBlank_whenFormulaWasASingleCellArrayFormula()
        {

        }
    }
}