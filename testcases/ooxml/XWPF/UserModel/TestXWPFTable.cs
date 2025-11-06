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
namespace TestCases.XWPF.UserModel
{
    using NPOI.OpenXmlFormats.Wordprocessing;
    using NPOI.XWPF.UserModel;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using static NPOI.XWPF.UserModel.XWPFTable;


    /**
     * Tests for XWPF Tables
     */
    [TestFixture]
    public class TestXWPFTable
    {
        [Test]
        public void TestConstructor()
        {
            XWPFDocument doc = new XWPFDocument();
            CT_Tbl ctTable = new CT_Tbl();
            XWPFTable xtab = new XWPFTable(ctTable, doc);
            ClassicAssert.IsNotNull(xtab);
            ClassicAssert.AreEqual(1, ctTable.SizeOfTrArray());
            ClassicAssert.AreEqual(1, ctTable.GetTrArray(0).SizeOfTcArray());
            ClassicAssert.IsNotNull(ctTable.GetTrArray(0).GetTcArray(0).GetPArray(0));

            ctTable = new CT_Tbl();
            xtab = new XWPFTable(ctTable, doc, 3, 2);
            ClassicAssert.IsNotNull(xtab);
            ClassicAssert.AreEqual(3, ctTable.SizeOfTrArray());
            ClassicAssert.AreEqual(2, ctTable.GetTrArray(0).SizeOfTcArray());
            ClassicAssert.IsNotNull(ctTable.GetTrArray(0).GetTcArray(0).GetPArray(0));

            try
            {
                doc.Close();
            }
            catch(IOException e)
            {
                Assert.Fail("Unable to close doc");
            }
        }

        [Test]
        public void TestTblGrid()
        {
            XWPFDocument doc = new XWPFDocument();
            CT_Tbl ctTable = new CT_Tbl();
            CT_TblGrid cttblgrid = ctTable.AddNewTblGrid();
            cttblgrid.AddNewGridCol().w = 123;
            cttblgrid.AddNewGridCol().w = 321;

            XWPFTable xtab = new XWPFTable(ctTable, doc);
            ClassicAssert.AreEqual(123, xtab.GetCTTbl().tblGrid.gridCol[0].w);
            ClassicAssert.AreEqual(321, xtab.GetCTTbl().tblGrid.gridCol[1].w);

            try
            {
                doc.Close();
            }
            catch(IOException e)
            {
                Assert.Fail("Unable to close doc");
            }
        }

        [Test]
        public void TestGetText()
        {
            XWPFDocument doc = new XWPFDocument();
            CT_Tbl table = new CT_Tbl();
            CT_Row row = table.AddNewTr();
            CT_Tc cell = row.AddNewTc();
            CT_P paragraph = cell.AddNewP();
            CT_R run = paragraph.AddNewR();
            CT_Text text = run.AddNewT();
            text.Value = ("finally I can Write!");

            XWPFTable xtab = new XWPFTable(table, doc);
            ClassicAssert.AreEqual("finally I can Write!\n", xtab.Text);

            try
            {
                doc.Close();
            }
            catch(IOException e)
            {
                Assert.Fail("Unable to close doc");
            }
        }


        [Test]
        public void TestCreateRow()
        {
            XWPFDocument doc = new XWPFDocument();

            CT_Tbl table = new CT_Tbl();
            CT_Row r1 = table.AddNewTr();
            r1.AddNewTc().AddNewP();
            r1.AddNewTc().AddNewP();
            CT_Row r2 = table.AddNewTr();
            r2.AddNewTc().AddNewP();
            r2.AddNewTc().AddNewP();
            CT_Row r3 = table.AddNewTr();
            r3.AddNewTc().AddNewP();
            r3.AddNewTc().AddNewP();

            XWPFTable xtab = new XWPFTable(table, doc);
            ClassicAssert.AreEqual(3, xtab.NumberOfRows);
            ClassicAssert.IsNotNull(xtab.GetRow(2));

            //add a new row
            xtab.CreateRow();
            ClassicAssert.AreEqual(4, xtab.NumberOfRows);

            //check number of cols
            ClassicAssert.AreEqual(2, xtab.NumberOfColumns);

            //check creation of first row
            xtab = new XWPFTable(new CT_Tbl(), doc);
            ClassicAssert.AreEqual(1, xtab.GetCTTbl().GetTrArray(0).SizeOfTcArray());

            try
            {
                doc.Close();
            }
            catch(IOException e)
            {
                Assert.Fail("Unable to close doc");
            }
        }
        [Test]
        public void TestInsertRow()
        {
            XWPFDocument doc = new XWPFDocument();

            CT_Tbl table = new CT_Tbl();
            CT_Row r1 = table.AddNewTr();
            r1.AddNewTc().AddNewP();
            r1.AddNewTc().AddNewP();
            CT_Row r2 = table.AddNewTr();
            r2.AddNewTc().AddNewP();
            r2.AddNewTc().AddNewP();
            CT_Row r3 = table.AddNewTr();
            r3.AddNewTc().AddNewP();
            r3.AddNewTc().AddNewP();

            XWPFTable xtab = new XWPFTable(table, doc);
            ClassicAssert.AreEqual(3, xtab.NumberOfRows);
            ClassicAssert.IsNotNull(xtab.GetRow(2));

            //add a new row
            xtab.CreateRow();
            ClassicAssert.AreEqual(4, xtab.NumberOfRows);

            xtab.InsertNewTableRow(0);
            ClassicAssert.AreEqual(5, xtab.NumberOfRows);
            xtab.InsertNewTableRow(xtab.NumberOfRows);
            ClassicAssert.AreEqual(6, xtab.NumberOfRows);


        }

        [Test]
        public void TestSetGetWidth()
        {
            XWPFDocument doc = new XWPFDocument();

            CT_Tbl table = new CT_Tbl();
            table.AddNewTblPr().AddNewTblW().w = "1000";

            XWPFTable xtab = new XWPFTable(table, doc);

            ClassicAssert.AreEqual(1000, xtab.Width);

            xtab.Width = 100;
            ClassicAssert.AreEqual(100, int.Parse(table.tblPr.tblW.w));

            try
            {
                doc.Close();
            }
            catch(IOException e)
            {
                Assert.Fail("Unable to close doc");
            }
        }
        [Test]
        public void TestSetGetHeight()
        {
            XWPFDocument doc = new XWPFDocument();

            CT_Tbl table = new CT_Tbl();

            XWPFTable xtab = new XWPFTable(table, doc);
            XWPFTableRow row = xtab.CreateRow();
            row.Height = (20);
            ClassicAssert.AreEqual(20, row.Height);

            try
            {
                doc.Close();
            }
            catch(IOException e)
            {
                Assert.Fail("Unable to close doc");
            }
        }
        [Test]
        public void TestSetGetMargins()
        {
            // instantiate the following class so it'll Get picked up by
            // the XmlBean process and Added to the jar file. it's required
            // for the following XWPFTable methods.
            CT_TblCellMar ctm = new CT_TblCellMar();
            ClassicAssert.IsNotNull(ctm);
            // create a table
            XWPFDocument doc = new XWPFDocument();
            CT_Tbl ctTable = new CT_Tbl();
            XWPFTable table = new XWPFTable(ctTable, doc);
            // Set margins
            table.SetCellMargins(50, 50, 250, 450);
            // Get margin components
            int t = table.CellMarginTop;
            ClassicAssert.AreEqual(50, t);
            int l = table.CellMarginLeft;
            ClassicAssert.AreEqual(50, l);
            int b = table.CellMarginBottom;
            ClassicAssert.AreEqual(250, b);
            int r = table.CellMarginRight;
            ClassicAssert.AreEqual(450, r);

            try
            {
                doc.Close();
            }
            catch(IOException e)
            {
                Assert.Fail("Unable to close doc");
            }
        }

        [Test]
        public void TestSetGetHBorders()
        {
            // instantiate the following classes so they'll Get picked up by
            // the XmlBean process and added to the jar file. they are required
            // for the following XWPFTable methods.
            CT_TblBorders cttb = new CT_TblBorders();
            ClassicAssert.IsNotNull(cttb);
            ST_Border stb = new ST_Border();
            ClassicAssert.IsNotNull(stb);
            // create a table
            XWPFDocument doc = new XWPFDocument();
            CT_Tbl ctTable = new CT_Tbl();
            XWPFTable table = new XWPFTable(ctTable, doc);
            // check initial state
            XWPFBorderType? bt = table.InsideHBorderType;
            ClassicAssert.AreEqual(XWPFBorderType.SINGLE, bt);
            int sz = table.InsideHBorderSize;
            ClassicAssert.AreEqual(-1, sz);
            int sp = table.InsideHBorderSpace;
            ClassicAssert.AreEqual(-1, sp);
            string clr = table.InsideHBorderColor;
            ClassicAssert.IsNull(clr);
            // Set inside horizontal border
            table.SetInsideHBorder(XWPFBorderType.SINGLE, 4, 0, "FF0000");
            // Get inside horizontal border components
            bt = table.InsideHBorderType;
            ClassicAssert.AreEqual(XWPFBorderType.SINGLE, bt);
            sz = table.InsideHBorderSize;
            ClassicAssert.AreEqual(4, sz);
            sp = table.InsideHBorderSpace;
            ClassicAssert.AreEqual(0, sp);
            clr = table.InsideHBorderColor;
            ClassicAssert.AreEqual("FF0000", clr);
            // remove the border and verify state
            table.RemoveInsideHBorder();
            bt = table.InsideHBorderType;
            ClassicAssert.IsNull(bt);
            sz = table.InsideHBorderSize;
            ClassicAssert.AreEqual(-1, sz);
            sp = table.InsideHBorderSpace;
            ClassicAssert.AreEqual(-1, sp);
            clr = table.InsideHBorderColor;
            ClassicAssert.IsNull(clr);
            // check other borders
            bt = table.InsideVBorderType;
            ClassicAssert.AreEqual(XWPFBorderType.SINGLE, bt);
            bt = table.TopBorderType;
            ClassicAssert.AreEqual(XWPFBorderType.SINGLE, bt);
            bt = table.BottomBorderType;
            ClassicAssert.AreEqual(XWPFBorderType.SINGLE, bt);
            bt = table.LeftBorderType;
            ClassicAssert.AreEqual(XWPFBorderType.SINGLE, bt);
            bt = table.RightBorderType;
            ClassicAssert.AreEqual(XWPFBorderType.SINGLE, bt);
            // remove the rest all at once and test
            table.RemoveBorders();
            bt = table.InsideVBorderType;
            ClassicAssert.IsNull(bt);
            bt = table.TopBorderType;
            ClassicAssert.IsNull(bt);
            bt = table.BottomBorderType;
            ClassicAssert.IsNull(bt);
            bt = table.LeftBorderType;
            ClassicAssert.IsNull(bt);
            bt = table.RightBorderType;
            ClassicAssert.IsNull(bt);
            try
            {
                doc.Close();
            }
            catch(IOException e)
            {
                Assert.Fail("Unable to close doc");
            }
        }

        [Test]
        public void TestSetGetVBorders()
        {
            // create a table
            XWPFDocument doc = new XWPFDocument();
            CT_Tbl ctTable = new CT_Tbl();
            XWPFTable table = new XWPFTable(ctTable, doc);
            // check initial state
            XWPFBorderType? bt = table.InsideVBorderType;
            ClassicAssert.AreEqual(XWPFBorderType.SINGLE, bt);
            int sz = table.InsideVBorderSize;
            ClassicAssert.AreEqual(-1, sz);
            int sp = table.InsideVBorderSpace;
            ClassicAssert.AreEqual(-1, sp);
            string clr = table.InsideVBorderColor;
            ClassicAssert.IsNull(clr);
            // Set inside vertical border
            table.SetInsideVBorder(XWPFBorderType.DOUBLE, 4, 0, "00FF00");
            // Get inside vertical border components
            bt = table.InsideVBorderType;
            ClassicAssert.AreEqual(XWPFBorderType.DOUBLE, bt);
            sz = table.InsideVBorderSize;
            ClassicAssert.AreEqual(4, sz);
            sp = table.InsideVBorderSpace;
            ClassicAssert.AreEqual(0, sp);
            clr = table.InsideVBorderColor;
            ClassicAssert.AreEqual("00FF00", clr);
            // remove the border and verify state
            table.RemoveInsideVBorder();
            bt = table.InsideVBorderType;
            ClassicAssert.IsNull(bt);
            sz = table.InsideVBorderSize;
            ClassicAssert.AreEqual(-1, sz);
            sp = table.InsideVBorderSpace;
            ClassicAssert.AreEqual(-1, sp);
            clr = table.InsideVBorderColor;
            ClassicAssert.IsNull(clr);
            // check the rest
            bt = table.InsideHBorderType;
            ClassicAssert.AreEqual(XWPFBorderType.SINGLE, bt);
            bt = table.TopBorderType;
            ClassicAssert.AreEqual(XWPFBorderType.SINGLE, bt);
            bt = table.BottomBorderType;
            ClassicAssert.AreEqual(XWPFBorderType.SINGLE, bt);
            bt = table.LeftBorderType;
            ClassicAssert.AreEqual(XWPFBorderType.SINGLE, bt);
            bt = table.RightBorderType;
            ClassicAssert.AreEqual(XWPFBorderType.SINGLE, bt);
            // remove the rest one at a time and test
            table.RemoveInsideHBorder();
            table.RemoveTopBorder();
            table.RemoveBottomBorder();
            table.RemoveLeftBorder();
            table.RemoveRightBorder();
            bt = table.InsideHBorderType;
            ClassicAssert.IsNull(bt);
            bt = table.TopBorderType;
            ClassicAssert.IsNull(bt);
            bt = table.BottomBorderType;
            ClassicAssert.IsNull(bt);
            bt = table.LeftBorderType;
            ClassicAssert.IsNull(bt);
            bt = table.RightBorderType;
            ClassicAssert.IsNull(bt);
            try
            {
                doc.Close();
            }
            catch(IOException e)
            {
                Assert.Fail("Unable to close doc");
            }
        }

        [Test]
        public void TestSetGetTopBorders()
        {
            // create a table
            XWPFDocument doc = new XWPFDocument();
            CT_Tbl ctTable = new CT_Tbl();
            XWPFTable table = new XWPFTable(ctTable, doc);
            // check initial state
            XWPFBorderType? bt = table.TopBorderType;
            ClassicAssert.AreEqual(XWPFBorderType.SINGLE, bt);
            int sz = table.TopBorderSize;
            ClassicAssert.AreEqual(-1, sz);
            int sp = table.TopBorderSpace;
            ClassicAssert.AreEqual(-1, sp);
            string clr = table.TopBorderColor;
            ClassicAssert.IsNull(clr);
            // Set top border
            table.SetTopBorder(XWPFBorderType.THICK, 4, 0, "00FF00");
            // Get inside vertical border components
            bt = table.TopBorderType;
            ClassicAssert.AreEqual(XWPFBorderType.THICK, bt);
            sz = table.TopBorderSize;
            ClassicAssert.AreEqual(4, sz);
            sp = table.TopBorderSpace;
            ClassicAssert.AreEqual(0, sp);
            clr = table.TopBorderColor;
            ClassicAssert.AreEqual("00FF00", clr);
            // remove the border and verify state
            table.RemoveTopBorder();
            bt = table.TopBorderType;
            ClassicAssert.IsNull(bt);
            sz = table.TopBorderSize;
            ClassicAssert.AreEqual(-1, sz);
            sp = table.TopBorderSpace;
            ClassicAssert.AreEqual(-1, sp);
            clr = table.TopBorderColor;
            ClassicAssert.IsNull(clr);
            try
            {
                doc.Close();
            }
            catch(IOException e)
            {
                Assert.Fail("Unable to close doc");
            }
        }

        [Test]
        public void TestSetGetBottomBorders()
        {
            // create a table
            XWPFDocument doc = new XWPFDocument();
            CT_Tbl ctTable = new CT_Tbl();
            XWPFTable table = new XWPFTable(ctTable, doc);
            // check initial state
            XWPFBorderType? bt = table.BottomBorderType;
            ClassicAssert.AreEqual(XWPFBorderType.SINGLE, bt);
            int sz = table.BottomBorderSize;
            ClassicAssert.AreEqual(-1, sz);
            int sp = table.BottomBorderSpace;
            ClassicAssert.AreEqual(-1, sp);
            string clr = table.BottomBorderColor;
            ClassicAssert.IsNull(clr);
            // Set inside vertical border
            table.SetBottomBorder(XWPFBorderType.DOTTED, 4, 0, "00FF00");
            // Get inside vertical border components
            bt = table.BottomBorderType;
            ClassicAssert.AreEqual(XWPFBorderType.DOTTED, bt);
            sz = table.BottomBorderSize;
            ClassicAssert.AreEqual(4, sz);
            sp = table.BottomBorderSpace;
            ClassicAssert.AreEqual(0, sp);
            clr = table.BottomBorderColor;
            ClassicAssert.AreEqual("00FF00", clr);
            // remove the border and verify state
            table.RemoveBottomBorder();
            bt = table.BottomBorderType;
            ClassicAssert.IsNull(bt);
            sz = table.BottomBorderSize;
            ClassicAssert.AreEqual(-1, sz);
            sp = table.BottomBorderSpace;
            ClassicAssert.AreEqual(-1, sp);
            clr = table.BottomBorderColor;
            ClassicAssert.IsNull(clr);
            try
            {
                doc.Close();
            }
            catch(IOException e)
            {
                Assert.Fail("Unable to close doc");
            }
        }

        [Test]
        public void TestSetGetLeftBorders()
        {
            // create a table
            XWPFDocument doc = new XWPFDocument();
            CT_Tbl ctTable = new CT_Tbl();
            XWPFTable table = new XWPFTable(ctTable, doc);
            // check initial state
            XWPFBorderType? bt = table.LeftBorderType;
            ClassicAssert.AreEqual(XWPFBorderType.SINGLE, bt);
            int sz = table.LeftBorderSize;
            ClassicAssert.AreEqual(-1, sz);
            int sp = table.LeftBorderSpace;
            ClassicAssert.AreEqual(-1, sp);
            string clr = table.LeftBorderColor;
            ClassicAssert.IsNull(clr);
            // Set inside vertical border
            table.SetLeftBorder(XWPFBorderType.DASHED, 4, 0, "00FF00");
            // Get inside vertical border components
            bt = table.LeftBorderType;
            ClassicAssert.AreEqual(XWPFBorderType.DASHED, bt);
            sz = table.LeftBorderSize;
            ClassicAssert.AreEqual(4, sz);
            sp = table.LeftBorderSpace;
            ClassicAssert.AreEqual(0, sp);
            clr = table.LeftBorderColor;
            ClassicAssert.AreEqual("00FF00", clr);
            // remove the border and verify state
            table.RemoveLeftBorder();
            bt = table.LeftBorderType;
            ClassicAssert.IsNull(bt);
            sz = table.LeftBorderSize;
            ClassicAssert.AreEqual(-1, sz);
            sp = table.LeftBorderSpace;
            ClassicAssert.AreEqual(-1, sp);
            clr = table.LeftBorderColor;
            ClassicAssert.IsNull(clr);
            try
            {
                doc.Close();
            }
            catch(IOException e)
            {
                Assert.Fail("Unable to close doc");
            }
        }

        [Test]
        public void TestSetGetRightBorders()
        {
            // create a table
            XWPFDocument doc = new XWPFDocument();
            CT_Tbl ctTable = new CT_Tbl();
            XWPFTable table = new XWPFTable(ctTable, doc);
            // check initial state
            XWPFBorderType? bt = table.RightBorderType;
            ClassicAssert.AreEqual(XWPFBorderType.SINGLE, bt);
            int sz = table.RightBorderSize;
            ClassicAssert.AreEqual(-1, sz);
            int sp = table.RightBorderSpace;
            ClassicAssert.AreEqual(-1, sp);
            string clr = table.RightBorderColor;
            ClassicAssert.IsNull(clr);
            // Set inside vertical border
            table.SetRightBorder(XWPFBorderType.DOT_DASH, 4, 0, "00FF00");
            // Get inside vertical border components
            bt = table.RightBorderType;
            ClassicAssert.AreEqual(XWPFBorderType.DOT_DASH, bt);
            sz = table.RightBorderSize;
            ClassicAssert.AreEqual(4, sz);
            sp = table.RightBorderSpace;
            ClassicAssert.AreEqual(0, sp);
            clr = table.RightBorderColor;
            ClassicAssert.AreEqual("00FF00", clr);
            // remove the border and verify state
            table.RemoveRightBorder();
            bt = table.RightBorderType;
            ClassicAssert.IsNull(bt);
            sz = table.RightBorderSize;
            ClassicAssert.AreEqual(-1, sz);
            sp = table.RightBorderSpace;
            ClassicAssert.AreEqual(-1, sp);
            clr = table.RightBorderColor;
            ClassicAssert.IsNull(clr);
            try
            {
                doc.Close();
            }
            catch(IOException e)
            {
                Assert.Fail("Unable to close doc");
            }
        }

        [Test]
        public void TestSetGetRowBandSize()
        {
            XWPFDocument doc = new XWPFDocument();
            CT_Tbl ctTable = new CT_Tbl();
            XWPFTable table = new XWPFTable(ctTable, doc);
            table.RowBandSize = 12;
            int sz = table.RowBandSize;
            ClassicAssert.AreEqual(12, sz);
            try
            {
                doc.Close();
            }
            catch(IOException e)
            {
                Assert.Fail("Unable to close doc");
            }
        }

        [Test]
        public void TestSetGetColBandSize()
        {
            XWPFDocument doc = new XWPFDocument();
            CT_Tbl ctTable = new CT_Tbl();
            XWPFTable table = new XWPFTable(ctTable, doc);
            table.ColBandSize = 16;
            int sz = table.ColBandSize;
            ClassicAssert.AreEqual(16, sz);
            try
            {
                doc.Close();
            }
            catch(IOException e)
            {
                Assert.Fail("Unable to close doc");
            }
        }
        [Test]
        public void TestCreateTable()
        {
            // open an empty document
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("sample.docx");

            // create a table with 5 rows and 7 coloumns
            int noRows = 5;
            int noCols = 7;
            XWPFTable table = doc.CreateTable(noRows, noCols);

            // assert the table is empty
            List<XWPFTableRow> rows = table.Rows;
            ClassicAssert.AreEqual(noRows, rows.Count, "Table has less rows than requested.");
            foreach(XWPFTableRow xwpfRow in rows)
            {
                ClassicAssert.IsNotNull(xwpfRow);
                for(int i = 0; i < 7; i++)
                {
                    XWPFTableCell xwpfCell = xwpfRow.GetCell(i);
                    ClassicAssert.IsNotNull(xwpfCell);
                    ClassicAssert.AreEqual(1, xwpfCell.Paragraphs.Count, "Empty cells should not have one paragraph.");
                    xwpfCell = xwpfRow.GetCell(i);
                    ClassicAssert.AreEqual(1, xwpfCell.Paragraphs.Count, "Calling 'getCell' must not modify cells content.");
                }
            }
            doc.Package.Revert();

            try
            {
                doc.Close();
            }
            catch(IOException e)
            {
                Assert.Fail("Unable to close doc");
            }
        }

        [Test]
        public void TestSetTableCaption()
        {
            // create a table
            XWPFDocument doc = new XWPFDocument();
            CT_Tbl ctTable = new CT_Tbl();
            XWPFTable table = new XWPFTable(ctTable, doc);

            //Set Caption
            table.TableCaption = "%TABLECAPTION%";

            //Check
            CT_String tblCaption = table.GetCTTbl().tblPr.tblCaption;
            ClassicAssert.IsNotNull(tblCaption);
            ClassicAssert.AreEqual("%TABLECAPTION%", tblCaption.val);
        }

        [Test]
        public void TestSetTableDescription()
        {
            // create a table
            XWPFDocument doc = new XWPFDocument();
            CT_Tbl ctTable = new CT_Tbl();
            XWPFTable table = new XWPFTable(ctTable, doc);

            //Set Description
            table.TableDescription = "%TABLEDESCRIPTION%";

            //Check
            CT_String tblDesc = table.GetCTTbl().tblPr.tblDescription;
            ClassicAssert.IsNotNull(tblDesc);
            ClassicAssert.AreEqual("%TABLEDESCRIPTION%", tblDesc.val);
        }

        [Test]
        public void TestReadTableCaptionAndDescription()
        {
            // open an document with table containing Table caption and Table Description
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("table_properties.docx");

            //Get Table
            var table = doc.Tables[0];

            //Assert Table Caption & Description
            ClassicAssert.AreEqual("Table Title", table.TableCaption);
            ClassicAssert.AreEqual("Table Description", table.TableDescription);
        }

        [Test]
        public void TestSetGetTableAlignment()
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFTable tbl = doc.CreateTable(1, 1);
            tbl.TableAlignment = (TableRowAlign.LEFT);
            ClassicAssert.AreEqual(TableRowAlign.LEFT, tbl.TableAlignment);
            tbl.TableAlignment = (TableRowAlign.CENTER);
            ClassicAssert.AreEqual(TableRowAlign.CENTER, tbl.TableAlignment);
            tbl.TableAlignment = (TableRowAlign.RIGHT);
            ClassicAssert.AreEqual(TableRowAlign.RIGHT, tbl.TableAlignment);
            tbl.RemoveTableAlignment();
            ClassicAssert.IsNull(tbl.TableAlignment);
            try
            {
                doc.Close();
            }
            catch(IOException e)
            {
                ClassicAssert.Fail("Unable to close doc");
            }
        }
    }
}