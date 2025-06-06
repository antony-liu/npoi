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
    using NPOI.OpenXml4Net.OPC;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.Util;
    using NPOI.XSSF;
    using NPOI.XSSF.Model;
    using NPOI.XSSF.Streaming;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;using NUnit.Framework.Legacy;
    using System;
    using System.IO;
    using TestCases;
    using TestCases.SS.UserModel;
    using TestCases.Util;

    [TestFixture]
    public class TestSXSSFWorkbook : BaseTestXWorkbook
    {

        public TestSXSSFWorkbook()
                : base(SXSSFITestDataProvider.instance)
        {
        }
        [TearDown]
        public void TearDown()
        {
            //((SXSSFITestDataProvider)_testDataProvider).Cleanup();
        }
        /**
         * cloning of sheets is not supported in SXSSF
         */

        [Test]
        public override void CloneSheet()
        {
            try
            {
                base.CloneSheet();
                Assert.Fail("expected exception");
            }
            catch(RuntimeException e)
            {
                ClassicAssert.AreEqual("NotImplemented", e.Message);
            }
        }
        [Test]
        public override void SheetClone()
        {
            try
            {
                base.SheetClone();
                Assert.Fail("expected exception");
            }
            catch(RuntimeException e)
            {
                ClassicAssert.AreEqual("NotImplemented", e.Message);
            }
        }
        /**
         * Skip this test, as SXSSF doesn't update formulas on sheet name
         *  Changes.
         */

        [Test]
        public override void TestSetSheetName()
        {
            Assume.That(false, "SXSSF doesn't update formulas on sheet name Changes, as most cells probably aren't in memory at the time");
        }

        [Test]
        public void ExistingWorkbook()
        {
            XSSFWorkbook xssfWb1 = new XSSFWorkbook();
            xssfWb1.CreateSheet("S1");
            SXSSFWorkbook wb1 = new SXSSFWorkbook(xssfWb1);
            XSSFWorkbook xssfWb2 = SXSSFITestDataProvider.instance.WriteOutAndReadBack(wb1) as XSSFWorkbook;
            ClassicAssert.IsTrue(wb1.Dispose());

            SXSSFWorkbook wb2 = new SXSSFWorkbook(xssfWb2);
            ClassicAssert.AreEqual(1, wb2.NumberOfSheets);
            ISheet sheet = wb2.GetSheetAt(0);
            ClassicAssert.IsNotNull(sheet);
            ClassicAssert.AreEqual("S1", sheet.SheetName);
            ClassicAssert.IsTrue(wb2.Dispose());
            xssfWb2.Close();
            xssfWb1.Close();

            wb2.Close();
            wb1.Close();
        }

        [Test]
        public void UseSharedStringsTable()
        {
            SXSSFWorkbook wb = new SXSSFWorkbook(null, 10, false, true);

            SharedStringsTable sss = POITestCase.GetFieldValue<SharedStringsTable, SXSSFWorkbook>(typeof(SXSSFWorkbook), wb, typeof(SharedStringsTable), "_sharedStringSource");

            ClassicAssert.IsNotNull(sss);

            IRow row = wb.CreateSheet("S1").CreateRow(0);

            row.CreateCell(0).SetCellValue("A");
            row.CreateCell(1).SetCellValue("B");
            row.CreateCell(2).SetCellValue("A");

            XSSFWorkbook xssfWorkbook = SXSSFITestDataProvider.instance.WriteOutAndReadBack(wb) as XSSFWorkbook;
            sss = POITestCase.GetFieldValue<SharedStringsTable, SXSSFWorkbook>(typeof(SXSSFWorkbook), wb, typeof(SharedStringsTable), "_sharedStringSource");
            ClassicAssert.AreEqual(2, sss.UniqueCount);
            ClassicAssert.IsTrue(wb.Dispose());

            ISheet sheet1 = xssfWorkbook.GetSheetAt(0);
            ClassicAssert.AreEqual("S1", sheet1.SheetName);
            ClassicAssert.AreEqual(1, sheet1.PhysicalNumberOfRows);
            row = sheet1.GetRow(0);
            ClassicAssert.IsNotNull(row);
            ICell cell = row.GetCell(0);
            ClassicAssert.IsNotNull(cell);
            ClassicAssert.AreEqual("A", cell.StringCellValue);
            cell = row.GetCell(1);
            ClassicAssert.IsNotNull(cell);
            ClassicAssert.AreEqual("B", cell.StringCellValue);
            cell = row.GetCell(2);
            ClassicAssert.IsNotNull(cell);
            ClassicAssert.AreEqual("A", cell.StringCellValue);

            xssfWorkbook.Close();
            wb.Close();
        }

        [Test]
        public void AddToExistingWorkbook()
        {
            XSSFWorkbook xssfWb1 = new XSSFWorkbook();
            xssfWb1.CreateSheet("S1");
            ISheet sheet = xssfWb1.CreateSheet("S2");
            IRow row = sheet.CreateRow(1);
            ICell cell = row.CreateCell(1);
            cell.SetCellValue("value 2_1_1");
            SXSSFWorkbook wb1 = new SXSSFWorkbook(xssfWb1);
            XSSFWorkbook xssfWb2 = SXSSFITestDataProvider.instance.WriteOutAndReadBack(wb1) as XSSFWorkbook;
            ClassicAssert.IsTrue(wb1.Dispose());
            xssfWb1.Close();

            SXSSFWorkbook wb2 = new SXSSFWorkbook(xssfWb2);
            // Add a row to the existing empty sheet
            ISheet sheet1 = wb2.GetSheetAt(0);
            IRow row1_1 = sheet1.CreateRow(1);
            ICell cell1_1_1 = row1_1.CreateCell(1);
            cell1_1_1.SetCellValue("value 1_1_1");

            // Add a row to the existing non-empty sheet
            ISheet sheet2 = wb2.GetSheetAt(1);
            IRow row2_2 = sheet2.CreateRow(2);
            ICell cell2_2_1 = row2_2.CreateCell(1);
            cell2_2_1.SetCellValue("value 2_2_1");

            // Add a sheet with one row
            ISheet sheet3 = wb2.CreateSheet("S3");
            IRow row3_1 = sheet3.CreateRow(1);
            ICell cell3_1_1 = row3_1.CreateCell(1);
            cell3_1_1.SetCellValue("value 3_1_1");

            XSSFWorkbook xssfWb3 = SXSSFITestDataProvider.instance.WriteOutAndReadBack(wb2) as XSSFWorkbook;
            wb2.Close();

            ClassicAssert.AreEqual(3, xssfWb3.NumberOfSheets);
            // Verify sheet 1
            sheet1 = xssfWb3.GetSheetAt(0);
            ClassicAssert.AreEqual("S1", sheet1.SheetName);
            ClassicAssert.AreEqual(1, sheet1.PhysicalNumberOfRows);
            row1_1 = sheet1.GetRow(1);
            ClassicAssert.IsNotNull(row1_1);
            cell1_1_1 = row1_1.GetCell(1);
            ClassicAssert.IsNotNull(cell1_1_1);
            ClassicAssert.AreEqual("value 1_1_1", cell1_1_1.StringCellValue);
            // Verify sheet 2
            sheet2 = xssfWb3.GetSheetAt(1);
            ClassicAssert.AreEqual("S2", sheet2.SheetName);
            ClassicAssert.AreEqual(2, sheet2.PhysicalNumberOfRows);
            IRow row2_1 = sheet2.GetRow(1);
            ClassicAssert.IsNotNull(row2_1);
            ICell cell2_1_1 = row2_1.GetCell(1);
            ClassicAssert.IsNotNull(cell2_1_1);
            ClassicAssert.AreEqual("value 2_1_1", cell2_1_1.StringCellValue);
            row2_2 = sheet2.GetRow(2);
            ClassicAssert.IsNotNull(row2_2);
            cell2_2_1 = row2_2.GetCell(1);
            ClassicAssert.IsNotNull(cell2_2_1);
            ClassicAssert.AreEqual("value 2_2_1", cell2_2_1.StringCellValue);
            // Verify sheet 3
            sheet3 = xssfWb3.GetSheetAt(2);
            ClassicAssert.AreEqual("S3", sheet3.SheetName);
            ClassicAssert.AreEqual(1, sheet3.PhysicalNumberOfRows);
            row3_1 = sheet3.GetRow(1);
            ClassicAssert.IsNotNull(row3_1);
            cell3_1_1 = row3_1.GetCell(1);
            ClassicAssert.IsNotNull(cell3_1_1);
            ClassicAssert.AreEqual("value 3_1_1", cell3_1_1.StringCellValue);

            xssfWb2.Close();
            xssfWb3.Close();
            wb1.Close();
        }

        [Test]
        public void SheetdataWriter()
        {
            using(SXSSFWorkbook wb = new SXSSFWorkbook())
            {
                SXSSFSheet sh = wb.CreateSheet() as SXSSFSheet;
                SheetDataWriter wr = sh.SheetDataWriter;
                ClassicAssert.IsTrue(wr.GetType() == typeof(SheetDataWriter));
                FileInfo tmp = wr.TempFileInfo;
                ClassicAssert.IsTrue(tmp.Name.StartsWith("poi-sxssf-sheet"));
                ClassicAssert.IsTrue(tmp.Name.EndsWith(".xml"));
            }

            using(var wb2 = new SXSSFWorkbook())
            {
                wb2.CompressTempFiles = (/*setter*/true);
                var sh2 = wb2.CreateSheet() as SXSSFSheet;
                var wr2 = sh2.SheetDataWriter;
                ClassicAssert.IsTrue(wr2.GetType() == typeof(GZIPSheetDataWriter));
                var tmp2 = wr2.TempFileInfo;
                ClassicAssert.IsTrue(tmp2.Name.StartsWith("poi-sxssf-sheet-xml"));
                ClassicAssert.IsTrue(tmp2.Name.EndsWith(".gz"));
            }

            //Test escaping of Unicode control characters
            using(var wb3 = new SXSSFWorkbook())
            {
                wb3.CreateSheet("S1").CreateRow(0).CreateCell(0).SetCellValue("value\u0019");
                using(var xssfWorkbook = SXSSFITestDataProvider.instance.WriteOutAndReadBack(wb3) as XSSFWorkbook)
                {
                    ICell cell = xssfWorkbook.GetSheet("S1").GetRow(0).GetCell(0);
                    ClassicAssert.AreEqual("value?", cell.StringCellValue);
                }
            }
            ;
        }

        [Test]
        public void GzipSheetdataWriter()
        {
            using(var wb = new SXSSFWorkbook())
            {
                wb.CompressTempFiles = true;
                int rowNum = 1000;
                int sheetNum = 5;
                for(int i = 0; i < sheetNum; i++)
                {
                    ISheet sh = wb.CreateSheet("sheet" + i);
                    for(int j = 0; j < rowNum; j++)
                    {
                        IRow row = sh.CreateRow(j);
                        ICell cell1 = row.CreateCell(0);
                        cell1.SetCellValue(new CellReference(cell1).FormatAsString());

                        ICell cell2 = row.CreateCell(1);
                        cell2.SetCellValue(i);

                        ICell cell3 = row.CreateCell(2);
                        cell3.SetCellValue(j);
                    }
                }


                using(XSSFWorkbook xwb = SXSSFITestDataProvider.instance.WriteOutAndReadBack(wb) as XSSFWorkbook)
                {
                    for(int i = 0; i < sheetNum; i++)
                    {
                        ISheet sh = xwb.GetSheetAt(i);
                        ClassicAssert.AreEqual("sheet" + i, sh.SheetName);
                        for(int j = 0; j < rowNum; j++)
                        {
                            IRow row = sh.GetRow(j);
                            ClassicAssert.IsNotNull(row, "row[" + j + "]");
                            ICell cell1 = row.GetCell(0);
                            ClassicAssert.AreEqual(new CellReference(cell1).FormatAsString(), cell1.StringCellValue);

                            ICell cell2 = row.GetCell(1);
                            ClassicAssert.AreEqual(i, (int) cell2.NumericCellValue);

                            ICell cell3 = row.GetCell(2);
                            ClassicAssert.AreEqual(j, (int) cell3.NumericCellValue);
                        }
                    }
                }
            }
        }

        protected static void assertWorkbookDispose(SXSSFWorkbook wb)
        {
            int rowNum = 1000;
            int sheetNum = 5;
            for(int i = 0; i < sheetNum; i++)
            {
                ISheet sh = wb.CreateSheet("sheet" + i);
                for(int j = 0; j < rowNum; j++)
                {
                    IRow row = sh.CreateRow(j);
                    ICell cell1 = row.CreateCell(0);
                    cell1.SetCellValue(new CellReference(cell1).FormatAsString());

                    ICell cell2 = row.CreateCell(1);
                    cell2.SetCellValue(i);

                    ICell cell3 = row.CreateCell(2);
                    cell3.SetCellValue(j);
                }
            }

            foreach(ISheet sheet in wb)
            {
                SXSSFSheet sxSheet = (SXSSFSheet)sheet;
                ClassicAssert.IsTrue(sxSheet.SheetDataWriter.TempFileInfo.Exists);
            }

            ClassicAssert.IsTrue(wb.Dispose());

            foreach(ISheet sheet in wb)
            {
                SXSSFSheet sxSheet = (SXSSFSheet)sheet;
                ClassicAssert.IsFalse(sxSheet.SheetDataWriter.TempFileInfo.Exists);
            }
        }

        [Test]
        public void WorkbookDispose()
        {
            SXSSFWorkbook wb1 = new SXSSFWorkbook();
            // the underlying Writer is SheetDataWriter
            assertWorkbookDispose(wb1);
            wb1.Close();
            SXSSFWorkbook wb2 = new SXSSFWorkbook();
            wb2.CompressTempFiles = (/*setter*/true);
            // the underlying Writer is GZIPSheetDataWriter
            assertWorkbookDispose(wb2);
            wb2.Close();
        }

        [Ignore("currently writing the same sheet multiple times is not supported...")]
        [Test]
        public void Bug53515()
        {
            IWorkbook wb1 = new SXSSFWorkbook(10);
            populateWorkbook(wb1);
            saveTwice(wb1);
            IWorkbook wb2 = new XSSFWorkbook();
            populateWorkbook(wb2);
            saveTwice(wb2);
            wb2.Close();
            wb1.Close();
        }

        [Ignore("Crashes the JVM because of documented JVM behavior with concurrent writing/reading of zip-files, "
                + "see http://www.oracle.com/technetwork/java/javase/documentation/overview-156328.html")]
        [Test]
        public void Bug53515a()
        {
            FileInfo out1 = new FileInfo("Test.xlsx");
            out1.Delete();
            for(int i = 0; i < 2; i++)
            {
                Console.WriteLine("Iteration " + i);
                SXSSFWorkbook wb;
                if(out1.Exists)
                {
                    wb = new SXSSFWorkbook(
                            (XSSFWorkbook) WorkbookFactory.Create(out1.FullName));
                }
                else
                {
                    wb = new SXSSFWorkbook(10);
                }

                try
                {
                    FileStream outSteam = new FileStream(out1.FullName, FileMode.Create, FileAccess.ReadWrite);
                    if(i == 0)
                    {
                        populateWorkbook(wb);
                    }
                    else
                    {
                        //System.Gc();
                        //System.Gc();
                        //System.Gc();
                        GC.Collect();
                        GC.Collect();
                        GC.Collect();
                    }

                    wb.Write(outSteam);
                    // ClassicAssert.IsTrue(wb.Dispose());
                    outSteam.Close();
                }
                finally
                {
                    ClassicAssert.IsTrue(wb.Dispose());
                }
                wb.Close();
            }
            out1.Delete();
        }

        private static void populateWorkbook(IWorkbook wb)
        {
            ISheet sh = wb.CreateSheet();
            for(int rownum = 0; rownum < 100; rownum++)
            {
                IRow row = sh.CreateRow(rownum);
                for(int cellnum = 0; cellnum < 10; cellnum++)
                {
                    ICell cell = row.CreateCell(cellnum);
                    String Address = new CellReference(cell).FormatAsString();
                    cell.SetCellValue(Address);
                }
            }
        }

        private static void saveTwice(IWorkbook wb)
        {
            for(int i = 0; i < 2; i++)
            {
                try
                {
                    NullOutputStream out1 = new NullOutputStream();
                    wb.Write(out1, false);
                    out1.Close();
                }
                catch(Exception e)
                {
                    throw new Exception("ERROR: failed on " + (i + 1)
                            + "th time calling " + wb.GetType().Name
                            + ".Write() with exception " + e.Message, e);
                }
            }
        }

        [Ignore("Just a local test for http://stackoverflow.com/questions/33627329/apache-poi-streaming-api-using-xssf-template")]
        [Test]
        public void TestTemplateFile()
        {
            XSSFWorkbook workBook = XSSFTestDataSamples.OpenSampleWorkbook("sample.xlsx");
            SXSSFWorkbook streamingWorkBook = new SXSSFWorkbook(workBook, 10);
            ISheet sheet = streamingWorkBook.GetSheet("Sheet1");
            for(int rowNum = 10; rowNum < 1000000; rowNum++)
            {
                IRow row = sheet.CreateRow(rowNum);
                for(int cellNum = 0; cellNum < 700; cellNum++)
                {
                    ICell cell = row.CreateCell(cellNum);
                    cell.SetCellValue("somEvalue");
                }

                if(rowNum % 100 == 0)
                {
                    Console.Write(".");
                    if(rowNum % 10000 == 0)
                    {
                        Console.WriteLine(rowNum);
                    }
                }
            }

            FileStream fos = new FileStream("C:\\temp\\streaming.xlsx", FileMode.Create, FileAccess.ReadWrite);
            streamingWorkBook.Write(fos);
            fos.Close();

            streamingWorkBook.Close();
            workBook.Close();
        }

        [Test]
        public void closeDoesNotModifyWorkbook()
        {
            String filename = "SampleSS.xlsx";
            FileInfo file = POIDataSamples.GetSpreadSheetInstance().GetFileInfo(filename);
            SXSSFWorkbook wb = null;
            XSSFWorkbook xwb = null;

            // Some tests commented out because close() modifies the file
            // See bug 58779

            // String
            //wb = new SXSSFWorkbook(new XSSFWorkbook(file.Path));
            //assertCloseDoesNotModifyFile(filename, wb);

            // File
            //wb = new SXSSFWorkbook(new XSSFWorkbook(file));
            //assertCloseDoesNotModifyFile(filename, wb);

            // InputStream
            FileStream fis = file.Open(FileMode.Open, FileAccess.ReadWrite);
            try
            {
                xwb = new XSSFWorkbook(fis);
                wb = new SXSSFWorkbook(xwb);
                assertCloseDoesNotModifyFile(filename, wb);
            }
            finally
            {
                if(xwb != null)
                {
                    xwb.Close();
                }
                if(wb != null)
                {
                    wb.Close();
                }
                fis.Close();
            }

            // OPCPackage
            //wb = new SXSSFWorkbook(new XSSFWorkbook(OPCPackage.open(file)));
            //assertCloseDoesNotModifyFile(filename, wb);
        }
        /**
         * Bug #59743
         * 
         * this is only triggered on other files apart of sheet[1,2,...].xml
         * as those are either copied uncompressed or with the use of GZIPInputStream
         * so we use shared strings
         */
        [Test]
        public void TestZipBombNotTriggeredOnUselessContent()
        {
            SXSSFWorkbook swb = new SXSSFWorkbook(null, 1, true, true);
            SXSSFSheet s = swb.CreateSheet() as SXSSFSheet;
            char[] useless = new char[32767];
            Arrays.Fill(useless, ' ');

            for(int row = 0; row < 1; row++)
            {
                IRow r = s.CreateRow(row);
                for(int col = 0; col < 10; col++)
                {
                    char[] prefix = HexDump.ToHex(row * 1000 + col).ToCharArray();
                    Arrays.Fill(useless, 0, 10, ' ');
                    Array.Copy(prefix, 0, useless, 0, prefix.Length);
                    String ul = new String(useless);
                    r.CreateCell(col, CellType.String).SetCellValue(ul);
                    ul = null;
                }
            }

            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            swb.Write(bos);
            swb.Dispose();
            swb.Close();
        }

        /**
         * To avoid accident changes to the template, you should be able
         *  to create a SXSSFWorkbook from a read-only XSSF one, then
         *  change + save that (only). See bug #60010
         * TODO Fix this to work!
         */

        [Ignore("")]
        [Test]
        public void CreateFromReadOnlyWorkbook()
        {
            FileInfo input = XSSFTestDataSamples.GetSampleFile("sample.xlsx");
            OPCPackage pkg = OPCPackage.Open(input, PackageAccess.READ);
            XSSFWorkbook xssf = new XSSFWorkbook(pkg);
            SXSSFWorkbook wb = new SXSSFWorkbook(xssf, 2);

            String sheetName = "Test SXSSF";
            ISheet s = wb.CreateSheet(sheetName);
            for(int i = 0; i < 10; i++)
            {
                IRow r = s.CreateRow(i);
                r.CreateCell(0).SetCellValue(true);
                r.CreateCell(1).SetCellValue(2.4);
                r.CreateCell(2).SetCellValue("Test Row " + i);
            }
            ClassicAssert.AreEqual(10, s.LastRowNum);

            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            wb.Write(bos);
            wb.Dispose();
            wb.Close();

            xssf = new XSSFWorkbook(new ByteArrayInputStream(bos.ToByteArray()));
            s = xssf.GetSheet(sheetName);
            ClassicAssert.AreEqual(10, s.LastRowNum);
            ClassicAssert.AreEqual(true, s.GetRow(0).GetCell(0).BooleanCellValue);
            ClassicAssert.AreEqual("Test Row 9", s.GetRow(9).GetCell(2).StringCellValue);
        }

        [Test]
        public void Test56557()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("56557.xlsx");

            // Using streaming XSSFWorkbook makes the output file invalid
            wb = new SXSSFWorkbook((XSSFWorkbook)wb);
        
            // Should not throw POIXMLException: java.io.IOException: Unable to parse xml bean when reading back
            IWorkbook wbBack = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            ClassicAssert.IsNotNull(wbBack);
            wbBack.Close();
        
            wb.Close();
        }
    }
}