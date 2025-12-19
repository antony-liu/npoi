using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.Util;
using NPOI.XSSF;
using NPOI.XSSF.Streaming;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using NUnit.Framework.Legacy;
using System;
using System.IO;
using TestCases.HPSF.Basic;
using TestCases.SS;
using TestCases.SS.UserModel;

namespace TestCases.XSSF.Streaming
{
    [TestFixture]
    public class TestSXSSFBugs : BaseTestBugzillaIssues
    {
        public TestSXSSFBugs()
            : base(SXSSFITestDataProvider.instance)
        {

        }
        // override some tests which do not work for SXSSF
        [Ignore("cloneSheet() not implemented")] public override void Bug18800() { /* cloneSheet() not implemented */ }
        [Ignore("cloneSheet() not implemented")] public override void Bug22720() { /* cloneSheet() not implemented */ }
        [Ignore("Evaluation is not fully supported")] public override void Bug47815() { /* Evaluation is not supported */ }
        [Ignore("Evaluation is not fully supported")] public override void Bug46729_testMaxFunctionArguments() { /* Evaluation is not supported */ }
        [Ignore("Reading data is not supported")] public override void Bug57798() { /* Reading data is not supported */ }

        [Test]
        public void Bug49253()
        {
            IWorkbook wb1 = new SXSSFWorkbook();
            IWorkbook wb2 = new SXSSFWorkbook();
            CellRangeAddress cra = CellRangeAddress.ValueOf("C2:D3");

            // No print settings before repeating
            ISheet s1 = wb1.CreateSheet();
            s1.RepeatingColumns = (cra);
            s1.RepeatingRows = (cra);

            IPrintSetup ps1 = s1.PrintSetup;
            ClassicAssert.IsFalse(ps1.ValidSettings);
            ClassicAssert.IsFalse(ps1.Landscape);


            // Had valid print settings before repeating
            ISheet s2 = wb2.CreateSheet();
            IPrintSetup ps2 = s2.PrintSetup;

            ps2.Landscape = (false);
            ClassicAssert.IsTrue(ps2.ValidSettings);
            ClassicAssert.IsFalse(ps2.Landscape);
            s2.RepeatingColumns = (cra);
            s2.RepeatingRows = (cra);

            ps2 = s2.PrintSetup;
            ClassicAssert.IsTrue(ps2.ValidSettings);
            ClassicAssert.IsFalse(ps2.Landscape);

            wb1.Close();
            wb2.Close();
        }

        // bug 60197: setSheetOrder should update sheet-scoped named ranges to maintain references to the sheets before the re-order
        [Test]
        override
        public void bug60197_NamedRangesReferToCorrectSheetWhenSheetOrderIsChanged()
        {
            try
            {
                base.bug60197_NamedRangesReferToCorrectSheetWhenSheetOrderIsChanged();
            }
            catch(RuntimeException e)
            {
                var cause = e.InnerException;
                if(cause is IOException && cause.Message == "Stream closed")
                {
                    // expected on the second time that _testDataProvider.writeOutAndReadBack(SXSSFWorkbook) is called
                    // if the test makes it this far, then we know that XSSFName sheet indices are updated when sheet
                    // order is changed, which is the purpose of this test. Therefore, consider this a passing test.
                }
                else
                {
                    throw;
                }
            }
        }

        [Test]
        public void Bug61648()
        {
            // works as expected
            WriteWorkbook(new XSSFWorkbook(), XSSFITestDataProvider.instance);

            // does not work
            SXSSFWorkbook wb = new SXSSFWorkbook();
            try
            {
                WriteWorkbook(wb, SXSSFITestDataProvider.instance);
                ClassicAssert.Fail("Should catch exception here");
            }
            catch(RuntimeException e)
            {
                // this is not implemented yet
                wb.Close();
            }
        }

        void WriteWorkbook(IWorkbook wb, ITestDataProvider testDataProvider)
        {
            ISheet sheet = wb.CreateSheet("array formula test");

            int rowIndex = 0;
            int colIndex = 0;
            IRow row = sheet.CreateRow(rowIndex++);

            ICell cell = row.CreateCell(colIndex++);
            cell.SetCellType(CellType.String);
            cell.SetCellValue("multiple");
            cell = row.CreateCell(colIndex);
            cell.SetCellType(CellType.String);
            cell.SetCellValue("unique");

            WriteRow(sheet, rowIndex++, 80d, "INDEX(A2:A7, MATCH(FALSE, ISBLANK(A2:A7), 0))");
            WriteRow(sheet, rowIndex++, 30d, "IFERROR(INDEX(A2:A7, MATCH(1, (COUNTIF(B2:B2, A2:A7) = 0) * (NOT(ISBLANK(A2:A7))), 0)), \"\")");
            WriteRow(sheet, rowIndex++, 30d, "IFERROR(INDEX(A2:A7, MATCH(1, (COUNTIF(B2:B3, A2:A7) = 0) * (NOT(ISBLANK(A2:A7))), 0)), \"\")");
            WriteRow(sheet, rowIndex++, 2d, "IFERROR(INDEX(A2:A7, MATCH(1, (COUNTIF(B2:B4, A2:A7) = 0) * (NOT(ISBLANK(A2:A7))), 0)), \"\")");
            WriteRow(sheet, rowIndex++, 30d, "IFERROR(INDEX(A2:A7, MATCH(1, (COUNTIF(B2:B5, A2:A7) = 0) * (NOT(ISBLANK(A2:A7))), 0)), \"\")");
            WriteRow(sheet, rowIndex, 2d, "IFERROR(INDEX(A2:A7, MATCH(1, (COUNTIF(B2:B6, A2:A7) = 0) * (NOT(ISBLANK(A2:A7))), 0)), \"\")");

            /*FileOutputStream fileOut = new FileOutputStream(filename);
            wb.write(fileOut);
            fileOut.Close();*/

            IWorkbook wbBack = testDataProvider.WriteOutAndReadBack(wb);
            ClassicAssert.IsNotNull(wbBack);
            wbBack.Close();

            wb.Close();
        }

        void WriteRow(ISheet sheet, int rowIndex, double col0Value, string col1Value)
        {
            int colIndex = 0;
            IRow row = sheet.CreateRow(rowIndex);

            // numeric value cell
            ICell cell = row.CreateCell(colIndex++);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(col0Value);

            // formula value cell
            CellRangeAddress range = new CellRangeAddress(rowIndex, rowIndex, colIndex, colIndex);
            sheet.SetArrayFormula(col1Value, range);
        }

        [Test]
        [Ignore("takes too long for the normal test run")]
        public void Test62872()
        {
            int COLUMN_COUNT = 300;
            int ROW_COUNT = 600000;
            int TEN_MINUTES = 1000*60*10;

            SXSSFWorkbook workbook = new SXSSFWorkbook(100);
            //workbook.SetCompressTempFiles(true);
            SXSSFSheet sheet = workbook.CreateSheet("RawData") as SXSSFSheet;

            SXSSFRow row = sheet.CreateRow(0) as SXSSFRow;
            SXSSFCell cell;

            for(int i = 1; i <= COLUMN_COUNT; i++)
            {
                cell = row.CreateCell(i - 1) as SXSSFCell;
                cell.SetCellValue("Column " + i);
            }
            var tick = DateTime.Now.Ticks;
            for(int i = 1; i < ROW_COUNT; i++)
            {
                row = sheet.CreateRow(i) as SXSSFRow;
                for(int j = 1; j <= COLUMN_COUNT; j++)
                {
                    cell = row.CreateCell(j - 1) as SXSSFCell;
                    
                    //make some noise
                    cell.SetCellValue(new DateTime(tick + i*TEN_MINUTES+(j*TEN_MINUTES)/COLUMN_COUNT));
                }
                i++;
                // if (i % 1000 == 0)
                // logger.info("Created Row " + i);
            }
            FileStream out1 = new FileStream(TempFile.CreateTempFile("test62872", ".xlsx").FullName,
                FileMode.OpenOrCreate, FileAccess.ReadWrite);
            try
            {
                workbook.Write(out1);
                workbook.Dispose();
                workbook.Close();
                //out1.Flush();
            }
            finally { out1.Close(); }
            // logger.info("File written!");
        }
    }
}
