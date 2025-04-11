/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace TestCases.HSSF.Extractor
{
    using System;
    using System.IO;

    using NPOI.HSSF.UserModel;
    using NPOI.POIFS.FileSystem;
    using NPOI.HSSF.Extractor;
    using TestCases.HSSF;


    using NUnit.Framework;using NUnit.Framework.Legacy;
using NPOI.HSSF.Record.Crypto;

    [TestFixture]
    public class TestExcelExtractor
    {

        private static ExcelExtractor CreateExtractor(String sampleFileName)
        {

            Stream is1 = HSSFTestDataSamples.OpenSampleFileStream(sampleFileName);

            try
            {
                return new ExcelExtractor(new POIFSFileSystem(is1));
            }
            catch (IOException)
            {
                throw;
            }
        }

        [Test]
        public void TestSimple()
        {

            ExcelExtractor extractor = CreateExtractor("Simple.xls");

            ClassicAssert.AreEqual("Sheet1\nreplaceMe\nSheet2\nSheet3\n", extractor.Text);

            // Now turn off sheet names
            extractor.IncludeSheetNames=false;
            ClassicAssert.AreEqual("replaceMe\n", extractor.Text);

            extractor.Close();
        }
        [Test]
        public void TestNumericFormula()
        {

            ExcelExtractor extractor = CreateExtractor("sumifformula.xls");

            string expected = "Sheet1\n" +
                    "1000\t1\t5\n" +
                    "2000\t2\n" +
                    "3000\t3\n" +
                    "4000\t4\n" +
                    "5000\t5\n" +
                    "Sheet2\nSheet3\n";
            ClassicAssert.AreEqual(
                    expected,
                    extractor.Text
            );

            extractor.FormulasNotResults=(true);

            ClassicAssert.AreEqual(
                "Sheet1\n" +
                "1000\t1\tSUMIF(A1:A5,\">4000\",B1:B5)\n" +
                "2000\t2\n" +
                "3000\t3\n" +
                "4000\t4\n" +
                "5000\t5\n" +
                "Sheet2\nSheet3\n", 
                    extractor.Text
            );

            extractor.Close();
        }
        [Test]
        public void TestwithContinueRecords()
        {

            ExcelExtractor extractor = CreateExtractor("StringContinueRecords.xls");

            // Has masses of text
            // Until we fixed bug #41064, this would've
            //   failed by now
            ClassicAssert.IsTrue(extractor.Text.Length > 40960);

            extractor.Close();
        }
        [Test]
        public void TestStringConcat()
        {

            ExcelExtractor extractor = CreateExtractor("SimpleWithFormula.xls");

            // Comes out as NaN if treated as a number
            // And as XYZ if treated as a string
            ClassicAssert.AreEqual("Sheet1\nreplaceme\nreplaceme\nreplacemereplaceme\nSheet2\nSheet3\n", extractor.Text);

            extractor.FormulasNotResults = (true);

            ClassicAssert.AreEqual("Sheet1\nreplaceme\nreplaceme\nCONCATENATE(A1,A2)\nSheet2\nSheet3\n", extractor.Text);

            extractor.Close();
        }
        [Test]
        public void TestStringFormula()
        {

            ExcelExtractor extractor = CreateExtractor("StringFormulas.xls");

            // Comes out as NaN if treated as a number
            // And as XYZ if treated as a string
            ClassicAssert.AreEqual("Sheet1\nXYZ\nSheet2\nSheet3\n", extractor.Text);

            extractor.FormulasNotResults = (true);

            ClassicAssert.AreEqual("Sheet1\nUPPER(\"xyz\")\nSheet2\nSheet3\n", extractor.Text);

            extractor.Close();
        }

        [Test]
        public void TestEventExtractor()
        {
            EventBasedExcelExtractor extractor;

            // First up, a simple file with string
            //  based formulas in it
            extractor = new EventBasedExcelExtractor(
                    new POIFSFileSystem(
                            HSSFTestDataSamples.OpenSampleFileStream("SimpleWithFormula.xls")
                    )
            );
            extractor.IncludeSheetNames = (true);

            String text = extractor.Text;
            ClassicAssert.AreEqual("Sheet1\nreplaceme\nreplaceme\nreplacemereplaceme\nSheet2\nSheet3\n", text);

            extractor.IncludeSheetNames = (false);
            extractor.FormulasNotResults = (true);

            text = extractor.Text;
            ClassicAssert.AreEqual("replaceme\nreplaceme\nCONCATENATE(A1,A2)\n", text);
            extractor.Close();

            // Now, a slightly longer file with numeric formulas
            extractor = new EventBasedExcelExtractor(
                    new POIFSFileSystem(
                            HSSFTestDataSamples.OpenSampleFileStream("sumifformula.xls")
                    )
            );
            extractor.IncludeSheetNames = (false);
            extractor.FormulasNotResults = (true);

            text = extractor.Text;
            ClassicAssert.AreEqual(
                    "1000\t1\tSUMIF(A1:A5,\">4000\",B1:B5)\n" +
                    "2000\t2\n" +
                    "3000\t3\n" +
                    "4000\t4\n" +
                    "5000\t5\n",
                    text
            );

            extractor.Close();
        }
        [Test]
        public void TestWithComments()
        {
            ExcelExtractor extractor = CreateExtractor("SimpleWithComments.xls");
            extractor.IncludeSheetNames = (false);

            // Check without comments
            ClassicAssert.AreEqual(
                    "1\tone\n" +
                    "2\ttwo\n" +
                    "3\tthree\n",
                    extractor.Text
            );

            // Now with
            extractor.IncludeCellComments = (true);
            ClassicAssert.AreEqual(
                    "1\tone Comment by Yegor Kozlov: Yegor Kozlov: first cell\n" +
                    "2\ttwo Comment by Yegor Kozlov: Yegor Kozlov: second cell\n" +
                    "3\tthree Comment by Yegor Kozlov: Yegor Kozlov: third cell\n",
                    extractor.Text
            );

            extractor.Close();
        }
        [Test]
        public void TestWithBlank()
        {
            ExcelExtractor extractor = CreateExtractor("MissingBits.xls");
            String def = extractor.Text;
            extractor.IncludeBlankCells = (true);
            String padded = extractor.Text;

            ClassicAssert.IsTrue(def.StartsWith(
                "Sheet1\n" +
                "&[TAB]\t\n" +
                "Hello\n" +
                "11\t23\n"
            ));

            ClassicAssert.IsTrue(padded.StartsWith(
                "Sheet1\n" +
                "&[TAB]\t\n" +
                "Hello\n" +
                "11\t\t\t23\n"
            ));

            extractor.Close();
        }
        [Test]
        //[Ignore("TODO FIX CI TESTS")]
        public void TestFormatting()
        {
            ExcelExtractor extractor = CreateExtractor("Formatting.xls");
            extractor.IncludeBlankCells = (false);
            extractor.IncludeSheetNames = false;
            String text = extractor.Text;

            // Note - not all the formats in the file
            //  actually quite match what they claim to
            //  be, as some are auto-local builtins...

            POITestCase.AssertStartsWith(text, "Dates, all 24th November 2006\n");
            POITestCase.AssertContains(text, "yyyy/mm/dd\t2006/11/24\n");
            POITestCase.AssertContains(text, "yyyy-mm-dd\t2006-11-24\n");
            POITestCase.AssertContains(text, "dd-mm-yy\t24-11-06\n");

            POITestCase.AssertContains(text, "nn.nn\t10.52\n");
            POITestCase.AssertContains(text, "nn.nnn\t10.520\n");
            POITestCase.AssertContains(text, "\u00a3nn.nn\t\u00a310.52\n");

            extractor.Close();
        }

        /**
         * Embded in a non-excel file
         */
        [Test]
        public void TestWithEmbeded()
        {
            POIFSFileSystem fs = new POIFSFileSystem(
                    POIDataSamples.GetDocumentInstance().OpenResourceAsStream("word_with_embeded.doc")
            );

            DirectoryNode objPool = (DirectoryNode)fs.Root.GetEntry("ObjectPool");
            DirectoryNode dirA = (DirectoryNode)objPool.GetEntry("_1269427460");
            DirectoryNode dirB = (DirectoryNode)objPool.GetEntry("_1269427461");

            HSSFWorkbook wbA = new HSSFWorkbook(dirA, fs, true);
            HSSFWorkbook wbB = new HSSFWorkbook(dirB, fs, true);

            ExcelExtractor exA = new ExcelExtractor(wbA);
            ExcelExtractor exB = new ExcelExtractor(wbB);

            ClassicAssert.AreEqual("Sheet1\nTest excel file\nThis is the first file\nSheet2\nSheet3\n",
                    exA.Text);
            ClassicAssert.AreEqual("Sample Excel", exA.SummaryInformation.Title);

            ClassicAssert.AreEqual("Sheet1\nAnother excel file\nThis is the second file\nSheet2\nSheet3\n",
                    exB.Text);
            ClassicAssert.AreEqual("Sample Excel 2", exB.SummaryInformation.Title);
        }

        /**
         * Excel embeded in excel
         */
        [Test]
        public void TestWithEmbededInOwn()
        {
            POIFSFileSystem fs = null;
            HSSFWorkbook wbA = null, wbB = null;
            ExcelExtractor exA = null, exB = null, ex = null;
            try
            {
                fs = new POIFSFileSystem(
                    HSSFTestDataSamples.OpenSampleFileStream("excel_with_embeded.xls")
                );
                DirectoryNode dirA = (DirectoryNode)fs.Root.GetEntry("MBD0000A3B5");
                DirectoryNode dirB = (DirectoryNode)fs.Root.GetEntry("MBD0000A3B4");

                wbA = new HSSFWorkbook(dirA, fs, true);
                wbB = new HSSFWorkbook(dirB, fs, true);

                exA = new ExcelExtractor(wbA);
                exB = new ExcelExtractor(wbB);

                ClassicAssert.AreEqual("Sheet1\nTest excel file\nThis is the first file\nSheet2\nSheet3\n",
                        exA.Text);
                ClassicAssert.AreEqual("Sample Excel", exA.SummaryInformation.Title);

                ClassicAssert.AreEqual("Sheet1\nTest excel file\nThis is the first file\nSheet2\nSheet3\n",
                        exA.Text);
                ClassicAssert.AreEqual("Sample Excel", exA.SummaryInformation.Title);

                // And the base file too
                ex = new ExcelExtractor(fs);
                ClassicAssert.AreEqual("Sheet1\nI have lots of embeded files in me\nSheet2\nSheet3\n",
                            ex.Text);
                ClassicAssert.AreEqual("Excel With Embeded", ex.SummaryInformation.Title);

            }
            finally
            {
                if (ex != null) ex.Close();
                if (exB != null) exB.Close();
                if (exA != null) exA.Close();
                if (wbB != null) wbB.Close();
                if (wbA != null) wbA.Close();
                if (fs != null) fs.Close();
            }

        }

        /**
         * Test that we get text from headers and footers
         */
        [Test]
        public void Test45538()
        {
            String[] files = {
                "45538_classic_Footer.xls", "45538_form_Footer.xls",
                "45538_classic_Header.xls", "45538_form_Header.xls"
            };
            for (int i = 0; i < files.Length; i++)
            {
                ExcelExtractor extractor = CreateExtractor(files[i]);
                String text = extractor.Text;
                POITestCase.AssertContains("Unable to find expected word in text\n" + text, text, "testdoc");
                POITestCase.AssertContains("Unable to find expected word in text\n" + text, text, "test phrase");
                extractor.Close();
            }
        }
        [Test]
        [Ignore("encrypt/decrypt were not implemented")]
        public void TestPassword()
        {
            Biff8EncryptionKey.CurrentUserPassword = ("password");
            ExcelExtractor extractor = CreateExtractor("password.xls");
            String text = extractor.Text;
            Biff8EncryptionKey.CurrentUserPassword = (null);

            POITestCase.AssertContains(text,"ZIP");
            extractor.Close();
        }
        [Test]
        public void TestNullPointerException()
        {
            ExcelExtractor extractor = CreateExtractor("ar.org.apsme.www_Form%20Inscripcion%20Curso%20NO%20Socios.xls");
            ClassicAssert.IsNotNull(extractor);
            ClassicAssert.IsNotNull(extractor.Text);
            extractor.Close();
        }
    }
}