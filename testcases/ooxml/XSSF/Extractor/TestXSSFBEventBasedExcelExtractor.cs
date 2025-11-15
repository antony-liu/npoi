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

namespace TestCases.XSSF.Extractor
{
    using NPOI.XSSF;
    using NPOI.XSSF.Extractor;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;

    /// <summary>
    /// Tests for <see cref="XSSFBEventBasedExcelExtractor"/>
    /// </summary>
    [TestFixture]
    public class TestXSSFBEventBasedExcelExtractor
    {
        protected XSSFEventBasedExcelExtractor GetExtractor(string sampleName)
        {
            return new XSSFBEventBasedExcelExtractor(XSSFTestDataSamples.
                    OpenSamplePackage(sampleName));
        }

        /// <summary>
        /// Get text out of the simple file
        /// </summary>
        [Test]
        public void TestGetSimpleText()
        {
            // a very simple file
            XSSFEventBasedExcelExtractor extractor = GetExtractor("sample.xlsb");
            extractor.IncludeCellComments = true;
            _ = extractor.Text;

            string text = extractor.Text;
            ClassicAssert.IsTrue(text.Length > 0);

            // Check sheet names
            POITestCase.AssertStartsWith(text, "Sheet1");
            POITestCase.AssertEndsWith(text, "Sheet3\n");

            // Now without, will have text
            extractor.IncludeSheetNames = false;
            text = extractor.Text;
            string CHUNK1 =
                "Lorem\t111\n" +
                        "ipsum\t222\n" +
                        "dolor\t333\n" +
                        "sit\t444\n" +
                        "amet\t555\n" +
                        "consectetuer\t666\n" +
                        "adipiscing\t777\n" +
                        "elit\t888\n" +
                        "Nunc\t999\n";
            string CHUNK2 =
                "The quick brown fox jumps over the lazy dog\n" +
                        "hello, xssf	hello, xssf\n" +
                        "hello, xssf	hello, xssf\n" +
                        "hello, xssf	hello, xssf\n" +
                        "hello, xssf	hello, xssf\n";
            ClassicAssert.AreEqual(
                    CHUNK1 +
                            "at\t4995\n" +
                            CHUNK2
                    , text);

        }


        /// <summary>
        /// Test text extraction from text box using GetShapes()
        /// </summary>
        /// <exception cref="Exception"></exception>
        [Test]
        public void TestShapes()
        {
            XSSFEventBasedExcelExtractor ooxmlExtractor = GetExtractor("WithTextBox.xlsb");

            try
            {
                string text = ooxmlExtractor.Text;
                POITestCase.AssertContains(text, "Line 1");
                POITestCase.AssertContains(text, "Line 2");
                POITestCase.AssertContains(text, "Line 3");
            }
            finally
            {
                ooxmlExtractor.Close();
            }
        }

        [Test]
        public void TestBeta()
        {
            XSSFEventBasedExcelExtractor extractor = GetExtractor("Simple.xlsb");
            extractor.IncludeCellComments = true;
            string text = extractor.Text;
            POITestCase.AssertContains(text,
                    "This is an example spreadsheet created with Microsoft Excel 2007 Beta 2.");
        }
        [Test]
        public void Test62815()
        {
            //test file based on http://oss.sheetjs.com/test_files/RkNumber.xlsb
            XSSFEventBasedExcelExtractor extractor = GetExtractor("62815.xlsb");
            extractor.IncludeCellComments = true;
            String[] rows = extractor.Text.Split("\r\n".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            ClassicAssert.AreEqual(283, rows.Length);
            StreamReader reader = new StreamReader(XSSFTestDataSamples.GetSampleFile("62815.xlsb.txt").Open(FileMode.Open));
            String line = reader.ReadLine();
            for(int i = 0; i<rows.Length; i++)
            {
                ClassicAssert.AreEqual(line, rows[i]);
                line = reader.ReadLine();
                while(line != null && line.StartsWith("#"))
                {
                    line = reader.ReadLine();
                }
            }
        }
    }
}


