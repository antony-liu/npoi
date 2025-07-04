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

using System;
using NUnit.Framework;using NUnit.Framework.Legacy;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using NPOI.SS.UserModel;
using NPOI.XSSF.Model;
using NPOI.XSSF;

namespace TestCases.XSSF.Model
{
    [TestFixture]
    public class TestStylesTable
    {
        private String testFile = "Formatting.xlsx";
        private static String customDataFormat = "YYYY-mm-dd";

        [SetUp]
        public static void assumeCustomDataFormatIsNotBuiltIn()
        {
            ClassicAssert.AreEqual(-1, BuiltinFormats.GetBuiltinFormat(customDataFormat));
        }

        [Test]
        public void TestCreateNew()
        {
            StylesTable st = new StylesTable();

            // Check defaults
            ClassicAssert.IsNotNull(st.GetCTStylesheet());
            ClassicAssert.AreEqual(1, st.XfsSize);
            ClassicAssert.AreEqual(1, st.StyleXfsSize);
            ClassicAssert.AreEqual(0, st.NumDataFormats);
        }
        [Test]
        public void TestCreateSaveLoad()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            StylesTable st = wb.GetStylesSource();

            ClassicAssert.IsNotNull(st.GetCTStylesheet());
            ClassicAssert.AreEqual(1, st.XfsSize);
            ClassicAssert.AreEqual(1, st.StyleXfsSize);
            ClassicAssert.AreEqual(0, st.NumDataFormats);

            st = ((XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(wb)).GetStylesSource();

            ClassicAssert.IsNotNull(st.GetCTStylesheet());
            ClassicAssert.AreEqual(1, st.XfsSize);
            ClassicAssert.AreEqual(1, st.StyleXfsSize);
            ClassicAssert.AreEqual(0, st.NumDataFormats);

            ClassicAssert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(wb));
        }
        [Test]
        public void TestLoadExisting()
        {
            XSSFWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook(testFile);
            ClassicAssert.IsNotNull(workbook.GetStylesSource());

            StylesTable st = workbook.GetStylesSource();

            doTestExisting(st);

            ClassicAssert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(workbook));
        }
        [Test]
        public void TestLoadSaveLoad()
        {
            XSSFWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook(testFile);
            ClassicAssert.IsNotNull(workbook.GetStylesSource());

            StylesTable st = workbook.GetStylesSource();
            doTestExisting(st);

            st = ((XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(workbook)).GetStylesSource();
            doTestExisting(st);
        }
        public void doTestExisting(StylesTable st)
        {
            // Check contents
            ClassicAssert.IsNotNull(st.GetCTStylesheet());
            ClassicAssert.AreEqual(11, st.XfsSize);
            ClassicAssert.AreEqual(1, st.StyleXfsSize);
            ClassicAssert.AreEqual(8, st.NumDataFormats);

            ClassicAssert.AreEqual(2, st.GetFonts().Count);
            ClassicAssert.AreEqual(2, st.GetFills().Count);
            ClassicAssert.AreEqual(1, st.GetBorders().Count);

            ClassicAssert.AreEqual("yyyy/mm/dd", st.GetNumberFormatAt((short)165));
            ClassicAssert.AreEqual("yy/mm/dd", st.GetNumberFormatAt((short)167));

            ClassicAssert.IsNotNull(st.GetStyleAt(0));
            ClassicAssert.IsNotNull(st.GetStyleAt(1));
            ClassicAssert.IsNotNull(st.GetStyleAt(2));

            ClassicAssert.AreEqual(0, st.GetStyleAt(0).DataFormat);
            ClassicAssert.AreEqual(14, st.GetStyleAt(1).DataFormat);
            ClassicAssert.AreEqual(0, st.GetStyleAt(2).DataFormat);
            ClassicAssert.AreEqual(165, st.GetStyleAt(3).DataFormat);

            ClassicAssert.AreEqual("yyyy/mm/dd", st.GetStyleAt(3).GetDataFormatString());
        }
        [Test]
        public void TestPopulateNew()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            StylesTable st = wb.GetStylesSource();

            ClassicAssert.IsNotNull(st.GetCTStylesheet());
            ClassicAssert.AreEqual(1, st.XfsSize);
            ClassicAssert.AreEqual(1, st.StyleXfsSize);
            ClassicAssert.AreEqual(0, st.NumDataFormats);

            int nf1 = st.PutNumberFormat("yyyy-mm-dd");
            int nf2 = st.PutNumberFormat("yyyy-mm-DD");
            ClassicAssert.AreEqual(nf1, st.PutNumberFormat("yyyy-mm-dd"));

            st.PutStyle(new XSSFCellStyle(st));

            // Save and re-load
            st = ((XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(wb)).GetStylesSource();

            ClassicAssert.IsNotNull(st.GetCTStylesheet());
            ClassicAssert.AreEqual(2, st.XfsSize);
            ClassicAssert.AreEqual(1, st.StyleXfsSize);
            ClassicAssert.AreEqual(2, st.NumDataFormats);

            ClassicAssert.AreEqual("yyyy-mm-dd", st.GetNumberFormatAt((short)nf1));
            ClassicAssert.AreEqual(nf1, st.PutNumberFormat("yyyy-mm-dd"));
            ClassicAssert.AreEqual(nf2, st.PutNumberFormat("yyyy-mm-DD"));

            ClassicAssert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(wb));
        }
        [Test]
        public void TestPopulateExisting()
        {
            XSSFWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook(testFile);
            ClassicAssert.IsNotNull(workbook.GetStylesSource());

            StylesTable st = workbook.GetStylesSource();
            ClassicAssert.AreEqual(11, st.XfsSize);
            ClassicAssert.AreEqual(1, st.StyleXfsSize);
            ClassicAssert.AreEqual(8, st.NumDataFormats);

            int nf1 = st.PutNumberFormat("YYYY-mm-dd");
            int nf2 = st.PutNumberFormat("YYYY-mm-DD");
            ClassicAssert.AreEqual(nf1, st.PutNumberFormat("YYYY-mm-dd"));

            st = ((XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(workbook)).GetStylesSource();

            ClassicAssert.AreEqual(11, st.XfsSize);
            ClassicAssert.AreEqual(1, st.StyleXfsSize);
            ClassicAssert.AreEqual(10, st.NumDataFormats);

            ClassicAssert.AreEqual("YYYY-mm-dd", st.GetNumberFormatAt((short)nf1));
            ClassicAssert.AreEqual(nf1, st.PutNumberFormat("YYYY-mm-dd"));
            ClassicAssert.AreEqual(nf2, st.PutNumberFormat("YYYY-mm-DD"));

            ClassicAssert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(workbook));
        }

        [Test]
        public void ExceedNumberFormatLimit()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            try
            {
                StylesTable styles = wb.GetStylesSource();
                for (int i = 0; i < styles.MaxNumberOfDataFormats; i++)
                {
                    wb.GetStylesSource().PutNumberFormat("\"test" + i + " \"0");
                }
                try
                {
                    wb.GetStylesSource().PutNumberFormat("\"anotherformat \"0");
                }
                catch (InvalidOperationException e)
                {
                    if (e.Message.StartsWith("The maximum number of Data Formats was exceeded."))
                    {
                        //expected
                    }
                    else
                    {
                        throw;
                    }
                }
            }
            finally
            {
                wb.Close();
            }
        }

        private static void AssertNotContainsKey<K, V>(SortedDictionary<K, V> map, K key)
        {
            ClassicAssert.IsFalse(map.ContainsKey(key));
        }
        private static void AssertNotContainsValue<K, V>(SortedDictionary<K, V> map, V value)
        {
            ClassicAssert.IsFalse(map.ContainsValue(value));
        }

        [Test]
        public void RemoveNumberFormat()
        {
            XSSFWorkbook wb1 = new XSSFWorkbook();
            try
            {
                String fmt = customDataFormat;
                short fmtIdx = (short)wb1.GetStylesSource().PutNumberFormat(fmt);

                ICell cell = wb1.CreateSheet("test").CreateRow(0).CreateCell(0);
                cell.SetCellValue(5.25);
                ICellStyle style = wb1.CreateCellStyle();
                style.DataFormat = fmtIdx;
                cell.CellStyle = style;

                ClassicAssert.AreEqual(fmt, cell.CellStyle.GetDataFormatString());
                ClassicAssert.AreEqual(fmt, wb1.GetStylesSource().GetNumberFormatAt(fmtIdx));

                // remove the number format from the workbook
                wb1.GetStylesSource().RemoveNumberFormat(fmt);

                // number format in CellStyles should be restored to default number format
                short defaultFmtIdx = 0;
                String defaultFmt = BuiltinFormats.GetBuiltinFormat(0);
                ClassicAssert.AreEqual(defaultFmtIdx, style.DataFormat);
                ClassicAssert.AreEqual(defaultFmt, style.GetDataFormatString());

                // The custom number format should be entirely removed from the workbook
                SortedDictionary<short, String> numberFormats = wb1.GetStylesSource().GetNumberFormats() as SortedDictionary<short, String>;
                AssertNotContainsKey(numberFormats, fmtIdx);
                AssertNotContainsValue(numberFormats, fmt);

                // The default style shouldn't be added back to the styles source because it's built-in
                ClassicAssert.AreEqual(0, wb1.GetStylesSource().NumDataFormats);

                cell = null;
                style = null;
                numberFormats = null;
                XSSFWorkbook wb2 = XSSFTestDataSamples.WriteOutCloseAndReadBack(wb1);

                cell = wb2.GetSheet("test").GetRow(0).GetCell(0);
                style = cell.CellStyle;

                // number format in CellStyles should be restored to default number format
                ClassicAssert.AreEqual(defaultFmtIdx, style.DataFormat);
                ClassicAssert.AreEqual(defaultFmt, style.GetDataFormatString());

                // The custom number format should be entirely removed from the workbook
                numberFormats = wb2.GetStylesSource().GetNumberFormats() as SortedDictionary<short, String>;
                AssertNotContainsKey(numberFormats, fmtIdx);
                AssertNotContainsValue(numberFormats, fmt);

                // The default style shouldn't be added back to the styles source because it's built-in
                ClassicAssert.AreEqual(0, wb2.GetStylesSource().NumDataFormats);

                wb2.Close();
            }
            finally
            {
                wb1.Close();
            }
        }

        [Test]
        public void MaxNumberOfDataFormats()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            try
            {
                StylesTable styles = wb.GetStylesSource();

                // Check default limit
                int n = styles.MaxNumberOfDataFormats;
                // https://support.office.com/en-us/article/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
                ClassicAssert.IsTrue(200 <= n);
                ClassicAssert.IsTrue(n <= 250);

                // Check upper limit
                n = int.MaxValue;
                styles.MaxNumberOfDataFormats = (n);
                ClassicAssert.AreEqual(n, styles.MaxNumberOfDataFormats);

                // Check negative (illegal) limits
                try
                {
                    styles.MaxNumberOfDataFormats = (-1);
                    Assert.Fail("Expected to get an IllegalArgumentException(\"Maximum Number of Data Formats must be greater than or equal to 0\")");
                }
                catch (ArgumentException e)
                {
                    if (e.Message.StartsWith("Maximum Number of Data Formats must be greater than or equal to 0"))
                    {
                        // expected
                    }
                    else
                    {
                        throw;
                    }
                }
            }
            finally
            {
                wb.Close();
            }
        }

        [Test]
        public void AddDataFormatsBeyondUpperLimit()
        {
            XSSFWorkbook wb = new XSSFWorkbook();

            try
            {
                StylesTable styles = wb.GetStylesSource();
                styles.MaxNumberOfDataFormats = (0);

                // Try adding a format beyond the upper limit
                try
                {
                    styles.PutNumberFormat("\"test \"0");
                    Assert.Fail("Expected to raise InvalidOperationException");
                }
                catch (InvalidOperationException e)
                {
                    if (e.Message.StartsWith("The maximum number of Data Formats was exceeded."))
                    {
                        // expected
                    }
                    else
                    {
                        throw;
                    }
                }
            }
            finally
            {
                wb.Close();
            }
        }

        [Test]
        public void DecreaseUpperLimitBelowCurrentNumDataFormats()
        {
            XSSFWorkbook wb = new XSSFWorkbook();

            try
            {
                StylesTable styles = wb.GetStylesSource();
                styles.PutNumberFormat(customDataFormat);

                // Try decreasing the upper limit below the current number of formats
                try
                {
                    styles.MaxNumberOfDataFormats = (0);
                    Assert.Fail("Expected to raise InvalidOperationException");
                }
                catch (InvalidOperationException e)
                {
                    if (e.Message.StartsWith("Cannot set the maximum number of data formats less than the current quantity."))
                    {
                        // expected
                    }
                    else
                    {
                        throw;
                    }
                }
            }
            finally
            {
                wb.Close();
            }
        }
        [Test]
        public void TestLoadWithAlternateContent()
        {
            XSSFWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("style-alternate-content.xlsx");
            ClassicAssert.IsNotNull(workbook.GetStylesSource());

            StylesTable st = workbook.GetStylesSource();

            ClassicAssert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(workbook));
        }
    }
}
