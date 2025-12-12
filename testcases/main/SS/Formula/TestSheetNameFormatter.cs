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

using NUnit.Framework.Legacy;

namespace TestCases.SS.Formula
{

    using NPOI.SS.Formula;
    using NPOI.Util;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;
    using System;
    using System.IO;
    using System.Text;

    /**
     * Tests for {@link SheetNameFormatter}
     * 
     * @author Josh Micich
     */
    [TestFixture]
    public class TestSheetNameFormatter
    {
        /**
         * Tests main public method 'format' 
         */
        [Test]
        public void TestFormat()
        {

            ConfirmFormat("abc", "abc");
            ConfirmFormat("123", "'123'");

            ConfirmFormat("my sheet", "'my sheet'"); // space
            ConfirmFormat("A:MEM", "'A:MEM'"); // colon

            ConfirmFormat("O'Brian", "'O''Brian'"); // single quote Gets doubled


            ConfirmFormat("3rdTimeLucky", "'3rdTimeLucky'"); // digit in first pos
            ConfirmFormat("_", "_"); // plain underscore OK
            ConfirmFormat("my_3rd_sheet", "my_3rd_sheet"); // underscores and digits OK
            ConfirmFormat("A12220", "'A12220'");
            ConfirmFormat("TAXRETURN19980415", "TAXRETURN19980415");

            ConfirmFormat(null, "#REF");
        }

        private static void ConfirmFormat(string rawSheetName, string expectedSheetNameEncoding)
        {
            // test all variants

            ClassicAssert.AreEqual(expectedSheetNameEncoding, SheetNameFormatter.Format(rawSheetName));

            StringBuilder sb = new StringBuilder();
            SheetNameFormatter.AppendFormat(sb, rawSheetName);
            ClassicAssert.AreEqual(expectedSheetNameEncoding, sb.ToString());
            sb = new StringBuilder();
            SheetNameFormatter.AppendFormat(sb, rawSheetName);
            ClassicAssert.AreEqual(expectedSheetNameEncoding, sb.ToString());

            StringBuilder sbf = new StringBuilder();
            //noinspection deprecation
            SheetNameFormatter.AppendFormat(sbf, rawSheetName);
            ClassicAssert.AreEqual(expectedSheetNameEncoding, sbf.ToString());
        }

        [Test]
        public void TestFormatWithWorkbookName()
        {

            confirmFormat("abc", "abc", "[abc]abc");
            confirmFormat("abc", "123", "'[abc]123'");

            confirmFormat("abc", "my sheet", "'[abc]my sheet'"); // space
            confirmFormat("abc", "A:MEM", "'[abc]A:MEM'"); // colon

            confirmFormat("abc", "O'Brian", "'[abc]O''Brian'"); // single quote Gets doubled

            confirmFormat("abc", "3rdTimeLucky", "'[abc]3rdTimeLucky'"); // digit in first pos
            confirmFormat("abc", "_", "[abc]_"); // plain underscore OK
            confirmFormat("abc", "my_3rd_sheet", "[abc]my_3rd_sheet"); // underscores and digits OK
            confirmFormat("abc", "A12220", "'[abc]A12220'");
            confirmFormat("abc", "TAXRETURN19980415", "[abc]TAXRETURN19980415");

            confirmFormat("abc", null, "[abc]#REF");
            confirmFormat(null, "abc", "[#REF]abc");
            confirmFormat(null, null, "[#REF]#REF");
        }

        private static void confirmFormat(string workbookName, string rawSheetName, string expectedSheetNameEncoding)
        {
            // test all variants

            StringBuilder sb = new StringBuilder();
            SheetNameFormatter.AppendFormat(sb, workbookName, rawSheetName);
            ClassicAssert.AreEqual(expectedSheetNameEncoding, sb.ToString());

            sb = new StringBuilder();
            SheetNameFormatter.AppendFormat(sb, workbookName, rawSheetName);
            ClassicAssert.AreEqual(expectedSheetNameEncoding, sb.ToString());

            StringBuilder sbf = new StringBuilder();
            //noinspection deprecation
            SheetNameFormatter.AppendFormat(sbf, workbookName, rawSheetName);
            ClassicAssert.AreEqual(expectedSheetNameEncoding, sbf.ToString());
        }

        [Test]
        public void TestFormatException()
        {
            //        Appendable mock = new Appendable() {

            //        public Appendable append(CharSequence csq)
            //    {
            //        throw new IOException("Test exception");
            //    }
            //    public Appendable append(CharSequence csq, int start, int end)
            //    {
            //        throw new IOException("Test exception");
            //    }
            //    public Appendable append(char c)
            //    {
            //        throw new IOException("Test exception");
            //    }
            //};

            //try
            //{
            //    SheetNameFormatter.AppendFormat(mock, null, null);
            //    ClassicAssert.Fail("Should catch exception here");
            //}
            //catch(RuntimeException e)
            //{
            //    // expected here
            //}

            //try
            //{
            //    SheetNameFormatter.AppendFormat(mock, null);
            //    ClassicAssert.Fail("Should catch exception here");
            //}
            //catch(RuntimeException e)
            //{
            //    // expected here
            //}
        }
        [Test]
        public void TestBooleanLiterals()
        {
            ConfirmFormat("TRUE", "'TRUE'");
            ConfirmFormat("FALSE", "'FALSE'");
            ConfirmFormat("True", "'True'");
            ConfirmFormat("fAlse", "'fAlse'");

            ConfirmFormat("Yes", "Yes");
            ConfirmFormat("No", "No");
        }

        private static void ConfirmCellNameMatch(String rawSheetName, bool expected)
        {
            ClassicAssert.AreEqual(expected, SheetNameFormatter.NameLooksLikePlainCellReference(rawSheetName));
        }

        /**
         * Tests functionality to determine whether a sheet name Containing only letters and digits
         * would look (to Excel) like a cell name.
         */
        [Test]
        public void TestLooksLikePlainCellReference()
        {

            ConfirmCellNameMatch("A1", true);
            ConfirmCellNameMatch("a111", true);
            ConfirmCellNameMatch("AA", false);
            ConfirmCellNameMatch("aa1", true);
            ConfirmCellNameMatch("A1A", false);
            ConfirmCellNameMatch("A1A1", false);
            ConfirmCellNameMatch("Sh3", false);
            ConfirmCellNameMatch("SALES20080101", false); // out of range
        }

        private static void ConfirmCellRange(String text, int numberOfPrefixLetters, bool expected)
        {
            String prefix = text.Substring(0, numberOfPrefixLetters);
            String suffix = text.Substring(numberOfPrefixLetters);
            ClassicAssert.AreEqual(expected, SheetNameFormatter.CellReferenceIsWithinRange(prefix, suffix));
        }

        /**
         * Tests exact boundaries for names that look very close to cell names (i.e. contain 1 or more
         * letters followed by one or more digits).
         */
        [Test]
        public void TestCellRange()
        {
            ConfirmCellRange("A1", 1, true);
            ConfirmCellRange("a111", 1, true);
            ConfirmCellRange("A65536", 1, true);
            ConfirmCellRange("A65537", 1, false);
            ConfirmCellRange("iv1", 2, true);
            ConfirmCellRange("IW1", 2, false);
            ConfirmCellRange("AAA1", 3, false);
            ConfirmCellRange("a111", 1, true);
            ConfirmCellRange("Sheet1", 6, false);
            ConfirmCellRange("iV65536", 2, true);  // max cell in Excel 97-2003
            ConfirmCellRange("IW65537", 2, false);
        }
    }

}