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

namespace TestCases.XSSF.UserModel
{
    using NPOI.SS.UserModel;
    using NPOI.Util;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;
    [TestFixture]
    public class TestXSSFUnicodeSurrogates
    {

        // "𝝊𝝋𝝌𝝍𝝎𝝏𝝐𝝑𝝒𝝓𝝔𝝕𝝖𝝗𝝘𝝙𝝚𝝛𝝜𝝝𝝞𝝟𝝠𝝡𝝢𝝣𝝤𝝥𝝦𝝧𝝨𝝩𝝪𝝫𝝬𝝭𝝮𝝯𝝰𝝱𝝲𝝳𝝴𝝵𝝶𝝷𝝸𝝹𝝺";
        private static string unicodeText =
        "\uD835\uDF4A\uD835\uDF4B\uD835\uDF4C\uD835\uDF4D\uD835\uDF4E\uD835\uDF4F\uD835\uDF50\uD835" +
        "\uDF51\uD835\uDF52\uD835\uDF53\uD835\uDF54\uD835\uDF55\uD835\uDF56\uD835\uDF57\uD835\uDF58" +
        "\uD835\uDF59\uD835\uDF5A\uD835\uDF5B\uD835\uDF5C\uD835\uDF5D\uD835\uDF5E\uD835\uDF5F\uD835" +
        "\uDF60\uD835\uDF61\uD835\uDF62\uD835\uDF63\uD835\uDF64\uD835\uDF65\uD835\uDF66\uD835\uDF67" +
        "\uD835\uDF68\uD835\uDF69\uD835\uDF6A\uD835\uDF6B\uD835\uDF6C\uD835\uDF6D\uD835\uDF6E\uD835" +
        "\uDF6F\uD835\uDF70\uD835\uDF71\uD835\uDF72\uD835\uDF73\uD835\uDF74\uD835\uDF75\uD835\uDF76" +
        "\uD835\uDF77\uD835\uDF78\uD835\uDF79\uD835\uDF7A";

        [Test]
        public void TestWriteUnicodeSurrogates()
        {
            string sheetName = "Sheet1";
            FileInfo tf = TempFile.CreateTempFile("poi-xmlbeans-test", ".xlsx");
            XSSFWorkbook wb = new XSSFWorkbook();
            try
            {
                ISheet sheet = wb.CreateSheet(sheetName);
                IRow row = sheet.CreateRow(0);
                ICell cell = row.CreateCell(0);
                cell.SetCellValue(unicodeText);
                using(FileStream os = new FileStream(tf.FullName, FileMode.OpenOrCreate))
                {
                    wb.Write(os);
                }
                using(FileStream fis = new FileStream(tf.FullName, FileMode.Open))
                {
                    XSSFWorkbook wb2 = new XSSFWorkbook(fis);
                    ISheet sheet2 = wb2.GetSheet(sheetName);
                    ICell cell2 = sheet2.GetRow(0).GetCell(0);
                    ClassicAssert.AreEqual(unicodeText, cell2.StringCellValue);
                }

            }
            finally
            {
                tf.Delete();
            }
        }
    }
}


