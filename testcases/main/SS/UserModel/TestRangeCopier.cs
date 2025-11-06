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
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;
    using TestCases.SS;

    public abstract class TestRangeCopier
    {
        protected ISheet sheet1;
        protected ISheet sheet2;
        protected IWorkbook workbook;
        protected RangeCopier rangeCopier;
        protected RangeCopier transSheetRangeCopier;
        protected ITestDataProvider testDataProvider;

        protected void initSheets()
        {
            sheet1 = workbook.GetSheet("sheet1");
            sheet2 = workbook.GetSheet("sheet2");
        }

        [Test]
        public void CopySheetRangeWithoutFormulas()
        {
            CellRangeAddress rangeToCopy = CellRangeAddress.ValueOf("B1:C2");   //2x2
            CellRangeAddress destRange = CellRangeAddress.ValueOf("C2:D3");     //2x2
            rangeCopier.CopyRange(rangeToCopy, destRange);
            ClassicAssert.AreEqual("1.1", sheet1.GetRow(2).GetCell(2).ToString());
            ClassicAssert.AreEqual("2.1", sheet1.GetRow(2).GetCell(3).ToString());
        }

        [Test]
        public void TileTheRangeAway()
        {
            CellRangeAddress tileRange = CellRangeAddress.ValueOf("C4:D5");
            CellRangeAddress destRange = CellRangeAddress.ValueOf("F4:K5");
            rangeCopier.CopyRange(tileRange, destRange);
            ClassicAssert.AreEqual("1.3", GetCellContent(sheet1, "H4"));
            ClassicAssert.AreEqual("1.3", GetCellContent(sheet1, "J4"));
            ClassicAssert.AreEqual("$C1+G$2", GetCellContent(sheet1, "G5"));
            ClassicAssert.AreEqual("SUM(G3:I3)", GetCellContent(sheet1, "H5"));
            ClassicAssert.AreEqual("$C1+I$2", GetCellContent(sheet1, "I5"));
            ClassicAssert.AreEqual("", GetCellContent(sheet1, "L5"));  //out of borders
            ClassicAssert.AreEqual("", GetCellContent(sheet1, "G7")); //out of borders
        }

        [Test]
        public void TileTheRangeOver()
        {
            CellRangeAddress tileRange = CellRangeAddress.ValueOf("C4:D5");
            CellRangeAddress destRange = CellRangeAddress.ValueOf("A4:C5");
            rangeCopier.CopyRange(tileRange, destRange);
            ClassicAssert.AreEqual("1.3", GetCellContent(sheet1, "A4"));
            ClassicAssert.AreEqual("$C1+B$2", GetCellContent(sheet1, "B5"));
            ClassicAssert.AreEqual("SUM(B3:D3)", GetCellContent(sheet1, "C5"));
        }

        [Test]
        public void CopyRangeToOtherSheet()
        {
            ISheet destSheet = sheet2;
            CellRangeAddress tileRange = CellRangeAddress.ValueOf("C4:D5"); // on sheet1
            CellRangeAddress destRange = CellRangeAddress.ValueOf("F4:J6"); // on sheet2 
            transSheetRangeCopier.CopyRange(tileRange, destRange);
            ClassicAssert.AreEqual("1.3", GetCellContent(destSheet, "H4"));
            ClassicAssert.AreEqual("1.3", GetCellContent(destSheet, "J4"));
            ClassicAssert.AreEqual("$C1+G$2", GetCellContent(destSheet, "G5"));
            ClassicAssert.AreEqual("SUM(G3:I3)", GetCellContent(destSheet, "H5"));
            ClassicAssert.AreEqual("$C1+I$2", GetCellContent(destSheet, "I5"));
        }

        protected static string GetCellContent(ISheet sheet, string coordinates)
        {
            try
            {
                CellReference p = new CellReference(coordinates);
                return sheet.GetRow(p.Row).GetCell(p.Col).ToString();
            }
            catch(NullReferenceException e)
            { // row or cell does not exist
                return "";
            }
        }
    }
}


