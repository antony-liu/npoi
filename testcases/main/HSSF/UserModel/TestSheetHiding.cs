/*
* Licensed to the Apache Software Foundation (ASF) under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for Additional information regarding copyright ownership.
* The ASF licenses this file to You under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed under the License is distributed on an "AS IS" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations under the License.
*/


namespace TestCases.HSSF.UserModel
{
    using System;
    using TestCases.SS.UserModel;
    using NUnit.Framework;using NUnit.Framework.Legacy;
    using NPOI.HSSF.UserModel;
    using NPOI.HSSF.Model;
    using NPOI.SS.UserModel;
    using NPOI.Util;

    [TestFixture]
    public class TestSheetHiding:BaseTestSheetHiding
    {
        public TestSheetHiding():base(HSSFITestDataProvider.Instance,
                    "TwoSheetsOneHidden.xls", "TwoSheetsNoneHidden.xls")
        {

        }

        [Test]
        public void TestInternalWorkbookHidden()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            wb.CreateSheet("MySheet");
            InternalWorkbook intWb = wb.Workbook;

            ClassicAssert.IsFalse(intWb.IsSheetHidden(0));
            ClassicAssert.IsFalse(intWb.IsSheetVeryHidden(0));
            ClassicAssert.AreEqual(SheetVisibility.Visible, intWb.GetSheetVisibility(0));

            intWb.SetSheetHidden(0, SheetVisibility.Hidden);
            ClassicAssert.IsTrue(intWb.IsSheetHidden(0));
            ClassicAssert.IsFalse(intWb.IsSheetVeryHidden(0));
            ClassicAssert.AreEqual(SheetVisibility.Hidden, intWb.GetSheetVisibility(0));

            // InternalWorkbook currently behaves slightly different
            // than HSSFWorkbook, but it should not have effect in normal usage
            // as checked limits are more strict in HSSFWorkbook

            // check sheet-index with one more will work and add the sheet
            intWb.SetSheetHidden(1, SheetVisibility.Hidden);
            ClassicAssert.IsTrue(intWb.IsSheetHidden(1));
            ClassicAssert.IsFalse(intWb.IsSheetVeryHidden(1));
            ClassicAssert.AreEqual(SheetVisibility.Hidden, intWb.GetSheetVisibility(1));

            // check sheet-index with index out of bounds => throws exception
            try
            {
                wb.SetSheetVisibility(10, SheetVisibility.Hidden);
                ClassicAssert.Fail("Should catch exception here");
            }
            catch(ArgumentException)
            {
                // expected here
            }
        }
    }
}