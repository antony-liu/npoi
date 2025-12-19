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

using NUnit.Framework;
using NUnit.Framework.Legacy;

namespace TestCases.HSSF.UserModel
{
    using NPOI.HSSF.Record;
    using NPOI.HSSF.UserModel;
    using NPOI.HSSF.Util;
    using NPOI.SS.UserModel;
    using NPOI.Util;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;
    using System;
    using System.IO;
    using TestCases.HSSF;

    /**
     * @author aviks
     * 
     * This Testcase contains Tests for bugs that are yet to be fixed. Therefore,
     * the standard ant Test target does not run these Tests. Run this Testcase with
     * the single-Test target. The names of the Tests usually correspond to the
     * Bugzilla id's PLEASE MOVE Tests from this class to TestBugs once the bugs are
     * fixed, so that they are then run automatically.
     */
    [TestFixture]
    [Ignore("UnfixedBugs")]
    public class TestUnfixedBugs
    {
        [Test]
        public void TestFormulaRecordAggregate_1()
        {
            // Assert.Fails at formula "=MEHRFACH.OPERATIONEN(E$3;$B$5;$D4)"
            IWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("44958_1.xls");
            for(int i = 0; i < wb.NumberOfSheets; i++)
            {
                ISheet sheet = wb.GetSheetAt(i);
                ClassicAssert.IsNotNull(wb.GetSheet(sheet.SheetName));
                sheet.GroupColumn((short) 4, (short) 5);
                sheet.SetColumnGroupCollapsed(4, true);
                sheet.SetColumnGroupCollapsed(4, false);

                foreach(IRow row in sheet)
                {
                    foreach(ICell cell in row)
                    {
                        try
                        {
                            ClassicAssert.IsNotNull(cell.ToString());
                        }
                        catch(Exception e)
                        {
                            throw new Exception("While handling: " + sheet.SheetName + "/" + row.RowNum + "/" + cell.ColumnIndex, e);
                        }
                    }
                }
            }
        }


        [Test]
        public void TestFormulaRecordAggregate()
        {
            // Assert.Fails at formula "=MEHRFACH.OPERATIONEN(E$3;$B$5;$D4)"
            IWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("44958.xls");
            {
                for(int i = 0; i < wb.NumberOfSheets; i++)
                {
                    ISheet sheet = wb.GetSheetAt(i);
                    ClassicAssert.IsNotNull(wb.GetSheet(sheet.SheetName));
                    sheet.GroupColumn((short) 4, (short) 5);
                    sheet.SetColumnGroupCollapsed(4, true);
                    sheet.SetColumnGroupCollapsed(4, false);

                    foreach(IRow row in sheet)
                    {
                        foreach(ICell cell in row)
                        {
                            try
                            {
                                ClassicAssert.IsNotNull(cell.ToString());
                            }
                            catch(Exception e)
                            {
                                throw new Exception("While handling: " + sheet.SheetName + "/" + row.RowNum + "/" + cell.ColumnIndex, e);
                            }
                        }
                    }
                }
            }
        }

        [Test]
        public void TestBug57074()
        {
            IWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("57074.xls");
            ISheet sheet = wb.GetSheet("Sheet1");
            IRow row = sheet.GetRow(0);
            ICell cell = row.GetCell(0);

            HSSFColor bgColor = (HSSFColor) cell.CellStyle.FillBackgroundColorColor;
            string bgColorStr = bgColor.GetTriplet()[0]+", "+bgColor.GetTriplet()[1]+", "+bgColor.GetTriplet()[2];
            //Console.WriteLine(bgColorStr);
            ClassicAssert.AreEqual("215, 228, 188", bgColorStr);

            HSSFColor fontColor = (HSSFColor) cell.CellStyle.FillForegroundColorColor;
            string fontColorStr = fontColor.GetTriplet()[0]+", "+fontColor.GetTriplet()[1]+", "+fontColor.GetTriplet()[2];
            //Console.WriteLine(fontColorStr);
            ClassicAssert.AreEqual("0, 128, 128", fontColorStr);
            wb.Close();
        }
    }
}