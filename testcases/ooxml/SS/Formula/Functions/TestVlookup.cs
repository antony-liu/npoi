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

namespace TestCases.SS.Formula.Functions
{
    using NPOI.SS.UserModel;
    using NPOI.XSSF;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;

    /// <summary>
    /// Test the VLOOKUP function
    /// </summary>
    [TestFixture]
    public class TestVlookup
    {

        [Test]
        public void TestFullColumnAreaRef61841()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("VLookupFullColumn.xlsx");
            IFormulaEvaluator feval = wb.GetCreationHelper().CreateFormulaEvaluator();
            feval.EvaluateAll();
            ClassicAssert.AreEqual("Value1", feval.Evaluate(wb.GetSheetAt(0).GetRow(3).GetCell(1)).StringValue, "Wrong lookup value");
            ClassicAssert.AreEqual(CellType.Error, feval.Evaluate(wb.GetSheetAt(0).GetRow(4).GetCell(1)).CellType, "Lookup should return #N/A");
        }

        [Test]
        public void Bug62275_true()
        {
            IWorkbook wb = new XSSFWorkbook();
            try
            {
                ISheet sheet = wb.CreateSheet();
                IRow row = sheet.CreateRow(0);

                ICell cell = row.CreateCell(0);
                cell.SetCellFormula("vlookup(A2,B1:B5,2,true)");

                ICreationHelper createHelper = wb.GetCreationHelper();
                IFormulaEvaluator eval = createHelper.CreateFormulaEvaluator();
                CellValue value = eval.Evaluate(cell);

                ClassicAssert.IsFalse(value.BooleanValue);
            }
            finally
            {
                wb.Close();
            }
        }
        [Test]
        public void Bug62275_false()
        {
            IWorkbook wb = new XSSFWorkbook();
            try
            {
                ISheet sheet = wb.CreateSheet();
                IRow row = sheet.CreateRow(0);

                ICell cell = row.CreateCell(0);
                cell.SetCellFormula("vlookup(A2,B1:B5,2,false)");

                ICreationHelper crateHelper = wb.GetCreationHelper();
                IFormulaEvaluator eval = crateHelper.CreateFormulaEvaluator();
                CellValue value = eval.Evaluate(cell);

                ClassicAssert.IsFalse(value.BooleanValue);
            }
            finally
            {
                wb.Close();
            }
        }

        [Test]
        public void Bug62275_empty_3args()
        {
            IWorkbook wb = new XSSFWorkbook();
            try
            {
                ISheet sheet = wb.CreateSheet();
                IRow row = sheet.CreateRow(0);

                ICell cell = row.CreateCell(0);
                cell.SetCellFormula("vlookup(A2,B1:B5,2,)");

                ICreationHelper crateHelper = wb.GetCreationHelper();
                IFormulaEvaluator eval = crateHelper.CreateFormulaEvaluator();
                CellValue value = eval.Evaluate(cell);

                ClassicAssert.IsFalse(value.BooleanValue);
            }
            finally
            {
                wb.Close();
            }
        }

        [Test]
        public void Bug62275_empty_2args()
        {
            IWorkbook wb = new XSSFWorkbook();
            try
            {
                ISheet sheet = wb.CreateSheet();
                IRow row = sheet.CreateRow(0);

                ICell cell = row.CreateCell(0);
                cell.SetCellFormula("vlookup(A2,B1:B5,,)");

                ICreationHelper crateHelper = wb.GetCreationHelper();
                IFormulaEvaluator eval = crateHelper.CreateFormulaEvaluator();
                CellValue value = eval.Evaluate(cell);

                ClassicAssert.IsFalse(value.BooleanValue);
            }
            finally
            {
                wb.Close();
            }
        }

        [Test]
        public void Bug62275_empty_1arg()
        {
            IWorkbook wb = new XSSFWorkbook();
            try
            {
                ISheet sheet = wb.CreateSheet();
                IRow row = sheet.CreateRow(0);

                ICell cell = row.CreateCell(0);
                cell.SetCellFormula("vlookup(A2,,,)");

                ICreationHelper crateHelper = wb.GetCreationHelper();
                IFormulaEvaluator eval = crateHelper.CreateFormulaEvaluator();
                CellValue value = eval.Evaluate(cell);

                ClassicAssert.IsFalse(value.BooleanValue);
            }
            finally
            {
                wb.Close();
            }
        }
    }
}

