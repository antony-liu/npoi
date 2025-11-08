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

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NUnit.Framework;using NUnit.Framework.Legacy;
using System;
using System.Collections.Generic;
using System.Text;

namespace TestCases.SS.Formula.Functions
{
    [TestFixture]
    public class TestAreas
    {
        [Test]
        public void TestBasic()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ICell cell = wb.CreateSheet().CreateRow(0).CreateCell(0);
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);

            String formulaText = "AREAS(B1)";
            confirmResult(fe, cell, formulaText, 1.0);

            formulaText = "AREAS(B2:D4)";
            confirmResult(fe, cell, formulaText, 1.0);

            formulaText = "AREAS((B2:D4,E5,F6:I9))";
            confirmResult(fe, cell, formulaText, 3.0);

            formulaText = "AREAS((B2:D4,E5,C3,E4))";
            confirmResult(fe, cell, formulaText, 4.0);

            formulaText = "AREAS((I9))";
            confirmResult(fe, cell, formulaText, 1.0);
        }

        private static void confirmResult(HSSFFormulaEvaluator fe, ICell cell, String formulaText, Double expectedResult)
        {
            cell.SetCellFormula(formulaText);
            fe.NotifyUpdateCell(cell);
            CellValue result = fe.Evaluate(cell);
            ClassicAssert.AreEqual(CellType.Numeric,result.CellType);
            ClassicAssert.AreEqual(expectedResult, result.NumberValue);
        }
    }
}
