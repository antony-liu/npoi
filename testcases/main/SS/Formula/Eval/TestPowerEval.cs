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

namespace TestCases.SS.Formula.Eval
{
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.Functions;
    using NPOI.SS.UserModel;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;

    /// <summary>
    /// Tests for power operator evaluator.
    /// </summary>
    /// <remarks>
    /// @author Bob van den Berge
    /// </remarks>
    [TestFixture]
    public sealed class TestPowerEval
    {
        [Test]
        public void TestPositiveValues()
        {
            confirm(0, 0, 1);
            confirm(1, 1, 0);
            confirm(9, 3, 2);
        }
        [Test]
        public void TestNegativeValues()
        {
            confirm(-1, -1, 1);
            confirm(1, 1, -1);
            confirm(1, -10, 0);
            confirm((1.0/3), 3, -1);
        }
        [Test]
        public void TestPositiveDecimalValues()
        {
            confirm(3, 27, (1/3.0));
        }
        [Test]
        public void TestNegativeDecimalValues()
        {
            confirm(-3, -27, (1/3.0));
        }
        [Test]
        public void TestErrorValues()
        {
            confirmError(-1.00001, 1.1);
        }
        [Test]
        public void TestInSpreadSheet()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.CreateSheet("Sheet1") as HSSFSheet;
            HSSFRow row = sheet.CreateRow(0) as HSSFRow;
            HSSFCell cell = row.CreateCell(0) as HSSFCell;
            cell.SetCellFormula("B1^C1");
            row.CreateCell(1).SetCellValue(-27);
            row.CreateCell(2).SetCellValue((1/3.0));

            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            CellValue cv = fe.Evaluate(cell);

            ClassicAssert.AreEqual(CellType.Numeric, cv.CellType);
            ClassicAssert.AreEqual(-3.0, cv.NumberValue);
        }

        private void confirm(double expected, double a, double b)
        {
            NumberEval result = (NumberEval) evaluate(EvalInstances.Power, a, b);

            ClassicAssert.AreEqual(expected, result.NumberValue);
        }

        private void confirmError(double a, double b)
        {
            ErrorEval result = (ErrorEval) evaluate(EvalInstances.Power, a, b);

            ClassicAssert.AreEqual("#NUM!", result.ErrorString);
        }

        private static ValueEval evaluate(Function instance, params double[] dArgs)
        {
            ValueEval[] evalArgs;
            evalArgs = new ValueEval[dArgs.Length];
            for(int i = 0; i < evalArgs.Length; i++)
            {
                evalArgs[i] = new NumberEval(dArgs[i]);
            }
            return instance.Evaluate(evalArgs, -1, (short) -1);
        }
    }
}


