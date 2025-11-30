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
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.Functions;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;


    /// <summary>
    /// <para>
    /// From Excel documentation at https://support.office.com/en-us/article/geomean-function-db1ac48d-25a5-40a0-ab83-0b38980e40d5:
    /// 1. Arguments can either be numbers or names, arrays, or references that contain numbers.
    /// 2. Logical values and text representations of numbers that you type directly into the list of arguments are counted.
    /// 3. If an array or reference argument contains text, logical values, or empty cells, those values are ignored; however, cells with the value zero are included.
    /// 4. Arguments that are error values or text that cannot be translated into numbers cause errors.
    /// 5. If any data point ≤ 0, GEOMEAN returns the #NUM! error value.
    /// </para>
    /// <para>
    /// Remarks:
    /// Actually, 5. is not true. If an error is encountered before a 0 value, the error is returned.
    /// </para>
    /// </summary>
    [TestFixture]
    public class TestGeomean
    {
        [Test]
        public void AcceptanceTest()
        {
            Function geomean = GetInstance();

            ValueEval result = geomean.Evaluate(new ValueEval[]{new NumberEval(2), new NumberEval(3)}, 0, 0);
            VerifyNumericResult(2.449489742783178, result);
        }

        [Test]
        public void BoolsByValueAreCoerced()
        {
            ValueEval[] args = {BoolEval.TRUE};
            ValueEval result = GetInstance().Evaluate(args, 0, 0);
            VerifyNumericResult(1.0, result);
        }

        [Test]
        public void StringsByValueAreCoerced()
        {
            ValueEval[] args = {new StringEval("2")};
            ValueEval result = GetInstance().Evaluate(args, 0, 0);
            VerifyNumericResult(2.0, result);
        }

        [Test]
        public void NonCoerceableStringsByValueCauseValueInvalid()
        {
            ValueEval[] args = {new StringEval("foo")};
            ValueEval result = GetInstance().Evaluate(args, 0, 0);
            ClassicAssert.AreEqual(ErrorEval.VALUE_INVALID, result);
        }

        [Test]
        public void BoolsByReferenceAreSkipped()
        {
            ValueEval[] args = new ValueEval[]{new NumberEval(2.0), EvalFactory.CreateRefEval("A1", BoolEval.TRUE)};
            ValueEval result = GetInstance().Evaluate(args, 0, 0);
            VerifyNumericResult(2.0, result);
        }

        [Test]
        public void BoolsStringsAndBlanksByReferenceAreSkipped()
        {
            ValueEval ref1 = EvalFactory.CreateAreaEval("A1:A3", new ValueEval[]{new StringEval("foo"), BoolEval.FALSE, BlankEval.instance});
            ValueEval[] args = {ref1, new NumberEval(2.0)};
            ValueEval result = GetInstance().Evaluate(args, 0, 0);
            VerifyNumericResult(2.0, result);
        }

        [Test]
        public void StringsByValueAreCounted()
        {
            ValueEval[] args = {new StringEval("2.0")};
            ValueEval result = GetInstance().Evaluate(args, 0, 0);
            VerifyNumericResult(2.0, result);
        }

        [Test]
        public void MissingArgCountAsZero()
        {
            // and, naturally, produces a NUM_ERROR
            ValueEval[] args = {new NumberEval(1.0), MissingArgEval.instance};
            ValueEval result = GetInstance().Evaluate(args, 0, 0);
            ClassicAssert.AreEqual(ErrorEval.NUM_ERROR, result);
        }

        /// <summary>
        /// Implementation-specific: the math lib returns 0 for the input [1.0, 0.0], but a NUM_ERROR should be returned.
        /// </summary>
        [Test]
        public void Sequence_1_0_shouldReturnError()
        {
            ValueEval[] args = {new NumberEval(1.0), new NumberEval(0)};
            ValueEval result = GetInstance().Evaluate(args, 0, 0);
            ClassicAssert.AreEqual(ErrorEval.NUM_ERROR, result);
        }

        [Test]
        public void MinusOneShouldReturnError()
        {
            ValueEval[] args = {new NumberEval(1.0), new NumberEval(-1.0)};
            ValueEval result = GetInstance().Evaluate(args, 0, 0);
            ClassicAssert.AreEqual(ErrorEval.NUM_ERROR, result);
        }

        [Test]
        public void FirstErrorPropagates()
        {
            ValueEval[] args = {ErrorEval.DIV_ZERO, ErrorEval.NUM_ERROR};
            ValueEval result = GetInstance().Evaluate(args, 0, 0);
            ClassicAssert.AreEqual(ErrorEval.DIV_ZERO, result);
        }

        private void VerifyNumericResult(double expected, ValueEval result)
        {
            ClassicAssert.IsTrue(result is NumberEval);
            ClassicAssert.AreEqual(expected, ((NumberEval) result).NumberValue, 1e-15);
        }

        private Function GetInstance()
        {
            return AggregateFunction.GEOMEAN;
        }
    }
}
