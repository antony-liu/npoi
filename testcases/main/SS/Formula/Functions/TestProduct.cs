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
    [TestFixture]
    public class TestProduct
    {
        [Test]
        public void MissingArgsAreIgnored()
        {
            ValueEval result = GetInstance().Evaluate(new ValueEval[]{new NumberEval(2.0), MissingArgEval.instance}, 0, 0);
            ClassicAssert.IsTrue(result is NumberEval);
            ClassicAssert.AreEqual(2, ((NumberEval) result).NumberValue, 0);
        }

        /// <summary>
        /// Actually PRODUCT() requires at least one arg but the checks are performed elsewhere.
        /// However, PRODUCT(,) is a valid call (which should return 0). So it makes sense to
        /// assert that PRODUCT() is also 0 (at least, nothing explodes).
        /// </summary>
        public void MissingArgEvalReturns0()
        {
            ValueEval result = GetInstance().Evaluate(new ValueEval[0], 0, 0);
            ClassicAssert.IsTrue(result is NumberEval);
            ClassicAssert.AreEqual(0, ((NumberEval) result).NumberValue, 0);
        }

        [Test]
        public void TwoMissingArgEvalsReturn0()
        {
            ValueEval result = GetInstance().Evaluate(new ValueEval[]{MissingArgEval.instance, MissingArgEval.instance}, 0, 0);
            ClassicAssert.IsTrue(result is NumberEval);
            ClassicAssert.AreEqual(0, ((NumberEval) result).NumberValue, 0);
        }

        [Test]
        public void AcceptanceTest()
        {
            ValueEval[] args = {
                new NumberEval(2.0),
                MissingArgEval.instance,
                new StringEval("6"),
                BoolEval.TRUE};
            ValueEval result = GetInstance().Evaluate(args, 0, 0);
            ClassicAssert.IsTrue(result is NumberEval);
            ClassicAssert.AreEqual(12, ((NumberEval) result).NumberValue, 0);
        }

        private static Function GetInstance()
        {
            return AggregateFunction.PRODUCT;
        }
    }
}


