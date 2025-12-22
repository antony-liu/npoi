/*
* Licensed to the Apache Software Foundation (ASF) Under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for Additional information regarding copyright ownership.
* The ASF licenses this file to You Under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed Under the License is distributed on an "AS Is" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations Under the License.
*/
/*
 * Created on May 10, 2005
 *
 */
namespace NPOI.SS.Formula.Eval
{
    using System;
    using NPOI.SS.Formula.Functions;
    using NPOI.SS.Util;

    //public class RelationalValues
    //{
    //    public double[] ds = new Double[2];
    //    public bool[] bs = new bool[2];
    //    public string[] ss = new String[3];
    //    public ErrorEval ee = null;
    //}
    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo Dot com &gt;
     *
     */
    public abstract class RelationalOperationEval : Fixed2ArgFunction, IArrayFunction
    {
        private static int DoCompare(ValueEval va, ValueEval vb)
        {
            // special cases when one operand is blank or missing
            if (va == BlankEval.instance|| va is MissingArgEval)
            {
                return CompareBlank(vb);
            }
            if (vb == BlankEval.instance || vb is MissingArgEval)
            {
                return -CompareBlank(va);
            }

            if (va is BoolEval bA)
            {
                if (vb is BoolEval bB)
                {
                    if (bA.BooleanValue == bB.BooleanValue)
                    {
                        return 0;
                    }
                    return bA.BooleanValue ? 1 : -1;
                }
                return 1;
            }
            if (vb is BoolEval)
            {
                return -1;
            }
            if (va is StringEval sA)
            {
                if (vb is StringEval sB)
                {
                    return string.Compare(sA.StringValue, sB.StringValue, StringComparison.OrdinalIgnoreCase);
                }
                return 1;
            }
            if (vb is StringEval)
            {
                return -1;
            }
            if (va is NumberEval nA)
            {
                if (vb is NumberEval nB)
                {
                    if (nA.NumberValue == nB.NumberValue)
                    {
                        // Excel considers -0.0 == 0.0 which is different to Double.compare()
                        return 0;
                    }
                    return NumberComparer.Compare(nA.NumberValue, nB.NumberValue);
                }
            }
            throw new ArgumentException("Bad operand types (" + va.GetType().Name + "), ("
                    + vb.GetType().Name + ")");
        }
        private static int CompareBlank(ValueEval v)
        {
            if (v == BlankEval.instance|| v is MissingArgEval)
            {
                return 0;
            }
            if (v is BoolEval boolEval)
            {
                return boolEval.BooleanValue ? -1 : 0;
            }
            if (v is NumberEval ne)
            {
                //return ne.NumberValue.CompareTo(0.0);
                return NumberComparer.Compare(0.0, ne.NumberValue);
            }
            if (v is StringEval se)
            {
                return se.StringValue.Length < 1 ? 0 : -1;
            }
            throw new ArgumentException("bad value class (" + v.GetType().Name + ")");
        }

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1)
        {

            ValueEval vA;
            ValueEval vB;
            try
            {
                vA = OperandResolver.GetSingleValue(arg0, srcRowIndex, srcColumnIndex);
                vB = OperandResolver.GetSingleValue(arg1, srcRowIndex, srcColumnIndex);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            int cmpResult = DoCompare(vA, vB);
            bool result = ConvertComparisonResult(cmpResult);
            return BoolEval.ValueOf(result);
        }
        public ValueEval EvaluateArray(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            ValueEval arg0 = args[0];
            ValueEval arg1 = args[1];

            return EvaluateTwoArrayArgs(arg0, arg1, srcRowIndex, srcColumnIndex, (vA, vB) => {
                int cmpResult = DoCompare(vA, vB);
                bool result = ConvertComparisonResult(cmpResult);
                return BoolEval.ValueOf(result);
            });
        }

        public ValueEval EvaluateTwoArrayArgs(ValueEval arg0, ValueEval arg1, int srcRowIndex, int srcColumnIndex,
                                           Func<ValueEval, ValueEval, ValueEval> evalFunc)
        {
            return new ArrayFunction().EvaluateTwoArrayArgs(arg0, arg1, srcRowIndex, srcColumnIndex, evalFunc);
        }
        public ValueEval EvaluateOneArrayArg(ValueEval arg0, int srcRowIndex, int srcColumnIndex,
                                          Func<ValueEval, ValueEval> evalFunc)
        {
            return new ArrayFunction().EvaluateOneArrayArg(arg0, srcRowIndex, srcColumnIndex, evalFunc);
        }

        public abstract bool ConvertComparisonResult(int cmpResult);


        public static Function EqualEval = new EqualEval();
        public static Function NotEqualEval = new NotEqualEval();
        public static Function LessEqualEval = new LessEqualEval();
        public static Function LessThanEval = new LessThanEval();
        public static Function GreaterEqualEval = new GreaterEqualEval();
        public static Function GreaterThanEval = new GreaterThanEval();
    }
}