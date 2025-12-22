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
 * Created on May 9, 2005
 *
 */
namespace NPOI.SS.Formula.Functions
{
    using NPOI.SS.Formula.Eval;
    using System;


    /*
     * @author Amol S. Deshmukh &lt; amol at apache dot org &gt;
     * The NOT bool function. Returns negation of specified value
     * (treated as a bool). If the specified arg Is a number,
     * then it Is true <=> 'number Is non-zero'
     */
    public class Not : Boolean1ArgFunction
    {
        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0)
        {
            bool boolArgVal;
            try
            {
                ValueEval ve = OperandResolver.GetSingleValue(arg0, srcRowIndex, srcColumnIndex);
                bool? b = OperandResolver.CoerceValueToBoolean(ve, false);
                boolArgVal = b??false;
            }
            catch(EvaluationException e)
            {
                return e.GetErrorEval();
            }

            return BoolEval.ValueOf(!boolArgVal);
        }
    }

    public abstract class Boolean1ArgFunction : Fixed1ArgFunction, IArrayFunction
    {
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
        public ValueEval EvaluateArray(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            if(args.Length != 1)
            {
                return ErrorEval.VALUE_INVALID;
            }
            return EvaluateOneArrayArg(args[0], srcRowIndex, srcColumnIndex,
                    (vA) => Evaluate(srcRowIndex, srcColumnIndex, vA));
        }
    }
}