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

namespace NPOI.SS.Formula.Eval
{
    using NPOI.SS.Formula.Functions;
    using System;

    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     *  
     */
    public class UnaryPlusEval : Fixed1ArgFunction, IArrayFunction
    {

        public static Function instance = new UnaryPlusEval();
    	
	    private UnaryPlusEval() {
	    }

        public override ValueEval Evaluate(int srcCellRow, int srcCellCol, ValueEval arg0)
        {
            double d;
            try
            {
                ValueEval ve = OperandResolver.GetSingleValue(arg0, srcCellRow, srcCellCol);
                if (ve is BlankEval)
                {
                    return NumberEval.ZERO;
                }
                if (ve is StringEval)
                {
                    // Note - asymmetric with UnaryMinus
                    // -"hello" Evaluates to #VALUE!
                    // but +"hello" Evaluates to "hello"
                    return ve;
                }
                d = OperandResolver.CoerceValueToDouble(ve);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            return new NumberEval(+d);
        }

        public ValueEval EvaluateArray(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            if(args.Length != 1)
            {
                return ErrorEval.VALUE_INVALID;
            }
            return EvaluateOneArrayArg(args[0], srcRowIndex, srcColumnIndex, (valA)=>
                    Evaluate(srcRowIndex, srcColumnIndex, valA)
            );
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
    }
}