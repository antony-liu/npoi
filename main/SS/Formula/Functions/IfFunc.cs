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
 * Created on Nov 25, 2006
 *
 */
namespace NPOI.SS.Formula.Functions
{
    using NPOI.SS.Formula.Eval;
    using System;

    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     * 
     */
    public class IfFunc : Var2or3ArgFunction, IArrayFunction
    {

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1)
        {
            bool b;
            try
            {
                b = EvaluateFirstArg(arg0, srcRowIndex, srcColumnIndex);
            }
            catch(EvaluationException e)
            {
                return e.GetErrorEval();
            }
            if(b)
            {
                if(arg1 == MissingArgEval.instance)
                {
                    return BlankEval.instance;
                }
                return arg1;
            }
            return BoolEval.FALSE;
        }

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1,
                ValueEval arg2)
        {
            bool b;
            try
            {
                b = EvaluateFirstArg(arg0, srcRowIndex, srcColumnIndex);
            }
            catch(EvaluationException e)
            {
                return e.GetErrorEval();
            }
            if(b)
            {
                if(arg1 == MissingArgEval.instance)
                {
                    return BlankEval.instance;
                }
                return arg1;
            }
            if(arg2 == MissingArgEval.instance)
            {
                return BlankEval.instance;
            }
            return arg2;
        }

        public static bool EvaluateFirstArg(ValueEval arg, int srcCellRow, int srcCellCol)
        {
            ValueEval ve = OperandResolver.GetSingleValue(arg, srcCellRow, srcCellCol);
            bool? b = OperandResolver.CoerceValueToBoolean(ve, false);
            if(b == null)
            {
                return false;
            }
            return (bool) b;
        }

        public ValueEval EvaluateTwoArrayArgs(ValueEval arg0, ValueEval arg1, int srcRowIndex, int srcColumnIndex,
                                   Func<ValueEval, ValueEval, ValueEval> evalFunc)
        {
            return new ArrayFunction().EvaluateTwoArrayArgs(arg0, arg1, srcRowIndex, srcColumnIndex, evalFunc);
        }
        public ValueEval EvaluateOneArrayArg(ValueEval[] args, int srcRowIndex, int srcColumnIndex,
                                          Func<ValueEval, ValueEval> evalFunc)
        {
            return new ArrayFunction().EvaluateOneArrayArg(args[0], srcRowIndex, srcColumnIndex, evalFunc);
        }

        public ValueEval EvaluateArray(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            if(args.Length < 2 || args.Length > 3)
            {
                return ErrorEval.VALUE_INVALID;
            }

            ValueEval arg0 = args[0];
            ValueEval arg1 = args[1];
            ValueEval arg2 = args.Length == 2 ? BoolEval.FALSE : args[2];
            return EvaluateArrayArgs(arg0, arg1, arg2, srcRowIndex, srcColumnIndex);
        }

        private static ValueEval EvaluateArrayArgs(ValueEval arg0, ValueEval arg1, ValueEval arg2, int srcRowIndex, int srcColumnIndex)
        {
            int w1, w2, h1, h2;
            int a1FirstCol = 0, a1FirstRow = 0;
            if(arg0 is AreaEval)
            {
                AreaEval ae = (AreaEval)arg0;
                w1 = ae.Width;
                h1 = ae.Height;
                a1FirstCol = ae.FirstColumn;
                a1FirstRow = ae.FirstRow;
            }
            else if(arg0 is RefEval)
            {
                RefEval ref1 = (RefEval)arg0;
                w1 = 1;
                h1 = 1;
                a1FirstCol = ref1.Column;
                a1FirstRow = ref1.Row;
            }
            else
            {
                w1 = 1;
                h1 = 1;
            }
            int a2FirstCol = 0, a2FirstRow = 0;
            if(arg1 is AreaEval)
            {
                AreaEval ae = (AreaEval)arg1;
                w2 = ae.Width;
                h2 = ae.Height;
                a2FirstCol = ae.FirstColumn;
                a2FirstRow = ae.FirstRow;
            }
            else if(arg1 is RefEval)
            {
                RefEval ref1 = (RefEval)arg1;
                w2 = 1;
                h2 = 1;
                a2FirstCol = ref1.Column;
                a2FirstRow = ref1.Row;
            }
            else
            {
                w2 = 1;
                h2 = 1;
            }

            int a3FirstCol = 0, a3FirstRow = 0;
            if(arg2 is AreaEval)
            {
                AreaEval ae = (AreaEval)arg2;
                a3FirstCol = ae.FirstColumn;
                a3FirstRow = ae.FirstRow;
            }
            else if(arg2 is RefEval)
            {
                RefEval ref1 = (RefEval)arg2;
                a3FirstCol = ref1.Column;
                a3FirstRow = ref1.Row;
            }

            int width = Math.Max(w1, w2);
            int height = Math.Max(h1, h2);

            ValueEval[] vals = new ValueEval[height * width];

            int idx = 0;
            for(int i = 0; i < height; i++)
            {
                for(int j = 0; j < width; j++)
                {
                    ValueEval vA;
                    try
                    {
                        vA = OperandResolver.GetSingleValue(arg0, a1FirstRow + i, a1FirstCol + j);
                    }
                    catch(FormulaParseException e)
                    {
                        vA = ErrorEval.NAME_INVALID;
                    }
                    catch(EvaluationException e)
                    {
                        vA = e.GetErrorEval();
                    }
                    ValueEval vB;
                    try
                    {
                        vB = OperandResolver.GetSingleValue(arg1, a2FirstRow + i, a2FirstCol + j);
                    }
                    catch(FormulaParseException e)
                    {
                        vB = ErrorEval.NAME_INVALID;
                    }
                    catch(EvaluationException e)
                    {
                        vB = e.GetErrorEval();
                    }

                    ValueEval vC;
                    try
                    {
                        vC = OperandResolver.GetSingleValue(arg2, a3FirstRow + i, a3FirstCol + j);
                    }
                    catch(FormulaParseException e)
                    {
                        vC = ErrorEval.NAME_INVALID;
                    }
                    catch(EvaluationException e)
                    {
                        vC = e.GetErrorEval();
                    }

                    if(vA is ErrorEval)
                    {
                        vals[idx++] = vA;
                    }
                    else if(vB is ErrorEval)
                    {
                        vals[idx++] = vB;
                    }
                    else
                    {
                        bool? b;
                        try
                        {
                            b = OperandResolver.CoerceValueToBoolean(vA, false);
                            vals[idx++] = b.HasValue && b.Value ? vB : vC;
                        }
                        catch(EvaluationException e)
                        {
                            vals[idx++] = e.GetErrorEval();
                        }
                    }

                }
            }

            if(vals.Length == 1)
            {
                return vals[0];
            }

            return new CacheAreaEval(srcRowIndex, srcColumnIndex, srcRowIndex + height - 1, srcColumnIndex + width - 1, vals);
        }
    }
}