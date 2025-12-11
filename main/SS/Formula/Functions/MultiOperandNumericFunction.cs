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

namespace NPOI.SS.Formula.Functions
{
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Eval;
    using NPOI.Util;
    using System;

    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     * This Is the base class for all excel function evaluator
     * classes that take variable number of operands, and
     * where the order of operands does not matter
     */
    public abstract class MultiOperandNumericFunction : Function
    {
        public enum Policy { COERCE, SKIP, ERROR }
        //private interface IEvalConsumer<TValue, TReceiver>
        //{
        //    void Accept(TValue value, TReceiver receiver);
        //}

        private Action<BoolEval, DoubleList> boolByRefConsumer;
        private Action<BoolEval, DoubleList> boolByValueConsumer;
        private Action<BlankEval, DoubleList> blankConsumer;
        private Action<MissingArgEval, DoubleList> missingArgConsumer = ConsumerFactory.CreateForMissingArg(Policy.SKIP);

        static readonly double[] EMPTY_DOUBLE_ARRAY = [];

        protected MultiOperandNumericFunction(bool isReferenceBoolCounted, bool isBlankCounted)
        {
            boolByRefConsumer = ConsumerFactory.CreateForBoolEval(isReferenceBoolCounted ? Policy.COERCE : Policy.SKIP);
            boolByValueConsumer = ConsumerFactory.CreateForBoolEval(Policy.COERCE);
            blankConsumer = ConsumerFactory.CreateForBlank(isBlankCounted ? Policy.COERCE : Policy.SKIP);
        }
        protected internal abstract double Evaluate(double[] values);

        public ValueEval Evaluate(ValueEval[] args, int srcCellRow, int srcCellCol)
        {
            try
            {
                double[] values = GetNumberArray(args);
                double d = Evaluate(values);
                if(Double.IsNaN(d) || Double.IsInfinity(d))
                    return ErrorEval.NUM_ERROR;

                return new NumberEval(d);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }

        private sealed class DoubleList
        {
            private double[] _array;
            private int _Count;

            public DoubleList()
            {
                _array = new double[8];
                _Count = 0;
            }

            public double[] ToArray()
            {
                if (_Count < 1)
                {
                    return EMPTY_DOUBLE_ARRAY;
                }
                double[] result = new double[_Count];
                Array.Copy(_array, 0, result, 0, _Count);
                return result;
            }

            public void Add(double[] values)
            {
                int AddLen = values.Length;
                EnsureCapacity(_Count + AddLen);
                Array.Copy(values, 0, _array, _Count, AddLen);
                _Count += AddLen;
            }

            private void EnsureCapacity(int reqSize)
            {
                if (reqSize > _array.Length)
                {
                    int newSize = reqSize * 3 / 2; // grow with 50% extra
                    double[] newArr = new double[newSize];
                    Array.Copy(_array, 0, newArr, 0, _Count);
                    _array = newArr;
                }
            }

            public void Add(double value)
            {
                EnsureCapacity(_Count + 1);
                _array[_Count] = value;
                _Count++;
            }
        }

        private static readonly int DEFAULT_MAX_NUM_OPERANDS = SpreadsheetVersion.EXCEL2007.MaxFunctionArgs;

        public void SetMissingArgPolicy(Policy policy)
        {
            missingArgConsumer = ConsumerFactory.CreateForMissingArg(policy);
        }

        public void SetBlankEvalPolicy(Policy policy)
        {
            blankConsumer = ConsumerFactory.CreateForBlank(policy);
        }
        /**
         * Maximum number of operands accepted by this function.
         * Subclasses may override to Change default value.
         */
        internal virtual int MaxNumOperands
        {
            get
            {
                return DEFAULT_MAX_NUM_OPERANDS;
            }
        }
        /**
     *  Whether to count nested subtotals.
     */
        public virtual bool IsSubtotalCounted
        {
            get
            {
                return true;
            }
        }
        /**
     * Collects values from a single argument
     */
        private void CollectValues(ValueEval operand, DoubleList temp)
        {
            if (operand is ThreeDEval eval)
            {
                for (int sIx = eval.FirstSheetIndex; sIx <= eval.LastSheetIndex; sIx++)
                {
                    int width = eval.Width;
                    int height = eval.Height;
                    for (int rrIx = 0; rrIx < height; rrIx++)
                    {
                        for (int rcIx = 0; rcIx < width; rcIx++)
                        {
                            ValueEval ve = eval.GetValue(sIx, rrIx, rcIx);
                            if (!IsSubtotalCounted && eval.IsSubTotal(rrIx, rcIx)) continue;
                            CollectValue(ve, true, temp);
                        }
                    }
                }
                return;
            }
            if (operand is TwoDEval ae)
            {
                int width = ae.Width;
                int height = ae.Height;
                for (int rrIx = 0; rrIx < height; rrIx++)
                {
                    for (int rcIx = 0; rcIx < width; rcIx++)
                    {
                        ValueEval ve = ae.GetValue(rrIx, rcIx);
                        if (!IsSubtotalCounted && ae.IsSubTotal(rrIx, rcIx)) continue;
                        CollectValue(ve, true, temp);
                    }
                }
                return;
            }
            if (operand is RefEval re)
            {
                for (int sIx = re.FirstSheetIndex; sIx <= re.LastSheetIndex; sIx++)
                {
                    CollectValue(re.GetInnerValueEval(sIx), true, temp);
                }
                return;
            }
            CollectValue((ValueEval)operand, false, temp);
        }
        private void CollectValue(ValueEval ve, bool isViaReference, DoubleList temp)
        {
            if (ve == null)
            {
                throw new ArgumentException("ve must not be null");
            }
            if (ve is BoolEval boolEval)
            {
                if(isViaReference)
                {
                    boolByRefConsumer.Invoke(boolEval, temp);
                }
                else
                {
                    boolByValueConsumer.Invoke(boolEval, temp);
                }
                return;
            }
            if (ve is NumberEval ne)
            {
                temp.Add(ne.NumberValue);
                return;
            }
            if (ve is StringEval eval)
            {
                if (isViaReference)
                {
                    // ignore all ref strings
                    return;
                }
                String s = eval.StringValue;
                Double d = OperandResolver.ParseDouble(s);
                if (double.IsNaN(d))
                {
                    throw new EvaluationException(ErrorEval.VALUE_INVALID);
                }
                temp.Add(d);
                return;
            }
            if (ve is ErrorEval errorEval)
            {
                throw new EvaluationException(errorEval);
            }
            if (ve == BlankEval.instance)
            {
                blankConsumer.Invoke((BlankEval) ve, temp);
                return;
            }
            if(ve == MissingArgEval.instance)
            {
                missingArgConsumer.Invoke((MissingArgEval) ve, temp);
                return;
            }
            throw new InvalidOperationException("Invalid ValueEval type passed for conversion: ("
                    + ve.GetType().Name + ")");
        }
        /**
         * Returns a double array that contains values for the numeric cells
         * from among the list of operands. Blanks and Blank equivalent cells
         * are ignored. Error operands or cells containing operands of type
         * that are considered invalid and would result in #VALUE! error in
         * excel cause this function to return <c>null</c>.
         *
         * @return never <c>null</c>
         */
        protected double[] GetNumberArray(ValueEval[] operands)
        {
            if (operands.Length > MaxNumOperands)
            {
                throw EvaluationException.InvalidValue();
            }
            DoubleList retval = new DoubleList();

            for (int i = 0, iSize = operands.Length; i < iSize; i++)
            {
                CollectValues(operands[i], retval);
            }
            return retval.ToArray();
        }
        /**
         * Ensures that a two dimensional array has all sub-arrays present and the same Length
         * @return <c>false</c> if any sub-array Is missing, or Is of different Length
         */
        protected static bool AreSubArraysConsistent(double[][] values)
        {

            if (values == null || values.Length < 1)
            {
                // TODO this doesn't seem right.  Fix or Add comment.
                return true;
            }

            if (values[0] == null)
            {
                return false;
            }
            int outerMax = values.Length;
            int innerMax = values[0].Length;
            for (int i = 1; i < outerMax; i++)
            { // note - 'i=1' start at second sub-array
                double[] subArr = values[i];
                if (subArr == null)
                {
                    return false;
                }
                if (innerMax != subArr.Length)
                {
                    return false;
                }
            }
            return true;
        }

        private static class ConsumerFactory
        {
            public static Action<MissingArgEval, DoubleList> CreateForMissingArg(Policy policy)
            {
                Action<MissingArgEval, DoubleList> coercer =
                        (MissingArgEval value, DoubleList receiver) => receiver.Add(0.0);
                return CreateAny(coercer, policy);
            }

            public static Action<BoolEval, DoubleList> CreateForBoolEval(Policy policy)
            {
                Action<BoolEval, DoubleList> coercer =
                        (BoolEval value, DoubleList receiver) => receiver.Add(value.NumberValue);
                return CreateAny(coercer, policy);
            }

            public static Action<BlankEval, DoubleList> CreateForBlank(Policy policy)
            {
                Action<BlankEval, DoubleList> coercer =
                        (BlankEval value, DoubleList receiver) => receiver.Add(0.0);
                return CreateAny(coercer, policy);
            }

            private static Action<T, DoubleList> CreateAny<T>(Action<T, DoubleList> coercer, Policy policy)
                where T : ValueEval
            {
                switch(policy)
                {
                    case Policy.COERCE:
                        return coercer;
                    case Policy.SKIP:
                        return doNothing<T>();
                    case Policy.ERROR:
                        return throwValueInvalid<T>();
                    default:
                        throw new AssertFailedException("");
                }
            }

            private static Action<T, DoubleList> doNothing<T>() where T : ValueEval
            {
                return (T value, DoubleList receiver) => {
                };
            }

            private static Action<T, DoubleList> throwValueInvalid<T>() where T : ValueEval
            {
                return (T value, DoubleList receiver) => {
                    throw new EvaluationException(ErrorEval.VALUE_INVALID);
                };
            }
        }
    }
}