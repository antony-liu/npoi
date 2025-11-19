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

/*
 * Notes:
 * Duplicate x values don't work most of the time because of the way the
 * math library handles multiple regression.
 * The math library currently Assert.Fails when the number of x variables is >=
 * the sample size (see https://github.com/Hipparchus-Math/hipparchus/issues/13).
 */


using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.SS.Formula.Functions
{
    using MathNet.Numerics.LinearRegression;
    using NPOI.SS.Formula;

    using NPOI.SS.Formula.Eval;
    using NPOI.Util;
    using System.Linq;

    /// <summary>
    /// <para>
    /// Implementation for the Excel function TREND
    /// </para>
    /// <para>
    /// Syntax:
    /// </para>
    /// <para>
    /// TREND(known_y's, known_x's, new_x's, constant)
    ///    <list type="table">
    /// <listheader><term>known_y's, known_x's, new_x's</term><description>constant</description></listheader>
    /// <item><term>known_y's, known_x's, new_x's</term><description>typically area references, possibly cell references or scalar values</description></item>
    /// <item><term>constant</term><description>TRUE or FALSE:
    ///      determines whether the regression line should include an intercept term</description></item>
    /// </para>
    /// <para>
    /// If <b>known_x's</b> is not given, it is assumed to be the default array {1, 2, 3, ...}
    /// of the same size as <b>known_y's</b>.
    /// If <b>new_x's</b> is not given, it is assumed to be the same as <b>known_x's</b>
    /// If <b>constant</b> is omitted, it is assumed to be <b>TRUE</b>
    /// </para>
    /// </summary>
    public sealed class Trend : Function
    {
        MatrixFunction.MutableValueCollector collector = new MatrixFunction.MutableValueCollector(false, false);
        private sealed class TrendResults
        {
            public double[] vals;
            public int resultWidth;
            public int resultHeight;

            public TrendResults(double[] vals, int resultWidth, int resultHeight)
            {
                this.vals = vals;
                this.resultWidth = resultWidth;
                this.resultHeight = resultHeight;
            }
        }

        public ValueEval Evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            if(args.Length < 1 || args.Length > 4)
            {
                return ErrorEval.VALUE_INVALID;
            }
            try
            {
                TrendResults tr = GetNewY(args);
                ValueEval[] vals = new ValueEval[tr.vals.Length];
                for(int i = 0; i < tr.vals.Length; i++)
                {
                    vals[i] = new NumberEval(tr.vals[i]);
                }
                if(tr.vals.Length == 1)
                {
                    return vals[0];
                }
                return new CacheAreaEval(srcRowIndex, srcColumnIndex, srcRowIndex + tr.resultHeight - 1, srcColumnIndex + tr.resultWidth - 1, vals);
            }
            catch(EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }

        private static double[][] EvalToArray(ValueEval arg)
        {
            double[][] ar;
            ValueEval eval;
            if(arg is MissingArgEval)
            {
                return [];
            }
            if(arg is RefEval)
            {
                RefEval re = (RefEval) arg;
                if(re.NumberOfSheets > 1)
                {
                    throw new EvaluationException(ErrorEval.VALUE_INVALID);
                }
                eval = re.GetInnerValueEval(re.FirstSheetIndex);
            }
            else
            {
                eval = arg;
            }
            if(eval == null)
            {
                throw new RuntimeException("Parameter may not be null.");
            }

            if(eval is AreaEval)
            {
                AreaEval ae = (AreaEval) eval;
                int w = ae.Width;
                int h = ae.Height;
                ar = Arrays.Allacate<double>(h, w);
                for(int i = 0; i < h; i++)
                {
                    for(int j = 0; j < w; j++)
                    {
                        ValueEval ve = ae.GetRelativeValue(i, j);
                        if(!(ve is NumericValueEval))
                        {
                            throw new EvaluationException(ErrorEval.VALUE_INVALID);
                        }
                        ar[i][j] = ((NumericValueEval) ve).NumberValue;
                    }
                }
            }
            else if(eval is NumericValueEval)
            {
                ar = Arrays.Allacate<double>(1, 1);
                ar[0][0] = ((NumericValueEval) eval).NumberValue;
            }
            else
            {
                throw new EvaluationException(ErrorEval.VALUE_INVALID);
            }

            return ar;
        }

        private static double[][] GetDefaultArrayOneD(int w)
        {
            double[][] array = Arrays.Allacate<double>(w, 1);
            for(int i = 0; i < w; i++)
            {
                array[i][0] = i + 1;
            }
            return array;
        }

        private static double[] FlattenArray(double[][] twoD)
        {
            if(twoD.Length < 1)
            {
                return Array.Empty<double>();
            }
            double[] oneD = new double[twoD.Length * twoD[0].Length];
            for(int i = 0; i < twoD.Length; i++)
            {
                for(int j = 0; j < twoD[0].Length; j++)
                {
                    oneD[i * twoD[0].Length + j] = twoD[i][j];
                }
            }
            return oneD;
        }

        private static double[][] FlattenArrayToRow(double[][] twoD)
        {
            if(twoD.Length < 1)
            {
                return [];
            }
            double[][] oneD = Arrays.Allacate<double>(twoD.Length * twoD[0].Length, 1);
            for(int i = 0; i < twoD.Length; i++)
            {
                for(int j = 0; j < twoD[0].Length; j++)
                {
                    oneD[i * twoD[0].Length + j][0] = twoD[i][j];
                }
            }
            return oneD;
        }

        private static double[][] SwitchRowsColumns(double[][] array)
        {
            //double[][] newArray = new double[array[0].Length][array.Length];
            double[][] newArray = Arrays.Allacate<double>(array[0].Length, array.Length);
            for(int i = 0; i < array.Length; i++)
            {
                for(int j = 0; j < array[0].Length; j++)
                {
                    newArray[j][i] = array[i][j];
                }
            }
            return newArray;
        }

        /// <summary>
        /// Check if all columns in a matrix contain the same values.
        /// Return true if the number of distinct values in each column is 1.
        /// </summary>
        /// <param name="matrix"> column-oriented matrix. A Row matrix should be transposed to column .</param>
        /// <returns>true if all columns contain the same value</returns>
        private static bool IsAllColumnsSame(double[][] matrix)
        {
            if(matrix.Length == 0)
                return false;

            bool[] cols = new bool[matrix[0].Length];
            for(int j = 0; j < matrix[0].Length; j++)
            {
                double prev = Double.NaN;
                for(int i = 0; i < matrix.Length; i++)
                {
                    double v = matrix[i][j];
                    if(i > 0 && v != prev)
                    {
                        cols[j] = true;
                        break;
                    }
                    prev = v;
                }
            }
            bool allEquals = true;
            foreach(bool x in cols)
            {
                if(x)
                {
                    allEquals = false;
                    break;
                }
            }
            ;
            return allEquals;

        }

        private static TrendResults GetNewY(ValueEval[] args)
        {
            double[][] xOrig;
            double[][] x;
            double[][] yOrig;
            double[] y;
            double[][] newXOrig;
            double[][] newX;
            double[][] resultSize;
            bool passThroughOrigin = false;
            switch(args.Length)
            {
                case 1:
                    yOrig = EvalToArray(args[0]);
                    xOrig = [];
                    newXOrig = [];
                    break;
                case 2:
                    yOrig = EvalToArray(args[0]);
                    xOrig = EvalToArray(args[1]);
                    newXOrig = [];
                    break;
                case 3:
                    yOrig = EvalToArray(args[0]);
                    xOrig = EvalToArray(args[1]);
                    newXOrig = EvalToArray(args[2]);
                    break;
                case 4:
                    yOrig = EvalToArray(args[0]);
                    xOrig = EvalToArray(args[1]);
                    newXOrig = EvalToArray(args[2]);
                    if(!(args[3] is BoolEval))
                    {
                        throw new EvaluationException(ErrorEval.VALUE_INVALID);
                    }
                    // The argument in Excel is false when it *should* pass through the origin.
                    passThroughOrigin = !((BoolEval) args[3]).BooleanValue;
                    break;
                default:
                    throw new EvaluationException(ErrorEval.VALUE_INVALID);
            }

            if(yOrig.Length < 1)
            {
                throw new EvaluationException(ErrorEval.VALUE_INVALID);
            }
            y = FlattenArray(yOrig);
            newX = newXOrig;

            if(newXOrig.Length > 0)
            {
                resultSize = newXOrig;
            }
            else
            {
                //resultSize = new double[1][1];
                resultSize = Arrays.Allacate<double>(1, 1);
            }

            if(y.Length == 1)
            {
                /* See comment at top of file
                if (xOrig.Length > 0 && !(xOrig.Length == 1 || xOrig[0].Length == 1)) {
                    throw new EvaluationException(ErrorEval.REF_INVALID);
                } else if (xOrig.Length < 1) {
                    x = new double[1][1];
                    x[0][0] = 1;
                } else {
                    x = new double[1][];
                    x[0] = flattenArray(xOrig);
                    if (newXOrig.Length < 1) {
                        resultSize = xOrig;
                    }
                }*/
                throw new NotImplementedException("Sample size too small");
            }
            else if(yOrig.Length == 1 || yOrig[0].Length == 1)
            {
                if(xOrig.Length < 1)
                {
                    x = GetDefaultArrayOneD(y.Length);
                    if(newXOrig.Length < 1)
                    {
                        resultSize = yOrig;
                    }
                }
                else
                {
                    x = xOrig;
                    if(xOrig[0].Length > 1 && yOrig.Length == 1)
                    {
                        x = SwitchRowsColumns(x);
                    }
                    if(newXOrig.Length < 1)
                    {
                        resultSize = xOrig;
                    }
                }
                if(newXOrig.Length > 0 && (x.Length == 1 || x[0].Length == 1))
                {
                    newX = FlattenArrayToRow(newXOrig);
                }
            }
            else
            {
                if(xOrig.Length < 1)
                {
                    x = GetDefaultArrayOneD(y.Length);
                    if(newXOrig.Length < 1)
                    {
                        resultSize = yOrig;
                    }
                }
                else
                {
                    x = FlattenArrayToRow(xOrig);
                    if(newXOrig.Length < 1)
                    {
                        resultSize = xOrig;
                    }
                }
                if(newXOrig.Length > 0)
                {
                    newX = FlattenArrayToRow(newXOrig);
                }
                if(y.Length != x.Length || yOrig.Length != xOrig.Length)
                {
                    throw new EvaluationException(ErrorEval.REF_INVALID);
                }
            }

            if(newXOrig.Length < 1)
            {
                newX = x;
            }
            else if(newXOrig.Length == 1 && newXOrig[0].Length > 1 && xOrig.Length > 1 && xOrig[0].Length == 1)
            {
                newX = SwitchRowsColumns(newXOrig);
            }

            if(newX[0].Length != x[0].Length)
            {
                throw new EvaluationException(ErrorEval.REF_INVALID);
            }

            if(x[0].Length >= x.Length)
            {
                /* See comment at top of file */
                throw new NotImplementedException("Sample size too small");
            }

            int resultHeight = resultSize.Length;
            int resultWidth = resultSize[0].Length;
            double[] result;
            if(IsAllColumnsSame(x))
            {
                result = new double[newX.Length];
                double avg = y.Average();
                for(int i = 0; i < result.Length; i++)
                    result[i] = avg;
                return new TrendResults(result, resultWidth, resultHeight);
            }

            //MultipleRegression reg = new MultipleRegression();
            //if(passThroughOrigin)
            //{
            //    reg.SetNoIntercept(true);
            //}

            //try
            //{
                
            //    reg.newSampleData(y, x);
            //}
            //catch(ArgumentException e)
            //{
            //    throw new EvaluationException(ErrorEval.REF_INVALID);
            //}
            double[] par;
            //try
            //{
            //    par = reg.estimateRegressionParameters();

            //}
            //catch(SingularMatrixException e)
            //{
            //    throw new NotImplementedException("Singular matrix in input");
            //}
            par = MultipleRegression.DirectMethod(x, y, !passThroughOrigin);
            result = new double[newX.Length];
            for(int i = 0; i < newX.Length; i++)
            {
                result[i] = 0;
                if(passThroughOrigin)
                {
                    for(int j = 0; j < par.Length; j++)
                    {
                        result[i] += par[j] * newX[i][j];
                    }
                }
                else
                {
                    result[i] = par[0];
                    for(int j = 1; j < par.Length; j++)
                    {
                        result[i] += par[j] * newX[i][j - 1];
                    }
                }
            }
            return new TrendResults(result, resultWidth, resultHeight);
        }
    }
}
