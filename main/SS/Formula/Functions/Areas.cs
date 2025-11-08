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

using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.PTG;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.Formula.Functions
{
    public class Areas : Function
    {
        public ValueEval Evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            if (args.Length == 0)
            {
                return ErrorEval.VALUE_INVALID;
            }
            try
            {
                ValueEval valueEval = args[0];
                int result = 1;
                if (valueEval is RefListEval refListEval) {
                    result = refListEval.GetList().Count;
                }
                NumberEval numberEval = new NumberEval(new NumberPtg(result));
                return numberEval;
            }
            catch
            {
                return ErrorEval.VALUE_INVALID;
            }
        }
    }
}
