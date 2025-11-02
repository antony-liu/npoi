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

namespace NPOI.HSSF.UserModel.Helpers
{
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.UserModel;
    using NPOI.SS.UserModel.Helpers;
    using NPOI.Util;

    /// <summary>
    /// Helper for shifting columns up or down
    /// </summary>
    /// <remarks>
    /// @since POI 4.0.0
    /// </remarks>
    // non-Javadoc: When possible, code should be implemented in the ColumnShifter abstract class to avoid duplication with
    // {@link NPOI.xssf.UserModel.helpers.XSSFColumnShifter}
    public sealed class HSSFColumnShifter : ColumnShifter
    {
        private static  POILogger logger = POILogFactory.GetLogger(typeof(HSSFColumnShifter));

        public HSSFColumnShifter(HSSFSheet sh)
            : base(sh)
        {

        }

        public override void UpdateNamedRanges(FormulaShifter formulaShifter)
        {
            throw new NotImplementedException("HSSFColumnShifter.updateNamedRanges");
        }

        public override void UpdateFormulas(FormulaShifter formulaShifter)
        {
            throw new NotImplementedException("updateFormulas");
        }

        public override void UpdateConditionalFormatting(FormulaShifter formulaShifter)
        {
            throw new NotImplementedException("updateConditionalFormatting");
        }

        public override void UpdateHyperlinks(FormulaShifter formulaShifter)
        {
            throw new NotImplementedException("updateHyperlinks");
        }

        public override void UpdateColumnFormulas(IColumn column, FormulaShifter formulaShifter)
        {
            throw new NotImplementedException("UpdateColumnFormulas");
        }
    }
}

