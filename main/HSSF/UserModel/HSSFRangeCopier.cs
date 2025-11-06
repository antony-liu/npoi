/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */


using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.HSSF.UserModel
{
    using NPOI.HSSF.Record.Aggregates;
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.UserModel;

    public class HSSFRangeCopier : RangeCopier
    {
        public HSSFRangeCopier(ISheet sourceSheet, ISheet destSheet)
            : base(sourceSheet, destSheet)
        {

        }

        protected override void AdjustCellReferencesInsideFormula(ICell cell, ISheet destSheet, int deltaX, int deltaY)
        {
            FormulaRecordAggregate fra = (FormulaRecordAggregate)((HSSFCell)cell).CellValueRecord;
            int destSheetIndex = destSheet.Workbook.GetSheetIndex(destSheet);
            Ptg[] ptgs = fra.FormulaTokens;
            if(adjustInBothDirections(ptgs, destSheetIndex, deltaX, deltaY))
                fra.SetParsedExpression(ptgs);
        }
    }
}