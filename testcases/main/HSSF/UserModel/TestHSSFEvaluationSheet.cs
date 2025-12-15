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

namespace TestCases.HSSF.UserModel
{
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula;

    using NPOI.SS.UserModel;
    using NUnit.Framework;
    using TestCases.HSSF;
    using TestCases.SS.UserModel;

    public class TestHSSFEvaluationSheet : BaseTestXEvaluationSheet
    {
        protected override KeyValuePair<ISheet, IEvaluationSheet> GetInstance()
        {
            HSSFSheet sheet = new HSSFWorkbook().CreateSheet() as HSSFSheet;
            return new KeyValuePair<ISheet, IEvaluationSheet>(sheet, new HSSFEvaluationSheet(sheet));
        }

        [Test]
        public void TestMissingExternalName()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("external_name.xls") as HSSFWorkbook;
            foreach(IName name in wb.GetAllNames())
            {
                // this sometimes causes exceptions
                if(!name.IsFunctionName)
                {
                    _ = name.RefersToFormula;
                }
            }
        }
    }
}


