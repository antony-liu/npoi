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

namespace TestCases.SS.UserModel
{
    using NPOI.SS.Formula;
    using NPOI.SS.UserModel;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;

    public abstract class BaseTestXEvaluationSheet
    {
        /// <summary>
        /// Get a pair of underlying sheet and evaluation sheet.
        /// </summary>
        protected abstract KeyValuePair<ISheet, IEvaluationSheet> GetInstance();

        [Test]
        public void LastRowNumIsUpdatedFromUnderlyingSheet_bug62993()
        {
            KeyValuePair<ISheet, IEvaluationSheet> sheetPair = GetInstance();
            ISheet underlyingSheet = sheetPair.Key;
            IEvaluationSheet instance = sheetPair.Value;

            ClassicAssert.AreEqual(0, instance.LastRowNum);

            underlyingSheet.CreateRow(0);
            underlyingSheet.CreateRow(1);
            underlyingSheet.CreateRow(2);
            ClassicAssert.AreEqual(2, instance.LastRowNum);

            underlyingSheet.RemoveRow(underlyingSheet.GetRow(2));
            ClassicAssert.AreEqual(1, instance.LastRowNum);
        }
    }
}
