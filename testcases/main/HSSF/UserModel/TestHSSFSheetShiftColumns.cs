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

namespace TestCases.HSSF.UserModel
{
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NUnit.Framework;
    using TestCases.SS.UserModel;

    public class TestHSSFSheetShiftColumns : BaseTestSheetShiftColumns
    {
        public TestHSSFSheetShiftColumns()
               : base()
        {
            _testDataProvider = HSSFITestDataProvider.Instance;
        }
        [SetUp]
        public override void Init()
        {
            workbook = new HSSFWorkbook();
            base.Init();
        }

        protected override IWorkbook openWorkbook(string spreadsheetFileName)
        {
            return HSSFTestDataSamples.OpenSampleWorkbook(spreadsheetFileName);
        }

        protected override IWorkbook GetReadBackWorkbook(IWorkbook wb)
        {
            return HSSFTestDataSamples.WriteOutAndReadBack((HSSFWorkbook) wb);
        }
        [Ignore("see <https://bz.apache.org/bugzilla/show_bug.cgi?id=62030>")]
        [Test]
        public override void ShiftMergedColumnsToMergedColumnsLeft()
        {
            // This override is used only in order to test Assert.Failing for HSSF. Please remove method After code is fixed on hssf, 
            // so that original method from BaseTestSheetShiftColumns can be executed. 
        }
        [Ignore("see <https://bz.apache.org/bugzilla/show_bug.cgi?id=62030>")]
        [Test]
        public override void ShiftMergedColumnsToMergedColumnsRight()
        {
            // This override is used only in order to test Assert.Failing for HSSF. Please remove method After code is fixed on hssf, 
            // so that original method from BaseTestSheetShiftColumns can be executed. 
        }
        [Ignore("see <https://bz.apache.org/bugzilla/show_bug.cgi?id=62030>")]
        [Test]
        public override void TestBug54524()
        {
            // This override is used only in order to test Assert.Failing for HSSF. Please remove method After code is fixed on hssf, 
            // so that original method from BaseTestSheetShiftColumns can be executed. 
        }
        [Ignore("see <https://bz.apache.org/bugzilla/show_bug.cgi?id=62030>")]
        [Test]
        public override void TestCommentsShifting()
        {
            // This override is used only in order to test Assert.Failing for HSSF. Please remove method After code is fixed on hssf, 
            // so that original method from BaseTestSheetShiftColumns can be executed. 
        }
        [Ignore("see <https://bz.apache.org/bugzilla/show_bug.cgi?id=62030>")]
        [Test]
        public override void TestShiftWithMergedRegions()
        {
            // This override is used only in order to test Assert.Failing for HSSF. Please remove method After code is fixed on hssf, 
            // so that original method from BaseTestSheetShiftColumns can be executed. 
            // After removing, you can re-add 'final' keyword to specification of original method. 
        }
        [Ignore("see <https://bz.apache.org/bugzilla/show_bug.cgi?id=62030>")]
        [Test]
        public override void TestShiftHyperlinks()
        {
        }
    }

}
