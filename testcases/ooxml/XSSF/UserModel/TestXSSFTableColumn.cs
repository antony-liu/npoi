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

namespace TestCases.XSSF.UserModel
{
    using NPOI.XSSF;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;
    [TestFixture]
    public sealed class TestXSSFTableColumn
    {

        [Test]
        public void TestGetColumnName()
        {
            XSSFWorkbook wb = XSSFTestDataSamples
                    .OpenSampleWorkbook("CustomXMLMappings-complex-type.xlsx");
            try
            {
                XSSFTable table = wb.GetTable("Tabella2");

                List<XSSFTableColumn> tableColumns = table.GetColumns();

                ClassicAssert.AreEqual("ID", tableColumns[0].Name);
                ClassicAssert.AreEqual("Unmapped Column", tableColumns[1].Name);
                ClassicAssert.AreEqual("SchemaRef", tableColumns[2].Name);
                ClassicAssert.AreEqual("Namespace", tableColumns[3].Name);

            }
            finally
            {
                wb.Close();
            }
        }

        [Test]
        public void TestGetColumnIndex()
        {
            XSSFWorkbook wb = XSSFTestDataSamples
                        .OpenSampleWorkbook("CustomXMLMappings-complex-type.xlsx");
            try
            {
                XSSFTable table = wb.GetTable("Tabella2");

                List<XSSFTableColumn> tableColumns = table.GetColumns();

                ClassicAssert.AreEqual(0, tableColumns[0].ColumnIndex);
                ClassicAssert.AreEqual(1, tableColumns[1].ColumnIndex);
                ClassicAssert.AreEqual(2, tableColumns[2].ColumnIndex);
                ClassicAssert.AreEqual(3, tableColumns[3].ColumnIndex);

            }
            finally { wb.Close(); }
        }

        [Test]
        public void TestGetXmlColumnPrs()
        {
            XSSFWorkbook wb = XSSFTestDataSamples
                            .OpenSampleWorkbook("CustomXMLMappings-complex-type.xlsx");
            try
            {
                XSSFTable table = wb.GetTable("Tabella2");

                List<XSSFTableColumn> tableColumns = table.GetColumns();

                ClassicAssert.IsNotNull(tableColumns[0].GetXmlColumnPr());
                ClassicAssert.IsNull(tableColumns[1].GetXmlColumnPr()); // unmapped column
                ClassicAssert.IsNotNull(tableColumns[2].ColumnIndex);
                ClassicAssert.IsNotNull(tableColumns[3].ColumnIndex);

            }
            finally { wb.Close(); }
        }
    }
}


