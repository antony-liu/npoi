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

namespace TestCases.SS.UserModel
{
    using NPOI.SS.UserModel;
    using NPOI.SS.UserModel.Helpers;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;

    public abstract class BaseTestColumnShifting
    {
        protected IWorkbook wb;
        protected ISheet sheet1;
        protected ColumnShifter columnShifter;

        public virtual void Init()
        {
            int rowIndex = 0;
            sheet1 = wb.CreateSheet("sheet1");
            IRow row = sheet1.CreateRow(rowIndex++);
            row.CreateCell(0, CellType.Numeric).SetCellValue(0);
            row.CreateCell(3, CellType.Numeric).SetCellValue(3);
            row.CreateCell(4, CellType.Numeric).SetCellValue(4);

            row = sheet1.CreateRow(rowIndex++);
            row.CreateCell(0, CellType.Numeric).SetCellValue(0.1);
            row.CreateCell(1, CellType.Numeric).SetCellValue(1.1);
            row.CreateCell(2, CellType.Numeric).SetCellValue(2.1);
            row.CreateCell(3, CellType.Numeric).SetCellValue(3.1);
            row.CreateCell(4, CellType.Numeric).SetCellValue(4.1);
            row.CreateCell(5, CellType.Numeric).SetCellValue(5.1);
            row.CreateCell(6, CellType.Numeric).SetCellValue(6.1);
            row.CreateCell(7, CellType.Numeric).SetCellValue(7.1);
            row = sheet1.CreateRow(rowIndex++);
            row.CreateCell(3, CellType.Numeric).SetCellValue(3.2);
            row.CreateCell(5, CellType.Numeric).SetCellValue(5.2);
            row.CreateCell(7, CellType.Numeric).SetCellValue(7.2);

            InitColumnShifter();
        }

        protected virtual void InitColumnShifter() { }

        [Test]
        public void TestShift3ColumnsRight()
        {
            columnShifter.ShiftColumns(1, 2, 3);

            ICell cell = sheet1.GetRow(0).GetCell(4);
            ClassicAssert.IsNull(cell);
            cell = sheet1.GetRow(1).GetCell(4);
            ClassicAssert.AreEqual(1.1, cell.NumericCellValue, 0.01);
            cell = sheet1.GetRow(1).GetCell(5);
            ClassicAssert.AreEqual(2.1, cell.NumericCellValue, 0.01);
            cell = sheet1.GetRow(2).GetCell(4);
            ClassicAssert.IsNull(cell);
        }

        [Test]
        public void testShiftLeft()
        {
            try
            {
                columnShifter.ShiftColumns(1, 2, -3);
                ClassicAssert.IsTrue(false, "Shift to negative indices should throw exception");
            }
            catch(InvalidOperationException e)
            {
                ClassicAssert.IsTrue(true);
            }
        }
    }
}

