using NPOI.HSSF.UserModel;
using NPOI.SS;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NUnit.Framework;
using NUnit.Framework.Legacy;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestCases.SS.Util
{
    [TestFixture]
    public class TestAreaReference
    {
        [Test]
        public void TestWholeColumn()
        {
            AreaReference oldStyle = AreaReference.GetWholeColumn(SpreadsheetVersion.EXCEL97, "A", "B");
            ClassicAssert.AreEqual(0, oldStyle.FirstCell.Col);
            ClassicAssert.AreEqual(0, oldStyle.FirstCell.Row);
            ClassicAssert.AreEqual(1, oldStyle.LastCell.Col);
            ClassicAssert.AreEqual(SpreadsheetVersion.EXCEL97.LastRowIndex, oldStyle.LastCell.Row);
            ClassicAssert.IsTrue(oldStyle.IsWholeColumnReference());

            AreaReference oldStyleNonWholeColumn = new AreaReference("A1:B23", SpreadsheetVersion.EXCEL97);
            ClassicAssert.IsFalse(oldStyleNonWholeColumn.IsWholeColumnReference());

            AreaReference newStyle = AreaReference.GetWholeColumn(SpreadsheetVersion.EXCEL2007, "A", "B");
            ClassicAssert.AreEqual(0, newStyle.FirstCell.Col);
            ClassicAssert.AreEqual(0, newStyle.FirstCell.Row);
            ClassicAssert.AreEqual(1, newStyle.LastCell.Col);
            ClassicAssert.AreEqual(SpreadsheetVersion.EXCEL2007.LastRowIndex, newStyle.LastCell.Row);
            ClassicAssert.IsTrue(newStyle.IsWholeColumnReference());

            AreaReference newStyleNonWholeColumn = new AreaReference("A1:B23", SpreadsheetVersion.EXCEL2007);
            ClassicAssert.IsFalse(newStyleNonWholeColumn.IsWholeColumnReference());
        }
        [Test]
        public void TestWholeRow()
        {
            AreaReference oldStyle = AreaReference.GetWholeRow(SpreadsheetVersion.EXCEL97, "1", "2");
            ClassicAssert.AreEqual(0, oldStyle.FirstCell.Col);
            ClassicAssert.AreEqual(0, oldStyle.FirstCell.Row);
            ClassicAssert.AreEqual(SpreadsheetVersion.EXCEL97.LastColumnIndex, oldStyle.LastCell.Col);
            ClassicAssert.AreEqual(1, oldStyle.LastCell.Row);

            AreaReference newStyle = AreaReference.GetWholeRow(SpreadsheetVersion.EXCEL2007, "1", "2");
            ClassicAssert.AreEqual(0, newStyle.FirstCell.Col);
            ClassicAssert.AreEqual(0, newStyle.FirstCell.Row);
            ClassicAssert.AreEqual(SpreadsheetVersion.EXCEL2007.LastColumnIndex, newStyle.LastCell.Col);
            ClassicAssert.AreEqual(1, newStyle.LastCell.Row);
        }

        [Test]
        public void Test62810()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Ctor test");
            String sheetName = sheet.SheetName;
            CellReference topLeft = new CellReference(sheetName, 1, 1, true, true);
            CellReference bottomRight = new CellReference(sheetName, 5, 10, true, true);
            AreaReference goodAreaRef = new AreaReference(topLeft, bottomRight, SpreadsheetVersion.EXCEL2007);
            AreaReference badAreaRef = new AreaReference(bottomRight, topLeft, SpreadsheetVersion.EXCEL2007);

            ClassicAssert.AreEqual("'Ctor test'!$B$2", topLeft.FormatAsString());
            ClassicAssert.AreEqual("'Ctor test'!$K$6", bottomRight.FormatAsString());
            ClassicAssert.AreEqual("'Ctor test'!$B$2:$K$6", goodAreaRef.FormatAsString());
            ClassicAssert.AreEqual("'Ctor test'!$B$2:$K$6", badAreaRef.FormatAsString());
        }
    }
}
