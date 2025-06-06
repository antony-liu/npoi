﻿using System;
using NPOI;
using NPOI.SS.Util;
using NUnit.Framework;using NUnit.Framework.Legacy;

namespace TestCases;

public class CellReferenceParserTest
{
    [TestCase("A", char.MinValue, "A", char.MinValue, "")]
    [TestCase("1", char.MinValue, "", char.MinValue, "1")]
    [TestCase("$A$1", '$', "A", '$', "1")]
    [TestCase("$AB$12", '$', "AB", '$', "12")]
    [TestCase("$A$12", '$', "A", '$', "12")]
    [TestCase("$AB$1", '$', "AB", '$', "1")]
    [TestCase("A1", char.MinValue, "A", char.MinValue, "1")]
    [TestCase("AB1", char.MinValue, "AB", char.MinValue, "1")]
    [TestCase("$A123", '$', "A", char.MinValue, "123")]
    [TestCase("$123", char.MinValue, "", '$', "123")]
    public void TryParseCellReferenceShouldSucceedForValidInput(string input, char expectedColumnPrefix, string expectedColumn, char expectedRowPrefix, string expectedRow)
    {
        ClassicAssert.True(CellReferenceParser.TryParseCellReference(input.AsSpan(), out var columnPrefix, out var column, out var rowPrefix, out var row));
        ClassicAssert.AreEqual(expectedColumnPrefix, columnPrefix, "Column prefix mismatch");
        ClassicAssert.AreEqual(expectedColumn, column.ToString(), "Column mismatch");
        ClassicAssert.AreEqual(expectedRowPrefix, rowPrefix, "Row prefix mismatch");
        ClassicAssert.AreEqual(expectedRow, row.ToString(), "Row mismatch");
    }

    [TestCase("$1$1")]
    [TestCase("1$1")]
    [TestCase("1$")]
    public void TryParseCellReferenceShouldFailForInvalidInput(string input)
    {
        ClassicAssert.False(CellReferenceParser.TryParseCellReference(input.AsSpan(), out _, out _, out _, out _));
    }

    [TestCase("$A$1", "A", "1")]
    [TestCase("$AB$12", "AB", "12")]
    [TestCase("$A$12", "A", "12")]
    [TestCase("$AB$1", "AB", "1")]
    [TestCase("A1", "A", "1")]
    [TestCase("AB1", "AB", "1")]
    [TestCase("$A123", "A", "123")]
    public void TryParseStrictCellReferenceShouldSucceedForValidInput(string input, string expectedColumn, string expectedRow)
    {
        ClassicAssert.True(CellReferenceParser.TryParseStrictCellReference(input.AsSpan(), out var column, out var row));
        ClassicAssert.AreEqual(expectedColumn, column.ToString(), "Column mismatch");
        ClassicAssert.AreEqual(expectedRow, row.ToString(), "Row mismatch");
    }

    [TestCase("$1$1")]
    [TestCase("1$1")]
    [TestCase("1$")]
    [TestCase("A")]
    [TestCase("1")]
    public void TryParseStrictCellReferenceShouldFailForInvalidInput(string input)
    {
        ClassicAssert.False(CellReferenceParser.TryParseStrictCellReference(input.AsSpan(), out _, out _));
    }

    [TestCase("A", "A")]
    [TestCase("$A", "A")]
    [TestCase("$ABC", "ABC")]
    public void TryParseColumnReferenceShouldSucceedForValidInput(string input, string expectedColumn)
    {
        ClassicAssert.True(CellReferenceParser.TryParseColumnReference(input.AsSpan(), out var column));
        ClassicAssert.AreEqual(expectedColumn, column.ToString(), "Column mismatch");
    }

    [TestCase("1")]
    [TestCase("$A$1")]
    [TestCase("$AB$12")]
    [TestCase("$AB$")]
    [TestCase("$A$12")]
    [TestCase("$AB$1")]
    [TestCase("A1")]
    [TestCase("AB1")]
    [TestCase("$A123")]
    public void TryParseColumnReferenceShouldFailForInvalidInput(string input)
    {
        ClassicAssert.False(CellReferenceParser.TryParseColumnReference(input.AsSpan(), out _));
    }


    [TestCase("1", "1")]
    [TestCase("123", "123")]
    [TestCase("$123", "123")]
    public void TryParseRowReferenceShouldSucceedForValidInput(string input, string expectedRow)
    {
        ClassicAssert.True(CellReferenceParser.TryParseRowReference(input.AsSpan(), out var row));
        ClassicAssert.AreEqual(expectedRow, row.ToString(), "Row mismatch");
    }

    [TestCase("$")]
    [TestCase("A$")]
    [TestCase("$A")]
    [TestCase("A")]
    public void TryParseRowReferenceShouldFailForInvalidInput(string input)
    {
        ClassicAssert.False(CellReferenceParser.TryParseRowReference(input.AsSpan(), out _));
    }
}