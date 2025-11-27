// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for reference and lookup functions: COLUMN, ROW, COLUMNS, ROWS, ADDRESS.
/// </summary>
public class ReferenceFunctionTests
{
    #region COLUMN Function Tests

    [Fact]
    public void Column_WithReference_ReturnsColumnNumber()
    {
        var func = ColumnFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("A1"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void Column_WithReferenceB10_ReturnsTwo()
    {
        var func = ColumnFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("B10"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(2.0, result.NumericValue);
    }

    [Fact]
    public void Column_WithReferenceZ1_Returns26()
    {
        var func = ColumnFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Z1"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(26.0, result.NumericValue);
    }

    [Fact]
    public void Column_WithReferenceAA1_Returns27()
    {
        var func = ColumnFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("AA1"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(27.0, result.NumericValue);
    }

    [Fact]
    public void Column_WithAbsoluteReference_ReturnsColumnNumber()
    {
        var func = ColumnFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("$C$5"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(3.0, result.NumericValue);
    }

    [Fact]
    public void Column_WithoutReference_UsesCurrentCell()
    {
        var func = ColumnFunction.Instance;
        var context = CreateMockContext("D10");
        var args = new FormulaResult[0];

        var result = func.Execute(context, args);

        Assert.Equal(4.0, result.NumericValue);
    }

    [Fact]
    public void Column_WithoutReferenceNoContext_ReturnsError()
    {
        var func = ColumnFunction.Instance;
        var args = new FormulaResult[0];

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Column_WithInvalidReference_ReturnsError()
    {
        var func = ColumnFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("INVALID"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Column_WithError_PropagatesError()
    {
        var func = ColumnFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#REF!"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    #endregion

    #region ROW Function Tests

    [Fact]
    public void Row_WithReference_ReturnsRowNumber()
    {
        var func = RowFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("A1"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void Row_WithReferenceB10_ReturnsTen()
    {
        var func = RowFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("B10"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void Row_WithReferenceZ100_Returns100()
    {
        var func = RowFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Z100"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(100.0, result.NumericValue);
    }

    [Fact]
    public void Row_WithAbsoluteReference_ReturnsRowNumber()
    {
        var func = RowFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("$C$5"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void Row_WithoutReference_UsesCurrentCell()
    {
        var func = RowFunction.Instance;
        var context = CreateMockContext("D10");
        var args = new FormulaResult[0];

        var result = func.Execute(context, args);

        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void Row_WithoutReferenceNoContext_ReturnsError()
    {
        var func = RowFunction.Instance;
        var args = new FormulaResult[0];

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Row_WithInvalidReference_ReturnsError()
    {
        var func = RowFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("INVALID"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Row_WithError_PropagatesError()
    {
        var func = RowFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#DIV/0!"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    #endregion

    #region COLUMNS Function Tests

    [Fact]
    public void Columns_SingleCell_ReturnsOne()
    {
        var func = ColumnsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void Columns_2x2Array_ReturnsTwo()
    {
        var func = ColumnsFunction.Instance;
        // 2x2 array
        var args = new[]
        {
            FormulaResult.FromNumber(10),
            FormulaResult.FromNumber(20),
            FormulaResult.FromNumber(30),
            FormulaResult.FromNumber(40),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(2.0, result.NumericValue);
    }

    [Fact]
    public void Columns_3x2Array_ReturnsTwo()
    {
        var func = ColumnsFunction.Instance;
        // 3x2 array (3 rows, 2 columns)
        var args = new[]
        {
            FormulaResult.FromNumber(10),
            FormulaResult.FromNumber(20),
            FormulaResult.FromNumber(30),
            FormulaResult.FromNumber(40),
            FormulaResult.FromNumber(50),
            FormulaResult.FromNumber(60),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(2.0, result.NumericValue);
    }

    [Fact]
    public void Columns_1x5Array_ReturnsFive()
    {
        var func = ColumnsFunction.Instance;
        // 1x5 array (1 row, 5 columns)
        var args = new[]
        {
            FormulaResult.FromNumber(10),
            FormulaResult.FromNumber(20),
            FormulaResult.FromNumber(30),
            FormulaResult.FromNumber(40),
            FormulaResult.FromNumber(50),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void Columns_WithError_PropagatesError()
    {
        var func = ColumnsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(10),
            FormulaResult.Error("#N/A"),
            FormulaResult.FromNumber(30),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    [Fact]
    public void Columns_NoArguments_ReturnsError()
    {
        var func = ColumnsFunction.Instance;
        var args = new FormulaResult[0];

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region ROWS Function Tests

    [Fact]
    public void Rows_SingleCell_ReturnsOne()
    {
        var func = RowsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void Rows_2x2Array_ReturnsTwo()
    {
        var func = RowsFunction.Instance;
        // 2x2 array
        var args = new[]
        {
            FormulaResult.FromNumber(10),
            FormulaResult.FromNumber(20),
            FormulaResult.FromNumber(30),
            FormulaResult.FromNumber(40),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(2.0, result.NumericValue);
    }

    [Fact]
    public void Rows_3x2Array_ReturnsThree()
    {
        var func = RowsFunction.Instance;
        // 3x2 array (3 rows, 2 columns)
        var args = new[]
        {
            FormulaResult.FromNumber(10),
            FormulaResult.FromNumber(20),
            FormulaResult.FromNumber(30),
            FormulaResult.FromNumber(40),
            FormulaResult.FromNumber(50),
            FormulaResult.FromNumber(60),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(3.0, result.NumericValue);
    }

    [Fact]
    public void Rows_5x1Array_ReturnsFive()
    {
        var func = RowsFunction.Instance;
        // 5x1 array (5 rows, 1 column)
        var args = new[]
        {
            FormulaResult.FromNumber(10),
            FormulaResult.FromNumber(20),
            FormulaResult.FromNumber(30),
            FormulaResult.FromNumber(40),
            FormulaResult.FromNumber(50),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void Rows_WithError_PropagatesError()
    {
        var func = RowsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(10),
            FormulaResult.Error("#REF!"),
            FormulaResult.FromNumber(30),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    [Fact]
    public void Rows_NoArguments_ReturnsError()
    {
        var func = RowsFunction.Instance;
        var args = new FormulaResult[0];

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region ADDRESS Function Tests

    [Fact]
    public void Address_Row1Col1_ReturnsAbsoluteA1()
    {
        var func = AddressFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(1),
            FormulaResult.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("$A$1", result.StringValue);
    }

    [Fact]
    public void Address_Row2Col3_ReturnsAbsoluteC2()
    {
        var func = AddressFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(2),
            FormulaResult.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("$C$2", result.StringValue);
    }

    [Fact]
    public void Address_WithAbsNum1_ReturnsAbsolute()
    {
        var func = AddressFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(5),
            FormulaResult.FromNumber(10),
            FormulaResult.FromNumber(1), // Absolute
        };

        var result = func.Execute(null!, args);

        Assert.Equal("$J$5", result.StringValue);
    }

    [Fact]
    public void Address_WithAbsNum2_ReturnsAbsColRelRow()
    {
        var func = AddressFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(5),
            FormulaResult.FromNumber(10),
            FormulaResult.FromNumber(2), // Absolute column, relative row
        };

        var result = func.Execute(null!, args);

        Assert.Equal("$J5", result.StringValue);
    }

    [Fact]
    public void Address_WithAbsNum3_ReturnsRelColAbsRow()
    {
        var func = AddressFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(5),
            FormulaResult.FromNumber(10),
            FormulaResult.FromNumber(3), // Relative column, absolute row
        };

        var result = func.Execute(null!, args);

        Assert.Equal("J$5", result.StringValue);
    }

    [Fact]
    public void Address_WithAbsNum4_ReturnsRelative()
    {
        var func = AddressFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(5),
            FormulaResult.FromNumber(10),
            FormulaResult.FromNumber(4), // Relative
        };

        var result = func.Execute(null!, args);

        Assert.Equal("J5", result.StringValue);
    }

    [Fact]
    public void Address_WithA1False_ReturnsR1C1Notation()
    {
        var func = AddressFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(5),
            FormulaResult.FromNumber(10),
            FormulaResult.FromNumber(1), // Absolute
            FormulaResult.FromBool(false), // R1C1 notation
        };

        var result = func.Execute(null!, args);

        Assert.Equal("R5C10", result.StringValue);
    }

    [Fact]
    public void Address_WithSheetName_IncludesSheet()
    {
        var func = AddressFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(2),
            FormulaResult.FromNumber(3),
            FormulaResult.FromNumber(1), // Absolute
            FormulaResult.FromBool(true), // A1 notation
            FormulaResult.FromString("Sheet1"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("Sheet1!$C$2", result.StringValue);
    }

    [Fact]
    public void Address_WithSheetNameWithSpace_QuotesSheet()
    {
        var func = AddressFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(2),
            FormulaResult.FromNumber(3),
            FormulaResult.FromNumber(1), // Absolute
            FormulaResult.FromBool(true), // A1 notation
            FormulaResult.FromString("My Sheet"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("'My Sheet'!$C$2", result.StringValue);
    }

    [Fact]
    public void Address_ColumnZ_ReturnsCorrectLetter()
    {
        var func = AddressFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(1),
            FormulaResult.FromNumber(26),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("$Z$1", result.StringValue);
    }

    [Fact]
    public void Address_ColumnAA_ReturnsCorrectLetters()
    {
        var func = AddressFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(1),
            FormulaResult.FromNumber(27),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("$AA$1", result.StringValue);
    }

    [Fact]
    public void Address_InvalidRowNum_ReturnsError()
    {
        var func = AddressFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(0), // Invalid row
            FormulaResult.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Address_InvalidColNum_ReturnsError()
    {
        var func = AddressFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(1),
            FormulaResult.FromNumber(0), // Invalid column
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Address_InvalidAbsNum_ReturnsError()
    {
        var func = AddressFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(1),
            FormulaResult.FromNumber(1),
            FormulaResult.FromNumber(5), // Invalid abs_num (must be 1-4)
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Address_InsufficientArguments_ReturnsError()
    {
        var func = AddressFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(1), // Missing column
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Address_ErrorInRowNum_PropagatesError()
    {
        var func = AddressFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#DIV/0!"),
            FormulaResult.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void Address_ErrorInColNum_PropagatesError()
    {
        var func = AddressFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(1),
            FormulaResult.Error("#REF!"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    #endregion

    #region Helper Methods

    private static CellContext CreateMockContext(string currentCellReference)
    {
        var worksheet = new Worksheet();
        var context = new CellContext(worksheet);
        context.CurrentCellReference = currentCellReference;
        return context;
    }

    #endregion
}
