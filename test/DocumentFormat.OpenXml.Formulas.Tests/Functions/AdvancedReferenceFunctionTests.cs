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
/// Tests for advanced reference functions: OFFSET and INDIRECT.
/// </summary>
public class AdvancedReferenceFunctionTests
{
    #region OFFSET Function Tests

    [Fact]
    public void Offset_BasicOffset_ReturnsCorrectCell()
    {
        var func = OffsetFunction.Instance;
        var context = CreateContextWithData();

        // Reference A1 with offset (2, 3) should give us D3
        var args = new[]
        {
            CellValue.FromString("A1"),
            CellValue.FromNumber(2), // rows offset
            CellValue.FromNumber(3), // cols offset
        };

        var result = func.Execute(context, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(43.0, result.NumericValue); // D3 = 43
    }

    [Fact]
    public void Offset_NoOffset_ReturnsSameCell()
    {
        var func = OffsetFunction.Instance;
        var context = CreateContextWithData();

        var args = new[]
        {
            CellValue.FromString("B2"),
            CellValue.FromNumber(0), // no row offset
            CellValue.FromNumber(0), // no col offset
        };

        var result = func.Execute(context, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(21.0, result.NumericValue); // B2 = 21
    }

    [Fact]
    public void Offset_NegativeOffset_ReturnsCorrectCell()
    {
        var func = OffsetFunction.Instance;
        var context = CreateContextWithData();

        // Reference C3 with offset (-1, -1) should give us B2
        var args = new[]
        {
            CellValue.FromString("C3"),
            CellValue.FromNumber(-1), // rows offset
            CellValue.FromNumber(-1), // cols offset
        };

        var result = func.Execute(context, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(21.0, result.NumericValue); // B2 = 21
    }

    [Fact]
    public void Offset_WithHeight_SingleCell()
    {
        var func = OffsetFunction.Instance;
        var context = CreateContextWithData();

        var args = new[]
        {
            CellValue.FromString("A1"),
            CellValue.FromNumber(1), // rows offset
            CellValue.FromNumber(1), // cols offset
            CellValue.FromNumber(1), // height (single cell)
        };

        var result = func.Execute(context, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(21.0, result.NumericValue); // B2 = 21
    }

    [Fact]
    public void Offset_WithHeightAndWidth_SingleCell()
    {
        var func = OffsetFunction.Instance;
        var context = CreateContextWithData();

        var args = new[]
        {
            CellValue.FromString("A1"),
            CellValue.FromNumber(2), // rows offset
            CellValue.FromNumber(1), // cols offset
            CellValue.FromNumber(1), // height
            CellValue.FromNumber(1), // width
        };

        var result = func.Execute(context, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(31.0, result.NumericValue); // B3 = 31
    }

    [Fact]
    public void Offset_OutOfBoundsPositive_ReturnsError()
    {
        var func = OffsetFunction.Instance;
        var context = CreateContextWithData();

        // Offset beyond valid Excel range
        var args = new[]
        {
            CellValue.FromString("A1"),
            CellValue.FromNumber(1048576), // beyond max row
            CellValue.FromNumber(0),
        };

        var result = func.Execute(context, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    [Fact]
    public void Offset_OutOfBoundsNegative_ReturnsError()
    {
        var func = OffsetFunction.Instance;
        var context = CreateContextWithData();

        // Offset to negative row
        var args = new[]
        {
            CellValue.FromString("A1"),
            CellValue.FromNumber(-1), // would give row 0
            CellValue.FromNumber(0),
        };

        var result = func.Execute(context, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    [Fact]
    public void Offset_InvalidHeight_ReturnsError()
    {
        var func = OffsetFunction.Instance;
        var context = CreateContextWithData();

        var args = new[]
        {
            CellValue.FromString("A1"),
            CellValue.FromNumber(0),
            CellValue.FromNumber(0),
            CellValue.FromNumber(0), // invalid height (must be >= 1)
        };

        var result = func.Execute(context, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    [Fact]
    public void Offset_InvalidWidth_ReturnsError()
    {
        var func = OffsetFunction.Instance;
        var context = CreateContextWithData();

        var args = new[]
        {
            CellValue.FromString("A1"),
            CellValue.FromNumber(0),
            CellValue.FromNumber(0),
            CellValue.FromNumber(1),
            CellValue.FromNumber(-1), // invalid width
        };

        var result = func.Execute(context, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    [Fact]
    public void Offset_InsufficientArguments_ReturnsError()
    {
        var func = OffsetFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("A1"),
            CellValue.FromNumber(1), // missing cols offset
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Offset_InvalidReference_ReturnsError()
    {
        var func = OffsetFunction.Instance;
        var context = CreateContextWithData();

        var args = new[]
        {
            CellValue.FromString("INVALID"),
            CellValue.FromNumber(1),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(context, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    [Fact]
    public void Offset_ErrorInRowsOffset_PropagatesError()
    {
        var func = OffsetFunction.Instance;
        var context = CreateContextWithData();

        var args = new[]
        {
            CellValue.FromString("A1"),
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(context, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void Offset_AbsoluteReference_Works()
    {
        var func = OffsetFunction.Instance;
        var context = CreateContextWithData();

        var args = new[]
        {
            CellValue.FromString("$B$2"),
            CellValue.FromNumber(1),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(context, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(32.0, result.NumericValue); // C3 = 32
    }

    #endregion

    #region INDIRECT Function Tests

    [Fact]
    public void Indirect_A1Notation_ReturnsCorrectCell()
    {
        var func = IndirectFunction.Instance;
        var context = CreateContextWithData();

        var args = new[]
        {
            CellValue.FromString("B2"),
        };

        var result = func.Execute(context, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(21.0, result.NumericValue); // B2 = 21
    }

    [Fact]
    public void Indirect_A1NotationWithDollarSigns_ReturnsCorrectCell()
    {
        var func = IndirectFunction.Instance;
        var context = CreateContextWithData();

        var args = new[]
        {
            CellValue.FromString("$C$3"),
        };

        var result = func.Execute(context, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(32.0, result.NumericValue); // C3 = 32
    }

    [Fact]
    public void Indirect_R1C1Notation_ReturnsCorrectCell()
    {
        var func = IndirectFunction.Instance;
        var context = CreateContextWithData();

        // R2C2 = B2
        var args = new[]
        {
            CellValue.FromString("R2C2"),
            CellValue.FromBool(false), // R1C1 notation
        };

        var result = func.Execute(context, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(21.0, result.NumericValue); // B2 = 21
    }

    [Fact]
    public void Indirect_R1C1NotationRelative_ReturnsCorrectCell()
    {
        var func = IndirectFunction.Instance;
        var context = CreateContextWithDataAndCurrentCell("B2");

        // R[1]C[1] from B2 (row 2, col 2) = C3 (row 3, col 3)
        var args = new[]
        {
            CellValue.FromString("R[1]C[1]"),
            CellValue.FromBool(false), // R1C1 notation
        };

        var result = func.Execute(context, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(32.0, result.NumericValue); // C3 = 32
    }

    [Fact]
    public void Indirect_R1C1NotationRelativeNegative_ReturnsCorrectCell()
    {
        var func = IndirectFunction.Instance;
        var context = CreateContextWithDataAndCurrentCell("C3");

        // R[-1]C[-1] from C3 (row 3, col 3) = B2 (row 2, col 2)
        var args = new[]
        {
            CellValue.FromString("R[-1]C[-1]"),
            CellValue.FromBool(false), // R1C1 notation
        };

        var result = func.Execute(context, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(21.0, result.NumericValue); // B2 = 21
    }

    [Fact]
    public void Indirect_R1C1NotationZeroOffset_ReturnsSameCell()
    {
        var func = IndirectFunction.Instance;
        var context = CreateContextWithDataAndCurrentCell("B2");

        // R[0]C[0] from B2 = B2
        var args = new[]
        {
            CellValue.FromString("R[0]C[0]"),
            CellValue.FromBool(false), // R1C1 notation
        };

        var result = func.Execute(context, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(21.0, result.NumericValue); // B2 = 21
    }

    [Fact]
    public void Indirect_DefaultA1Notation_Works()
    {
        var func = IndirectFunction.Instance;
        var context = CreateContextWithData();

        // Default is A1 notation (TRUE)
        var args = new[]
        {
            CellValue.FromString("D3"),
        };

        var result = func.Execute(context, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(43.0, result.NumericValue); // D3 = 43
    }

    [Fact]
    public void Indirect_A1NotationExplicitTrue_Works()
    {
        var func = IndirectFunction.Instance;
        var context = CreateContextWithData();

        var args = new[]
        {
            CellValue.FromString("A1"),
            CellValue.FromBool(true), // A1 notation
        };

        var result = func.Execute(context, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(10.0, result.NumericValue); // A1 = 10
    }

    [Fact]
    public void Indirect_A1NotationWithNumericTrue_Works()
    {
        var func = IndirectFunction.Instance;
        var context = CreateContextWithData();

        var args = new[]
        {
            CellValue.FromString("B3"),
            CellValue.FromNumber(1), // Non-zero = TRUE
        };

        var result = func.Execute(context, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(31.0, result.NumericValue); // B3 = 31
    }

    [Fact]
    public void Indirect_R1C1NotationWithNumericFalse_Works()
    {
        var func = IndirectFunction.Instance;
        var context = CreateContextWithData();

        var args = new[]
        {
            CellValue.FromString("R1C1"),
            CellValue.FromNumber(0), // Zero = FALSE
        };

        var result = func.Execute(context, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(10.0, result.NumericValue); // R1C1 = A1 = 10
    }

    [Fact]
    public void Indirect_InvalidReference_ReturnsError()
    {
        var func = IndirectFunction.Instance;
        var context = CreateContextWithData();

        var args = new[]
        {
            CellValue.FromString("INVALID"),
        };

        var result = func.Execute(context, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    [Fact]
    public void Indirect_OutOfBoundsR1C1_ReturnsError()
    {
        var func = IndirectFunction.Instance;
        var context = CreateContextWithData();

        // Row 0 is invalid
        var args = new[]
        {
            CellValue.FromString("R0C1"),
            CellValue.FromBool(false),
        };

        var result = func.Execute(context, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    [Fact]
    public void Indirect_NoArguments_ReturnsError()
    {
        var func = IndirectFunction.Instance;
        var args = new CellValue[0];

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Indirect_NonTextReference_ReturnsError()
    {
        var func = IndirectFunction.Instance;
        var context = CreateContextWithData();

        var args = new[]
        {
            CellValue.FromNumber(123), // Not a text reference
        };

        var result = func.Execute(context, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Indirect_ErrorInReference_PropagatesError()
    {
        var func = IndirectFunction.Instance;
        var context = CreateContextWithData();

        var args = new[]
        {
            CellValue.Error("#N/A"),
        };

        var result = func.Execute(context, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    [Fact]
    public void Indirect_WithSheetName_ParsesCorrectly()
    {
        var func = IndirectFunction.Instance;
        var context = CreateContextWithData();

        // Sheet names are stripped, so Sheet1!A1 becomes A1
        var args = new[]
        {
            CellValue.FromString("Sheet1!A1"),
        };

        var result = func.Execute(context, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(10.0, result.NumericValue); // A1 = 10
    }

    [Fact]
    public void Indirect_WithQuotedSheetName_ParsesCorrectly()
    {
        var func = IndirectFunction.Instance;
        var context = CreateContextWithData();

        // Quoted sheet names with spaces
        var args = new[]
        {
            CellValue.FromString("'My Sheet'!B2"),
        };

        var result = func.Execute(context, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(21.0, result.NumericValue); // B2 = 21
    }

    #endregion

    #region OFFSET and INDIRECT Combination Tests

    [Fact]
    public void OffsetWithIndirect_CombinedPattern_ReturnsCorrectValue()
    {
        var indirectFunc = IndirectFunction.Instance;
        var offsetFunc = OffsetFunction.Instance;
        var context = CreateContextWithData();

        // First, use INDIRECT to get reference to "A1"
        var indirectArgs = new[]
        {
            CellValue.FromString("A1"),
        };

        var indirectResult = indirectFunc.Execute(context, indirectArgs);
        Assert.Equal(10.0, indirectResult.NumericValue);

        // Then, use OFFSET from A1
        var offsetArgs = new[]
        {
            CellValue.FromString("A1"),
            CellValue.FromNumber(1),
            CellValue.FromNumber(1),
        };

        var offsetResult = offsetFunc.Execute(context, offsetArgs);
        Assert.Equal(21.0, offsetResult.NumericValue); // B2 = 21
    }

    #endregion

    #region Helper Methods

    private static CellContext CreateContextWithData()
    {
        var worksheet = new Worksheet();
        var sheetData = new SheetData();
        worksheet.Append(sheetData);

        // Create a 4x4 grid of test data
        // A1=10, B1=11, C1=12, D1=13
        // A2=20, B2=21, C2=22, D2=23
        // A3=30, B3=31, C3=32, D3=43
        // A4=40, B4=41, C4=42, D4=43

        for (int row = 1; row <= 4; row++)
        {
            var rowElement = new Row { RowIndex = (uint)row };
            for (int col = 1; col <= 4; col++)
            {
                var cellRef = GetColumnLetter(col) + row.ToString();
                var value = (row - 1) * 10 + col + (col - 1) * 10;

                var cell = new Cell
                {
                    CellReference = cellRef,
                    CellValue = new Spreadsheet.CellValue(value.ToString()),
                    DataType = null, // Number type
                };
                rowElement.Append(cell);
            }
            sheetData.Append(rowElement);
        }

        return new CellContext(worksheet);
    }

    private static CellContext CreateContextWithDataAndCurrentCell(string currentCellReference)
    {
        var context = CreateContextWithData();
        context.CurrentCellReference = currentCellReference;
        return context;
    }

    private static string GetColumnLetter(int column)
    {
        var result = string.Empty;

        while (column > 0)
        {
            var modulo = (column - 1) % 26;
            result = (char)('A' + modulo) + result;
            column = (column - modulo) / 26;
        }

        return result;
    }

    #endregion
}
