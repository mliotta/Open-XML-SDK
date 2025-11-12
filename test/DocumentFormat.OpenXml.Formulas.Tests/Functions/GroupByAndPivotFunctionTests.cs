// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for GROUPBY, PIVOTBY, TRIMRANGE, and ANCHORARRAY functions.
/// </summary>
public class GroupByAndPivotFunctionTests
{
    #region GROUPBY Function Tests

    [Fact]
    public void GroupBy_Sum_ReturnsFirstGroupSum()
    {
        var func = GroupByFunction.Instance;
        // row_fields: ["A", "B"], values: [10, 20], function: 1 (SUM)
        var args = new[]
        {
            CellValue.FromString("A"),
            CellValue.FromString("B"),
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(1), // SUM
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void GroupBy_Average_ReturnsFirstGroupAverage()
    {
        var func = GroupByFunction.Instance;
        // row_fields: ["A", "A"], values: [10, 20], function: 2 (AVERAGE)
        var args = new[]
        {
            CellValue.FromString("A"),
            CellValue.FromString("A"),
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(2), // AVERAGE
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(15.0, result.NumericValue);
    }

    [Fact]
    public void GroupBy_Count_ReturnsFirstGroupCount()
    {
        var func = GroupByFunction.Instance;
        // row_fields: ["A", "A"], values: [10, 20], function: 3 (COUNT)
        var args = new[]
        {
            CellValue.FromString("A"),
            CellValue.FromString("A"),
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(3), // COUNT
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(2.0, result.NumericValue);
    }

    [Fact]
    public void GroupBy_Max_ReturnsFirstGroupMax()
    {
        var func = GroupByFunction.Instance;
        // row_fields: ["A", "A"], values: [10, 20], function: 4 (MAX)
        var args = new[]
        {
            CellValue.FromString("A"),
            CellValue.FromString("A"),
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(4), // MAX
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(20.0, result.NumericValue);
    }

    [Fact]
    public void GroupBy_Min_ReturnsFirstGroupMin()
    {
        var func = GroupByFunction.Instance;
        // row_fields: ["A", "A"], values: [10, 20], function: 5 (MIN)
        var args = new[]
        {
            CellValue.FromString("A"),
            CellValue.FromString("A"),
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(5), // MIN
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void GroupBy_InsufficientArguments_ReturnsError()
    {
        var func = GroupByFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("A"),
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void GroupBy_InvalidFunctionType_ReturnsError()
    {
        var func = GroupByFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("A"),
            CellValue.FromNumber(10),
            CellValue.FromNumber(99), // Invalid function type
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void GroupBy_ErrorInData_PropagatesError()
    {
        var func = GroupByFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("A"),
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void GroupBy_NonNumericFunction_ReturnsError()
    {
        var func = GroupByFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("A"),
            CellValue.FromNumber(10),
            CellValue.FromString("invalid"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region PIVOTBY Function Tests

    [Fact]
    public void PivotBy_Sum_ReturnsFirstCellSum()
    {
        var func = PivotByFunction.Instance;
        // row_fields: ["A"], col_fields: ["X"], values: [10], function: 1 (SUM)
        var args = new[]
        {
            CellValue.FromString("A"),
            CellValue.FromString("X"),
            CellValue.FromNumber(10),
            CellValue.FromNumber(1), // SUM
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void PivotBy_Average_ReturnsFirstCellAverage()
    {
        var func = PivotByFunction.Instance;
        // row_fields: ["A", "A"], col_fields: ["X", "X"], values: [10, 20], function: 2 (AVERAGE)
        var args = new[]
        {
            CellValue.FromString("A"),
            CellValue.FromString("A"),
            CellValue.FromString("X"),
            CellValue.FromString("X"),
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(2), // AVERAGE
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(15.0, result.NumericValue);
    }

    [Fact]
    public void PivotBy_Count_ReturnsFirstCellCount()
    {
        var func = PivotByFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("A"),
            CellValue.FromString("X"),
            CellValue.FromNumber(10),
            CellValue.FromNumber(3), // COUNT
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void PivotBy_Max_ReturnsFirstCellMax()
    {
        var func = PivotByFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("A"),
            CellValue.FromString("A"),
            CellValue.FromString("X"),
            CellValue.FromString("X"),
            CellValue.FromNumber(10),
            CellValue.FromNumber(30),
            CellValue.FromNumber(4), // MAX
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(30.0, result.NumericValue);
    }

    [Fact]
    public void PivotBy_Min_ReturnsFirstCellMin()
    {
        var func = PivotByFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("A"),
            CellValue.FromString("A"),
            CellValue.FromString("X"),
            CellValue.FromString("X"),
            CellValue.FromNumber(10),
            CellValue.FromNumber(30),
            CellValue.FromNumber(5), // MIN
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void PivotBy_InsufficientArguments_ReturnsError()
    {
        var func = PivotByFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("A"),
            CellValue.FromString("X"),
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void PivotBy_InvalidFunctionType_ReturnsError()
    {
        var func = PivotByFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("A"),
            CellValue.FromString("X"),
            CellValue.FromNumber(10),
            CellValue.FromNumber(99), // Invalid function type
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void PivotBy_ErrorInData_PropagatesError()
    {
        var func = PivotByFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#REF!"),
            CellValue.FromString("X"),
            CellValue.FromNumber(10),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    #endregion

    #region TRIMRANGE Function Tests

    [Fact]
    public void TrimRange_ReturnsFirstNonEmptyValue()
    {
        var func = TrimRangeFunction.Instance;
        var args = new[]
        {
            CellValue.Empty(),
            CellValue.Empty(),
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.Empty(),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void TrimRange_AllEmpty_ReturnsEmpty()
    {
        var func = TrimRangeFunction.Instance;
        var args = new[]
        {
            CellValue.Empty(),
            CellValue.Empty(),
            CellValue.Empty(),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Empty, result.Type);
    }

    [Fact]
    public void TrimRange_WithTextValues_ReturnsFirstNonEmpty()
    {
        var func = TrimRangeFunction.Instance;
        var args = new[]
        {
            CellValue.Empty(),
            CellValue.FromString("Hello"),
            CellValue.FromString("World"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("Hello", result.StringValue);
    }

    [Fact]
    public void TrimRange_FirstValueNonEmpty_ReturnsThatValue()
    {
        var func = TrimRangeFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(42),
            CellValue.Empty(),
            CellValue.Empty(),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(42.0, result.NumericValue);
    }

    [Fact]
    public void TrimRange_WithParameters_ReturnsFirstNonEmpty()
    {
        var func = TrimRangeFunction.Instance;
        var args = new[]
        {
            CellValue.Empty(),
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(0), // rows_to_trim
            CellValue.FromNumber(0), // cols_to_trim
        };

        var result = func.Execute(null!, args);

        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void TrimRange_NoArguments_ReturnsError()
    {
        var func = TrimRangeFunction.Instance;
        var args = Array.Empty<CellValue>();

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void TrimRange_ErrorInArray_PropagatesError()
    {
        var func = TrimRangeFunction.Instance;
        var args = new[]
        {
            CellValue.Empty(),
            CellValue.Error("#N/A"),
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    #endregion

    #region ANCHORARRAY Function Tests

    [Fact]
    public void AnchorArray_ReturnsReferenceValue()
    {
        var func = AnchorArrayFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(42),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(42.0, result.NumericValue);
    }

    [Fact]
    public void AnchorArray_WithTextValue_ReturnsText()
    {
        var func = AnchorArrayFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("Test"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("Test", result.StringValue);
    }

    [Fact]
    public void AnchorArray_WithError_PropagatesError()
    {
        var func = AnchorArrayFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#REF!"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    [Fact]
    public void AnchorArray_NoArguments_ReturnsError()
    {
        var func = AnchorArrayFunction.Instance;
        var args = Array.Empty<CellValue>();

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void AnchorArray_WithBoolean_ReturnsBoolean()
    {
        var func = AnchorArrayFunction.Instance;
        var args = new[]
        {
            CellValue.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void AnchorArray_WithEmpty_ReturnsEmpty()
    {
        var func = AnchorArrayFunction.Instance;
        var args = new[]
        {
            CellValue.Empty,
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Empty, result.Type);
    }

    #endregion
}
