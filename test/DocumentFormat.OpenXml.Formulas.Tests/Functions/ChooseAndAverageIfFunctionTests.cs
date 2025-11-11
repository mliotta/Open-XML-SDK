// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for CHOOSE and AVERAGEIF functions.
/// </summary>
public class ChooseAndAverageIfFunctionTests
{
    #region CHOOSE Function Tests

    [Fact]
    public void Choose_ValidIndex_ReturnsCorrectValue()
    {
        var func = ChooseFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(2),
            CellValue.FromString("Red"),
            CellValue.FromString("Green"),
            CellValue.FromString("Blue"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("Green", result.StringValue);
    }

    [Fact]
    public void Choose_FirstIndex_ReturnsFirstValue()
    {
        var func = ChooseFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromString("Monday"),
            CellValue.FromString("Tuesday"),
            CellValue.FromString("Wednesday"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("Monday", result.StringValue);
    }

    [Fact]
    public void Choose_LastIndex_ReturnsLastValue()
    {
        var func = ChooseFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(3),
            CellValue.FromString("Red"),
            CellValue.FromString("Green"),
            CellValue.FromString("Blue"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("Blue", result.StringValue);
    }

    [Fact]
    public void Choose_NumericValues_ReturnsNumber()
    {
        var func = ChooseFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(2),
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(20.0, result.NumericValue);
    }

    [Fact]
    public void Choose_MixedTypes_ReturnsCorrectType()
    {
        var func = ChooseFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(3),
            CellValue.FromNumber(100),
            CellValue.FromString("Text"),
            CellValue.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void Choose_IndexOutOfRangeLow_ReturnsError()
    {
        var func = ChooseFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0),
            CellValue.FromString("Red"),
            CellValue.FromString("Green"),
            CellValue.FromString("Blue"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Choose_IndexOutOfRangeHigh_ReturnsError()
    {
        var func = ChooseFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(4),
            CellValue.FromString("Red"),
            CellValue.FromString("Green"),
            CellValue.FromString("Blue"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Choose_NegativeIndex_ReturnsError()
    {
        var func = ChooseFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-1),
            CellValue.FromString("Red"),
            CellValue.FromString("Green"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Choose_NonNumericIndex_ReturnsError()
    {
        var func = ChooseFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("text"),
            CellValue.FromString("Red"),
            CellValue.FromString("Green"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Choose_InsufficientArguments_ReturnsError()
    {
        var func = ChooseFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Choose_ErrorInIndex_PropagatesError()
    {
        var func = ChooseFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#DIV/0!"),
            CellValue.FromString("Red"),
            CellValue.FromString("Green"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void Choose_DecimalIndex_TruncatesToInteger()
    {
        var func = ChooseFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(2.9),
            CellValue.FromString("Red"),
            CellValue.FromString("Green"),
            CellValue.FromString("Blue"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("Green", result.StringValue);
    }

    #endregion

    #region AVERAGEIF Function Tests

    [Fact]
    public void AverageIf_GreaterThan_ReturnsAverage()
    {
        var func = AverageIfFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromString(">5"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void AverageIf_LessThan_ReturnsAverage()
    {
        var func = AverageIfFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(3),
            CellValue.FromString("<5"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(3.0, result.NumericValue);
    }

    [Fact]
    public void AverageIf_GreaterThanOrEqual_ReturnsAverage()
    {
        var func = AverageIfFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromString(">=5"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void AverageIf_LessThanOrEqual_ReturnsAverage()
    {
        var func = AverageIfFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromString("<=5"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void AverageIf_Equality_ReturnsAverage()
    {
        var func = AverageIfFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromString("=10"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void AverageIf_NotEqual_ReturnsAverage()
    {
        var func = AverageIfFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromString("<>5"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void AverageIf_NoMatches_ReturnsError()
    {
        var func = AverageIfFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(3),
            CellValue.FromString(">10"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void AverageIf_TextCriteria_MatchesText()
    {
        var func = AverageIfFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("Apple"),
            CellValue.FromString("Apple"),
        };

        var result = func.Execute(null!, args);

        // Text values don't contribute to average, so DIV/0
        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void AverageIf_NumericCriteria_WithoutOperator()
    {
        var func = AverageIfFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void AverageIf_ErrorInRange_PropagatesError()
    {
        var func = AverageIfFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#REF!"),
            CellValue.FromString(">5"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    [Fact]
    public void AverageIf_ErrorInCriteria_PropagatesError()
    {
        var func = AverageIfFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.Error("#VALUE!"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void AverageIf_InsufficientArguments_ReturnsError()
    {
        var func = AverageIfFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void AverageIf_TooManyArguments_ReturnsError()
    {
        var func = AverageIfFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromString(">5"),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void AverageIf_BooleanCriteria_Matches()
    {
        var func = AverageIfFunction.Instance;
        var args = new[]
        {
            CellValue.FromBool(true),
            CellValue.FromBool(true),
        };

        var result = func.Execute(null!, args);

        // Boolean values don't contribute to numeric average
        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    #endregion
}
