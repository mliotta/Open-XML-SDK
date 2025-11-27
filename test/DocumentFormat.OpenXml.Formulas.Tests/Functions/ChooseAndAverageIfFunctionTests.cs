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
            FormulaResult.FromNumber(2),
            FormulaResult.FromString("Red"),
            FormulaResult.FromString("Green"),
            FormulaResult.FromString("Blue"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("Green", result.StringValue);
    }

    [Fact]
    public void Choose_FirstIndex_ReturnsFirstValue()
    {
        var func = ChooseFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(1),
            FormulaResult.FromString("Monday"),
            FormulaResult.FromString("Tuesday"),
            FormulaResult.FromString("Wednesday"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("Monday", result.StringValue);
    }

    [Fact]
    public void Choose_LastIndex_ReturnsLastValue()
    {
        var func = ChooseFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(3),
            FormulaResult.FromString("Red"),
            FormulaResult.FromString("Green"),
            FormulaResult.FromString("Blue"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("Blue", result.StringValue);
    }

    [Fact]
    public void Choose_NumericValues_ReturnsNumber()
    {
        var func = ChooseFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(2),
            FormulaResult.FromNumber(10),
            FormulaResult.FromNumber(20),
            FormulaResult.FromNumber(30),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(20.0, result.NumericValue);
    }

    [Fact]
    public void Choose_MixedTypes_ReturnsCorrectType()
    {
        var func = ChooseFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(3),
            FormulaResult.FromNumber(100),
            FormulaResult.FromString("Text"),
            FormulaResult.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void Choose_IndexOutOfRangeLow_ReturnsError()
    {
        var func = ChooseFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(0),
            FormulaResult.FromString("Red"),
            FormulaResult.FromString("Green"),
            FormulaResult.FromString("Blue"),
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
            FormulaResult.FromNumber(4),
            FormulaResult.FromString("Red"),
            FormulaResult.FromString("Green"),
            FormulaResult.FromString("Blue"),
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
            FormulaResult.FromNumber(-1),
            FormulaResult.FromString("Red"),
            FormulaResult.FromString("Green"),
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
            FormulaResult.FromString("text"),
            FormulaResult.FromString("Red"),
            FormulaResult.FromString("Green"),
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
            FormulaResult.FromNumber(1),
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
            FormulaResult.Error("#DIV/0!"),
            FormulaResult.FromString("Red"),
            FormulaResult.FromString("Green"),
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
            FormulaResult.FromNumber(2.9),
            FormulaResult.FromString("Red"),
            FormulaResult.FromString("Green"),
            FormulaResult.FromString("Blue"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
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
            FormulaResult.FromNumber(10),
            FormulaResult.FromString(">5"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void AverageIf_LessThan_ReturnsAverage()
    {
        var func = AverageIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(3),
            FormulaResult.FromString("<5"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(3.0, result.NumericValue);
    }

    [Fact]
    public void AverageIf_GreaterThanOrEqual_ReturnsAverage()
    {
        var func = AverageIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(5),
            FormulaResult.FromString(">=5"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void AverageIf_LessThanOrEqual_ReturnsAverage()
    {
        var func = AverageIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(5),
            FormulaResult.FromString("<=5"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void AverageIf_Equality_ReturnsAverage()
    {
        var func = AverageIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(10),
            FormulaResult.FromString("=10"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void AverageIf_NotEqual_ReturnsAverage()
    {
        var func = AverageIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(10),
            FormulaResult.FromString("<>5"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void AverageIf_NoMatches_ReturnsError()
    {
        var func = AverageIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(3),
            FormulaResult.FromString(">10"),
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
            FormulaResult.FromString("Apple"),
            FormulaResult.FromString("Apple"),
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
            FormulaResult.FromNumber(10),
            FormulaResult.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void AverageIf_ErrorInRange_PropagatesError()
    {
        var func = AverageIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#REF!"),
            FormulaResult.FromString(">5"),
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
            FormulaResult.FromNumber(10),
            FormulaResult.Error("#VALUE!"),
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
            FormulaResult.FromNumber(10),
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
            FormulaResult.FromNumber(10),
            FormulaResult.FromString(">5"),
            FormulaResult.FromNumber(20),
            FormulaResult.FromNumber(30),
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
            FormulaResult.FromBool(true),
            FormulaResult.FromBool(true),
        };

        var result = func.Execute(null!, args);

        // Boolean values don't contribute to numeric average
        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    #endregion
}
