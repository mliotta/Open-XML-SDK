// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for conditional aggregation functions (SUMIF, COUNTIF).
/// </summary>
public class ConditionalFunctionTests
{
    #region COUNTIF Tests

    [Fact]
    public void CountIf_ExactTextMatch_ReturnsCount()
    {
        var func = CountIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Apple"),
            FormulaResult.FromString("Apple"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void CountIf_ExactTextMismatch_ReturnsZero()
    {
        var func = CountIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Apple"),
            FormulaResult.FromString("Orange"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void CountIf_TextMatchCaseInsensitive_ReturnsCount()
    {
        var func = CountIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("APPLE"),
            FormulaResult.FromString("apple"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void CountIf_NumericEquality_ReturnsCount()
    {
        var func = CountIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(5),
            FormulaResult.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void CountIf_GreaterThan_ReturnsCount()
    {
        var func = CountIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(10),
            FormulaResult.FromString(">5"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void CountIf_GreaterThanFails_ReturnsZero()
    {
        var func = CountIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(3),
            FormulaResult.FromString(">5"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void CountIf_LessThan_ReturnsCount()
    {
        var func = CountIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(3),
            FormulaResult.FromString("<5"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void CountIf_GreaterThanOrEqual_ReturnsCount()
    {
        var func = CountIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(5),
            FormulaResult.FromString(">=5"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void CountIf_LessThanOrEqual_ReturnsCount()
    {
        var func = CountIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(5),
            FormulaResult.FromString("<=5"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void CountIf_NotEqual_ReturnsCount()
    {
        var func = CountIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(10),
            FormulaResult.FromString("<>5"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void CountIf_NotEqualSameValue_ReturnsZero()
    {
        var func = CountIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(5),
            FormulaResult.FromString("<>5"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void CountIf_EqualOperator_ReturnsCount()
    {
        var func = CountIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(5),
            FormulaResult.FromString("=5"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void CountIf_EqualOperatorText_ReturnsCount()
    {
        var func = CountIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Apple"),
            FormulaResult.FromString("=Apple"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void CountIf_BooleanMatch_ReturnsCount()
    {
        var func = CountIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromBool(true),
            FormulaResult.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void CountIf_WrongArgumentCount_ReturnsError()
    {
        var func = CountIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void CountIf_ErrorValue_PropagatesError()
    {
        var func = CountIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#DIV/0!"),
            FormulaResult.FromString(">5"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    #endregion

    #region SUMIF Tests

    [Fact]
    public void SumIf_NumericMatch_ReturnsSum()
    {
        var func = SumIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(5),
            FormulaResult.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void SumIf_NumericMismatch_ReturnsZero()
    {
        var func = SumIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(3),
            FormulaResult.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void SumIf_GreaterThan_ReturnsSum()
    {
        var func = SumIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(10),
            FormulaResult.FromString(">5"),
            FormulaResult.FromNumber(100),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(100.0, result.NumericValue);
    }

    [Fact]
    public void SumIf_GreaterThanFails_ReturnsZero()
    {
        var func = SumIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(3),
            FormulaResult.FromString(">5"),
            FormulaResult.FromNumber(100),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void SumIf_LessThan_ReturnsSum()
    {
        var func = SumIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(3),
            FormulaResult.FromString("<5"),
            FormulaResult.FromNumber(50),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(50.0, result.NumericValue);
    }

    [Fact]
    public void SumIf_GreaterThanOrEqual_ReturnsSum()
    {
        var func = SumIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(5),
            FormulaResult.FromString(">=5"),
            FormulaResult.FromNumber(75),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(75.0, result.NumericValue);
    }

    [Fact]
    public void SumIf_LessThanOrEqual_ReturnsSum()
    {
        var func = SumIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(5),
            FormulaResult.FromString("<=5"),
            FormulaResult.FromNumber(25),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(25.0, result.NumericValue);
    }

    [Fact]
    public void SumIf_NotEqual_ReturnsSum()
    {
        var func = SumIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(10),
            FormulaResult.FromString("<>5"),
            FormulaResult.FromNumber(200),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(200.0, result.NumericValue);
    }

    [Fact]
    public void SumIf_EqualOperator_ReturnsSum()
    {
        var func = SumIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(5),
            FormulaResult.FromString("=5"),
            FormulaResult.FromNumber(150),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(150.0, result.NumericValue);
    }

    [Fact]
    public void SumIf_TextMatch_ReturnsSum()
    {
        var func = SumIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Apple"),
            FormulaResult.FromString("Apple"),
            FormulaResult.FromNumber(50),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(50.0, result.NumericValue);
    }

    [Fact]
    public void SumIf_TextMismatch_ReturnsZero()
    {
        var func = SumIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Apple"),
            FormulaResult.FromString("Orange"),
            FormulaResult.FromNumber(50),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void SumIf_TwoArguments_SumsCriteriaRange()
    {
        var func = SumIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(10),
            FormulaResult.FromString(">5"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void SumIf_NonNumericSumRange_ReturnsZero()
    {
        var func = SumIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(10),
            FormulaResult.FromString(">5"),
            FormulaResult.FromString("text"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void SumIf_WrongArgumentCount_ReturnsError()
    {
        var func = SumIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void SumIf_TooManyArguments_ReturnsError()
    {
        var func = SumIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(5),
            FormulaResult.FromString(">3"),
            FormulaResult.FromNumber(10),
            FormulaResult.FromNumber(20),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void SumIf_ErrorValue_PropagatesError()
    {
        var func = SumIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#DIV/0!"),
            FormulaResult.FromString(">5"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void SumIf_ErrorInCriteria_PropagatesError()
    {
        var func = SumIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(10),
            FormulaResult.Error("#REF!"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    [Fact]
    public void SumIf_ErrorInSumRange_PropagatesError()
    {
        var func = SumIfFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(10),
            FormulaResult.FromString(">5"),
            FormulaResult.Error("#N/A"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    #endregion
}
