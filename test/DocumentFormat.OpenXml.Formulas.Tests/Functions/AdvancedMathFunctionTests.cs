// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for advanced mathematical functions.
/// </summary>
public class AdvancedMathFunctionTests
{
    #region SUMSQ Tests

    [Fact]
    public void SumSq_TwoNumbers_ReturnsCorrectValue()
    {
        var func = SumSqFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(3),
            CellValue.FromNumber(4),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(25.0, result.NumericValue); // 3² + 4² = 9 + 16 = 25
    }

    [Fact]
    public void SumSq_MultipleNumbers_ReturnsCorrectValue()
    {
        var func = SumSqFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(2),
            CellValue.FromNumber(3),
            CellValue.FromNumber(4),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(30.0, result.NumericValue); // 1 + 4 + 9 + 16 = 30
    }

    [Fact]
    public void SumSq_NegativeNumbers_ReturnsCorrectValue()
    {
        var func = SumSqFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-3),
            CellValue.FromNumber(-4),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(25.0, result.NumericValue); // (-3)² + (-4)² = 9 + 16 = 25
    }

    [Fact]
    public void SumSq_NoArguments_ReturnsValueError()
    {
        var func = SumSqFunction.Instance;
        var args = Array.Empty<CellValue>();

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void SumSq_ErrorValue_PropagatesError()
    {
        var func = SumSqFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(3),
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(4),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void SumSq_IgnoresNonNumericValues()
    {
        var func = SumSqFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(3),
            CellValue.FromString("text"),
            CellValue.FromNumber(4),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(25.0, result.NumericValue); // Only 3² + 4² = 25
    }

    #endregion

    #region SUMX2MY2 Tests

    [Fact]
    public void SumX2MY2_StandardCase_ReturnsCorrectValue()
    {
        var func = SumX2MY2Function.Instance;
        var args = new[]
        {
            CellValue.FromNumber(2),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(-5.0, result.NumericValue); // 2² - 3² = 4 - 9 = -5
    }

    [Fact]
    public void SumX2MY2_ExampleCase_ReturnsCorrectValue()
    {
        var func = SumX2MY2Function.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(4),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(-15.0, result.NumericValue); // 1² - 4² = 1 - 16 = -15
    }

    [Fact]
    public void SumX2MY2_NotTwoArguments_ReturnsValueError()
    {
        var func = SumX2MY2Function.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void SumX2MY2_NonNumericArgument_ReturnsValueError()
    {
        var func = SumX2MY2Function.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromString("text"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void SumX2MY2_ErrorValue_PropagatesError()
    {
        var func = SumX2MY2Function.Instance;
        var args = new[]
        {
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    #endregion

    #region SUMX2PY2 Tests

    [Fact]
    public void SumX2PY2_StandardCase_ReturnsCorrectValue()
    {
        var func = SumX2PY2Function.Instance;
        var args = new[]
        {
            CellValue.FromNumber(2),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(13.0, result.NumericValue); // 2² + 3² = 4 + 9 = 13
    }

    [Fact]
    public void SumX2PY2_ExampleCase_ReturnsCorrectValue()
    {
        var func = SumX2PY2Function.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(4),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(17.0, result.NumericValue); // 1² + 4² = 1 + 16 = 17
    }

    [Fact]
    public void SumX2PY2_NotTwoArguments_ReturnsValueError()
    {
        var func = SumX2PY2Function.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(2),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void SumX2PY2_NonNumericArgument_ReturnsValueError()
    {
        var func = SumX2PY2Function.Instance;
        var args = new[]
        {
            CellValue.FromString("text"),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region SUMXMY2 Tests

    [Fact]
    public void SumXMY2_StandardCase_ReturnsCorrectValue()
    {
        var func = SumXMY2Function.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(4),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(9.0, result.NumericValue); // (1-4)² = (-3)² = 9
    }

    [Fact]
    public void SumXMY2_ExampleCase_ReturnsCorrectValue()
    {
        var func = SumXMY2Function.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(9.0, result.NumericValue); // (5-2)² = 3² = 9
    }

    [Fact]
    public void SumXMY2_ZeroDifference_ReturnsZero()
    {
        var func = SumXMY2Function.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.0, result.NumericValue); // (5-5)² = 0
    }

    [Fact]
    public void SumXMY2_NotTwoArguments_ReturnsValueError()
    {
        var func = SumXMY2Function.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void SumXMY2_NonNumericArgument_ReturnsValueError()
    {
        var func = SumXMY2Function.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region MULTINOMIAL Tests

    [Fact]
    public void Multinomial_StandardCase_ReturnsCorrectValue()
    {
        var func = MultinomialFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(2),
            CellValue.FromNumber(3),
            CellValue.FromNumber(4),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1260.0, result.NumericValue); // 9!/(2!*3!*4!) = 1260
    }

    [Fact]
    public void Multinomial_TwoNumbers_ReturnsCorrectValue()
    {
        var func = MultinomialFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(3),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(10.0, result.NumericValue); // 5!/(3!*2!) = 10
    }

    [Fact]
    public void Multinomial_SingleNumber_ReturnsOne()
    {
        var func = MultinomialFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue); // 5!/5! = 1
    }

    [Fact]
    public void Multinomial_WithZero_ReturnsCorrectValue()
    {
        var func = MultinomialFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(3),
            CellValue.FromNumber(0),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(10.0, result.NumericValue); // 5!/(3!*0!*2!) = 10
    }

    [Fact]
    public void Multinomial_TruncatesDecimals()
    {
        var func = MultinomialFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(3.9),
            CellValue.FromNumber(2.1),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(10.0, result.NumericValue); // Truncates to (3, 2)
    }

    [Fact]
    public void Multinomial_NegativeNumber_ReturnsNumError()
    {
        var func = MultinomialFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-1),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Multinomial_NoArguments_ReturnsValueError()
    {
        var func = MultinomialFunction.Instance;
        var args = Array.Empty<CellValue>();

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Multinomial_NonNumericArgument_ReturnsValueError()
    {
        var func = MultinomialFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(2),
            CellValue.FromString("text"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Multinomial_ErrorValue_PropagatesError()
    {
        var func = MultinomialFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(2),
            CellValue.Error("#REF!"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    #endregion

    #region SERIESSUM Tests

    [Fact]
    public void SeriesSum_StandardCase_ReturnsCorrectValue()
    {
        var func = SeriesSumFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(2),  // x
            CellValue.FromNumber(0),  // n
            CellValue.FromNumber(1),  // m
            CellValue.FromNumber(1),  // coefficient
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue); // 1 * 2^0 = 1
    }

    [Fact]
    public void SeriesSum_WithPower_ReturnsCorrectValue()
    {
        var func = SeriesSumFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(2),  // x
            CellValue.FromNumber(2),  // n
            CellValue.FromNumber(1),  // m
            CellValue.FromNumber(3),  // coefficient
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(12.0, result.NumericValue); // 3 * 2^2 = 12
    }

    [Fact]
    public void SeriesSum_NegativePower_ReturnsCorrectValue()
    {
        var func = SeriesSumFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(2),   // x
            CellValue.FromNumber(-1),  // n
            CellValue.FromNumber(1),   // m
            CellValue.FromNumber(4),   // coefficient
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(2.0, result.NumericValue); // 4 * 2^(-1) = 2
    }

    [Fact]
    public void SeriesSum_ZeroBase_WithPositivePower_ReturnsZero()
    {
        var func = SeriesSumFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0),  // x
            CellValue.FromNumber(2),  // n
            CellValue.FromNumber(1),  // m
            CellValue.FromNumber(5),  // coefficient
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.0, result.NumericValue); // 5 * 0^2 = 0
    }

    [Fact]
    public void SeriesSum_ZeroBase_WithNegativePower_ReturnsNumError()
    {
        var func = SeriesSumFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0),   // x
            CellValue.FromNumber(-1),  // n
            CellValue.FromNumber(1),   // m
            CellValue.FromNumber(5),   // coefficient
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue); // Division by zero
    }

    [Fact]
    public void SeriesSum_NotFourArguments_ReturnsValueError()
    {
        var func = SeriesSumFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(2),
            CellValue.FromNumber(0),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void SeriesSum_NonNumericArgument_ReturnsValueError()
    {
        var func = SeriesSumFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("text"),
            CellValue.FromNumber(0),
            CellValue.FromNumber(1),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void SeriesSum_ErrorValue_PropagatesError()
    {
        var func = SeriesSumFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(2),
            CellValue.Error("#N/A"),
            CellValue.FromNumber(1),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    #endregion

    #region SQRTPI Tests

    [Fact]
    public void SqrtPi_StandardCase_ReturnsCorrectValue()
    {
        var func = SqrtPiFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(Math.Sqrt(Math.PI), result.NumericValue, 10); // sqrt(1*pi)
    }

    [Fact]
    public void SqrtPi_WithTwo_ReturnsCorrectValue()
    {
        var func = SqrtPiFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(Math.Sqrt(2 * Math.PI), result.NumericValue, 10); // sqrt(2*pi)
    }

    [Fact]
    public void SqrtPi_WithZero_ReturnsZero()
    {
        var func = SqrtPiFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void SqrtPi_WithDecimal_ReturnsCorrectValue()
    {
        var func = SqrtPiFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(3.5),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(Math.Sqrt(3.5 * Math.PI), result.NumericValue, 10);
    }

    [Fact]
    public void SqrtPi_NegativeNumber_ReturnsNumError()
    {
        var func = SqrtPiFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void SqrtPi_NoArguments_ReturnsValueError()
    {
        var func = SqrtPiFunction.Instance;
        var args = Array.Empty<CellValue>();

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void SqrtPi_TooManyArguments_ReturnsValueError()
    {
        var func = SqrtPiFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void SqrtPi_NonNumericArgument_ReturnsValueError()
    {
        var func = SqrtPiFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("text"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void SqrtPi_ErrorValue_PropagatesError()
    {
        var func = SqrtPiFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#DIV/0!"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    #endregion
}
