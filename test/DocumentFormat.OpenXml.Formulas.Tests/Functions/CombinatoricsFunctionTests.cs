// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for combinatorics and advanced rounding functions.
/// </summary>
public class CombinatoricsFunctionTests
{
    #region COMBIN Tests

    [Fact]
    public void Combin_StandardCase_ReturnsCorrectValue()
    {
        var func = CombinFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(8),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(28.0, result.NumericValue); // 8!/(2!*6!) = 28
    }

    [Fact]
    public void Combin_ChooseZero_ReturnsOne()
    {
        var func = CombinFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue); // C(n, 0) = 1
    }

    [Fact]
    public void Combin_ChooseAll_ReturnsOne()
    {
        var func = CombinFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue); // C(n, n) = 1
    }

    [Fact]
    public void Combin_LargeNumbers_ReturnsCorrectValue()
    {
        var func = CombinFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(120.0, result.NumericValue); // 10!/(3!*7!) = 120
    }

    [Fact]
    public void Combin_TruncatesDecimals()
    {
        var func = CombinFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(8.9),
            CellValue.FromNumber(2.1),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(28.0, result.NumericValue); // Truncates to COMBIN(8, 2)
    }

    [Fact]
    public void Combin_KGreaterThanN_ReturnsNumError()
    {
        var func = CombinFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Combin_NegativeN_ReturnsNumError()
    {
        var func = CombinFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-5),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Combin_NegativeK_ReturnsNumError()
    {
        var func = CombinFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(-2),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Combin_WrongArgumentCount_ReturnsValueError()
    {
        var func = CombinFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Combin_NonNumericArgument_ReturnsValueError()
    {
        var func = CombinFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("text"),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Combin_ErrorValue_PropagatesError()
    {
        var func = CombinFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    #endregion

    #region PERMUT Tests

    [Fact]
    public void Permut_StandardCase_ReturnsCorrectValue()
    {
        var func = PermutFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(8),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(56.0, result.NumericValue); // 8!/(8-2)! = 8*7 = 56
    }

    [Fact]
    public void Permut_ChooseZero_ReturnsOne()
    {
        var func = PermutFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue); // P(n, 0) = 1
    }

    [Fact]
    public void Permut_ChooseAll_ReturnsFactorial()
    {
        var func = PermutFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(120.0, result.NumericValue); // 5! = 120
    }

    [Fact]
    public void Permut_LargeNumbers_ReturnsCorrectValue()
    {
        var func = PermutFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(720.0, result.NumericValue); // 10*9*8 = 720
    }

    [Fact]
    public void Permut_TruncatesDecimals()
    {
        var func = PermutFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(8.9),
            CellValue.FromNumber(2.1),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(56.0, result.NumericValue); // Truncates to PERMUT(8, 2)
    }

    [Fact]
    public void Permut_KGreaterThanN_ReturnsNumError()
    {
        var func = PermutFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Permut_NegativeN_ReturnsNumError()
    {
        var func = PermutFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-5),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Permut_NegativeK_ReturnsNumError()
    {
        var func = PermutFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(-2),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Permut_WrongArgumentCount_ReturnsValueError()
    {
        var func = PermutFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Permut_NonNumericArgument_ReturnsValueError()
    {
        var func = PermutFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("text"),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Permut_ErrorValue_PropagatesError()
    {
        var func = PermutFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    #endregion

    #region MROUND Tests

    [Fact]
    public void Mround_StandardCase_ReturnsNearestMultiple()
    {
        var func = MroundFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(9.0, result.NumericValue); // 10 rounded to nearest multiple of 3 is 9
    }

    [Fact]
    public void Mround_RoundUp_ReturnsCorrectValue()
    {
        var func = MroundFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(11),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(12.0, result.NumericValue); // 11 rounded to nearest multiple of 3 is 12
    }

    [Fact]
    public void Mround_ExactMultiple_ReturnsSameValue()
    {
        var func = MroundFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(15),
            CellValue.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(15.0, result.NumericValue);
    }

    [Fact]
    public void Mround_MidpointRoundsAwayFromZero()
    {
        var func = MroundFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(7.5),
            CellValue.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(10.0, result.NumericValue); // 7.5 rounds to 10 (away from zero)
    }

    [Fact]
    public void Mround_NegativeNumberAndMultiple_ReturnsCorrectValue()
    {
        var func = MroundFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-10),
            CellValue.FromNumber(-3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(-9.0, result.NumericValue);
    }

    [Fact]
    public void Mround_DecimalMultiple_ReturnsCorrectValue()
    {
        var func = MroundFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1.3),
            CellValue.FromNumber(0.2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.4, result.NumericValue, 10);
    }

    [Fact]
    public void Mround_ZeroMultiple_ReturnsNumError()
    {
        var func = MroundFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Mround_DifferentSigns_ReturnsNumError()
    {
        var func = MroundFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(-3),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Mround_WrongArgumentCount_ReturnsValueError()
    {
        var func = MroundFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Mround_NonNumericArgument_ReturnsValueError()
    {
        var func = MroundFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("text"),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Mround_ErrorValue_PropagatesError()
    {
        var func = MroundFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    #endregion

    #region QUOTIENT Tests

    [Fact]
    public void Quotient_StandardCase_ReturnsIntegerPart()
    {
        var func = QuotientFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(3.0, result.NumericValue); // 10/3 = 3.333... -> 3
    }

    [Fact]
    public void Quotient_ExactDivision_ReturnsQuotient()
    {
        var func = QuotientFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(15),
            CellValue.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(3.0, result.NumericValue);
    }

    [Fact]
    public void Quotient_NegativeNumerator_TruncatesTowardZero()
    {
        var func = QuotientFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-10),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(-3.0, result.NumericValue); // -10/3 = -3.333... -> -3 (toward zero)
    }

    [Fact]
    public void Quotient_NegativeDenominator_ReturnsCorrectValue()
    {
        var func = QuotientFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(-3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(-3.0, result.NumericValue);
    }

    [Fact]
    public void Quotient_BothNegative_ReturnsPositive()
    {
        var func = QuotientFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-10),
            CellValue.FromNumber(-3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(3.0, result.NumericValue);
    }

    [Fact]
    public void Quotient_DecimalValues_TruncatesCorrectly()
    {
        var func = QuotientFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10.5),
            CellValue.FromNumber(2.3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(4.0, result.NumericValue); // 10.5/2.3 = 4.565... -> 4
    }

    [Fact]
    public void Quotient_ZeroNumerator_ReturnsZero()
    {
        var func = QuotientFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0),
            CellValue.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void Quotient_ZeroDenominator_ReturnsDivZeroError()
    {
        var func = QuotientFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void Quotient_WrongArgumentCount_ReturnsValueError()
    {
        var func = QuotientFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Quotient_NonNumericArgument_ReturnsValueError()
    {
        var func = QuotientFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("text"),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Quotient_ErrorValue_PropagatesError()
    {
        var func = QuotientFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    #endregion
}
