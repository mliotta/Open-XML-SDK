// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for new Excel math functions: CEILING.MATH, CEILING.PRECISE, FLOOR.MATH, FLOOR.PRECISE,
/// ISO.CEILING, COMBINA, FACTDOUBLE, MDETERM, MINVERSE, MMULT, MUNIT, and LOOKUP.
/// </summary>
public class NewMathFunctions2Tests
{
    // CEILING.MATH Tests
    [Fact]
    public void CeilingMath_PositiveNumber_RoundsUp()
    {
        var func = CeilingMathFunction.Instance;
        var args = new[] { CellValue.FromNumber(4.3) };
        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void CeilingMath_NegativeNumberMode0_RoundsTowardZero()
    {
        var func = CeilingMathFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-4.3),
            CellValue.FromNumber(1),
            CellValue.FromNumber(0)
        };
        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(-4.0, result.NumericValue);
    }

    [Fact]
    public void CeilingMath_NegativeNumberMode1_RoundsAwayFromZero()
    {
        var func = CeilingMathFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-4.3),
            CellValue.FromNumber(1),
            CellValue.FromNumber(1)
        };
        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(-5.0, result.NumericValue);
    }

    [Fact]
    public void CeilingMath_WithSignificance_RoundsToMultiple()
    {
        var func = CeilingMathFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(6.7),
            CellValue.FromNumber(3)
        };
        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(9.0, result.NumericValue);
    }

    // CEILING.PRECISE Tests
    [Fact]
    public void CeilingPrecise_PositiveNumber_RoundsUp()
    {
        var func = CeilingPreciseFunction.Instance;
        var args = new[] { CellValue.FromNumber(4.3) };
        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void CeilingPrecise_NegativeNumber_RoundsTowardZero()
    {
        var func = CeilingPreciseFunction.Instance;
        var args = new[] { CellValue.FromNumber(-4.3) };
        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(-4.0, result.NumericValue);
    }

    // FLOOR.MATH Tests
    [Fact]
    public void FloorMath_PositiveNumber_RoundsDown()
    {
        var func = FloorMathFunction.Instance;
        var args = new[] { CellValue.FromNumber(4.8) };
        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(4.0, result.NumericValue);
    }

    [Fact]
    public void FloorMath_NegativeNumberMode0_RoundsAwayFromZero()
    {
        var func = FloorMathFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-4.3),
            CellValue.FromNumber(1),
            CellValue.FromNumber(0)
        };
        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(-5.0, result.NumericValue);
    }

    [Fact]
    public void FloorMath_NegativeNumberMode1_RoundsTowardZero()
    {
        var func = FloorMathFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-4.8),
            CellValue.FromNumber(1),
            CellValue.FromNumber(1)
        };
        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(-4.0, result.NumericValue);
    }

    // FLOOR.PRECISE Tests
    [Fact]
    public void FloorPrecise_PositiveNumber_RoundsDown()
    {
        var func = FloorPreciseFunction.Instance;
        var args = new[] { CellValue.FromNumber(4.8) };
        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(4.0, result.NumericValue);
    }

    [Fact]
    public void FloorPrecise_NegativeNumber_RoundsDown()
    {
        var func = FloorPreciseFunction.Instance;
        var args = new[] { CellValue.FromNumber(-4.3) };
        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(-5.0, result.NumericValue);
    }

    // ISO.CEILING Tests
    [Fact]
    public void IsoCeiling_PositiveNumber_RoundsUp()
    {
        var func = IsoCeilingFunction.Instance;
        var args = new[] { CellValue.FromNumber(4.3) };
        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void IsoCeiling_NegativeNumber_RoundsTowardZero()
    {
        var func = IsoCeilingFunction.Instance;
        var args = new[] { CellValue.FromNumber(-4.3) };
        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(-4.0, result.NumericValue);
    }

    // COMBINA Tests
    [Fact]
    public void Combina_BasicCase_ReturnsCorrectValue()
    {
        var func = CombinaFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(4),
            CellValue.FromNumber(3)
        };
        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // COMBINA(4,3) = COMBIN(4+3-1, 3) = COMBIN(6,3) = 20
        Assert.Equal(20.0, result.NumericValue);
    }

    [Fact]
    public void Combina_ZeroChosen_ReturnsOne()
    {
        var func = CombinaFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(0)
        };
        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void Combina_NegativeNumber_ReturnsNumError()
    {
        var func = CombinaFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-4),
            CellValue.FromNumber(2)
        };
        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    // FACTDOUBLE Tests
    [Fact]
    public void FactDouble_EvenNumber_ReturnsCorrectValue()
    {
        var func = FactDoubleFunction.Instance;
        var args = new[] { CellValue.FromNumber(6) };
        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // 6!! = 6 * 4 * 2 = 48
        Assert.Equal(48.0, result.NumericValue);
    }

    [Fact]
    public void FactDouble_OddNumber_ReturnsCorrectValue()
    {
        var func = FactDoubleFunction.Instance;
        var args = new[] { CellValue.FromNumber(7) };
        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // 7!! = 7 * 5 * 3 * 1 = 105
        Assert.Equal(105.0, result.NumericValue);
    }

    [Fact]
    public void FactDouble_ZeroOrOne_ReturnsOne()
    {
        var func = FactDoubleFunction.Instance;
        var args = new[] { CellValue.FromNumber(0) };
        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue);

        args = new[] { CellValue.FromNumber(1) };
        result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void FactDouble_NegativeNumber_ReturnsNumError()
    {
        var func = FactDoubleFunction.Instance;
        var args = new[] { CellValue.FromNumber(-5) };
        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    // MDETERM Tests
    [Fact]
    public void MDeterm_NonArray_ReturnsValueError()
    {
        var func = MDetermFunction.Instance;
        var args = new[] { CellValue.FromNumber(5) };
        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    // MINVERSE Tests
    [Fact]
    public void MInverse_NonArray_ReturnsValueError()
    {
        var func = MInverseFunction.Instance;
        var args = new[] { CellValue.FromNumber(5) };
        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    // MMULT Tests
    [Fact]
    public void MMult_TwoNumbers_MultipliesThem()
    {
        var func = MMultFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(3),
            CellValue.FromNumber(4)
        };
        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(12.0, result.NumericValue);
    }

    // MUNIT Tests
    [Fact]
    public void MUnit_PositiveDimension_ReturnsOne()
    {
        var func = MUnitFunction.Instance;
        var args = new[] { CellValue.FromNumber(3) };
        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void MUnit_ZeroOrNegativeDimension_ReturnsValueError()
    {
        var func = MUnitFunction.Instance;
        var args = new[] { CellValue.FromNumber(0) };
        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    // LOOKUP Tests
    [Fact]
    public void Lookup_WrongArgCount_ReturnsValueError()
    {
        var func = LookupFunction.Instance;
        var args = new[] { CellValue.FromNumber(5) };
        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Lookup_ErrorInArgs_PropagatesError()
    {
        var func = LookupFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(1)
        };
        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }
}
