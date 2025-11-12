// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;
using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for PERCENTOF function.
/// </summary>
public class PercentOfFunctionTests
{
    [Fact]
    public void PercentOf_BasicCalculation_ReturnsCorrectValue()
    {
        var func = PercentOfFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(25),
            CellValue.FromNumber(100),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.25, result.NumericValue);
    }

    [Fact]
    public void PercentOf_FiftyPercent_ReturnsHalf()
    {
        var func = PercentOfFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(50),
            CellValue.FromNumber(100),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(0.5, result.NumericValue);
    }

    [Fact]
    public void PercentOf_GreaterThanTotal_ReturnsValueGreaterThanOne()
    {
        var func = PercentOfFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(150),
            CellValue.FromNumber(100),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(1.5, result.NumericValue);
    }

    [Fact]
    public void PercentOf_NegativeValues_ReturnsCorrectValue()
    {
        var func = PercentOfFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-25),
            CellValue.FromNumber(100),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(-0.25, result.NumericValue);
    }

    [Fact]
    public void PercentOf_DivideByZero_ReturnsError()
    {
        var func = PercentOfFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(25),
            CellValue.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void PercentOf_WrongNumberOfArgs_ReturnsError()
    {
        var func = PercentOfFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(25),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void PercentOf_NonNumericArgument_ReturnsError()
    {
        var func = PercentOfFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("abc"),
            CellValue.FromNumber(100),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void PercentOf_ErrorPropagation_ReturnsError()
    {
        var func = PercentOfFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(100),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void PercentOf_DecimalValues_ReturnsCorrectValue()
    {
        var func = PercentOfFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(33.33),
            CellValue.FromNumber(100),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.3333, result.NumericValue, 4);
    }

    [Fact]
    public void PercentOf_ZeroSubset_ReturnsZero()
    {
        var func = PercentOfFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0),
            CellValue.FromNumber(100),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(0.0, result.NumericValue);
    }
}
