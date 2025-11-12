// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;
using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for BINOM.DIST.RANGE function.
/// </summary>
public class BinomDistRangeFunctionTests
{
    [Fact]
    public void BinomDistRange_SingleValue_ReturnsCorrectProbability()
    {
        var func = BinomDistRangeFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),    // trials
            CellValue.FromNumber(0.5),   // probability
            CellValue.FromNumber(5),     // number_s
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // P(X=5) for n=10, p=0.5 should be approximately 0.246
        Assert.True(result.NumericValue > 0.24 && result.NumericValue < 0.25);
    }

    [Fact]
    public void BinomDistRange_Range_ReturnsCorrectProbability()
    {
        var func = BinomDistRangeFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),    // trials
            CellValue.FromNumber(0.5),   // probability
            CellValue.FromNumber(4),     // number_s
            CellValue.FromNumber(6),     // number_s2
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // P(4 <= X <= 6) for n=10, p=0.5 should be approximately 0.656
        Assert.True(result.NumericValue > 0.65 && result.NumericValue < 0.66);
    }

    [Fact]
    public void BinomDistRange_InvalidRange_ReturnsError()
    {
        var func = BinomDistRangeFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),    // trials
            CellValue.FromNumber(0.5),   // probability
            CellValue.FromNumber(6),     // number_s
            CellValue.FromNumber(4),     // number_s2 (less than number_s)
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void BinomDistRange_NegativeTrials_ReturnsError()
    {
        var func = BinomDistRangeFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-5),    // trials
            CellValue.FromNumber(0.5),   // probability
            CellValue.FromNumber(2),     // number_s
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void BinomDistRange_ProbabilityOutOfRange_ReturnsError()
    {
        var func = BinomDistRangeFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),    // trials
            CellValue.FromNumber(1.5),   // probability (>1)
            CellValue.FromNumber(5),     // number_s
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void BinomDistRange_NumberSGreaterThanTrials_ReturnsError()
    {
        var func = BinomDistRangeFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),    // trials
            CellValue.FromNumber(0.5),   // probability
            CellValue.FromNumber(15),    // number_s (>trials)
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void BinomDistRange_WrongNumberOfArgs_ReturnsError()
    {
        var func = BinomDistRangeFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),    // trials
            CellValue.FromNumber(0.5),   // probability
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void BinomDistRange_ErrorPropagation_ReturnsError()
    {
        var func = BinomDistRangeFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(0.5),
            CellValue.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void BinomDistRange_ZeroProbability_ReturnsCorrectValue()
    {
        var func = BinomDistRangeFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),    // trials
            CellValue.FromNumber(0.0),   // probability
            CellValue.FromNumber(0),     // number_s
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // P(X=0) for n=10, p=0 should be 1.0
        Assert.Equal(1.0, result.NumericValue, 5);
    }

    [Fact]
    public void BinomDistRange_OneProbability_ReturnsCorrectValue()
    {
        var func = BinomDistRangeFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),    // trials
            CellValue.FromNumber(1.0),   // probability
            CellValue.FromNumber(10),    // number_s
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // P(X=10) for n=10, p=1 should be 1.0
        Assert.Equal(1.0, result.NumericValue, 5);
    }
}
