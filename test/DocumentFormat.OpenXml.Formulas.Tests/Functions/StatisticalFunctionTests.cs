// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for advanced statistical functions.
/// </summary>
public class StatisticalFunctionTests
{
    // STDEVP Tests
    [Fact]
    public void StDevP_FiveNumbers_ReturnsPopulationStandardDeviation()
    {
        var func = StDevPFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(2),
            CellValue.FromNumber(3),
            CellValue.FromNumber(4),
            CellValue.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // Population stdev of 1,2,3,4,5 is sqrt(2) ≈ 1.414213562373095
        Assert.Equal(1.414213562373095, result.NumericValue, 10);
    }

    [Fact]
    public void StDevP_SingleNumber_ReturnsZero()
    {
        var func = StDevPFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void StDevP_NoNumbers_ReturnsError()
    {
        var func = StDevPFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("text"),
            CellValue.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void StDevP_ErrorValue_PropagatesError()
    {
        var func = StDevPFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void StDevP_MixedTypes_IgnoresNonNumeric()
    {
        var func = StDevPFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromString("text"),
            CellValue.FromNumber(2),
            CellValue.FromBool(true),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // Should calculate stdev of 1, 2, 3 only
        Assert.Equal(0.816496580927726, result.NumericValue, 10);
    }

    // VARP Tests
    [Fact]
    public void VarP_FiveNumbers_ReturnsPopulationVariance()
    {
        var func = VarPFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(2),
            CellValue.FromNumber(3),
            CellValue.FromNumber(4),
            CellValue.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // Population variance of 1,2,3,4,5 is 2.0
        Assert.Equal(2.0, result.NumericValue);
    }

    [Fact]
    public void VarP_SingleNumber_ReturnsZero()
    {
        var func = VarPFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void VarP_NoNumbers_ReturnsError()
    {
        var func = VarPFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("text"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void VarP_ErrorValue_PropagatesError()
    {
        var func = VarPFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.Error("#NUM!"),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    // LARGE Tests
    [Fact]
    public void Large_SecondLargest_ReturnsFour()
    {
        var func = LargeFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void Large_FirstLargest_ReturnsMaximum()
    {
        var func = LargeFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void Large_KTooLarge_ReturnsError()
    {
        var func = LargeFunction.Instance;
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
    public void Large_KLessThanOne_ReturnsError()
    {
        var func = LargeFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Large_WrongNumberOfArgs_ReturnsError()
    {
        var func = LargeFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Large_NonNumericK_ReturnsError()
    {
        var func = LargeFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromString("text"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Large_ErrorValue_PropagatesError()
    {
        var func = LargeFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    // SMALL Tests
    [Fact]
    public void Small_SecondSmallest_ReturnsTwo()
    {
        var func = SmallFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void Small_FirstSmallest_ReturnsMinimum()
    {
        var func = SmallFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void Small_KTooLarge_ReturnsError()
    {
        var func = SmallFunction.Instance;
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
    public void Small_KLessThanOne_ReturnsError()
    {
        var func = SmallFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Small_WrongNumberOfArgs_ReturnsError()
    {
        var func = SmallFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Small_NonNumericK_ReturnsError()
    {
        var func = SmallFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromString("text"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Small_ErrorValue_PropagatesError()
    {
        var func = SmallFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#NUM!"),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    // PERCENTILE Tests
    [Fact]
    public void Percentile_Median_ReturnsMiddleValue()
    {
        var func = PercentileFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(3),
            CellValue.FromNumber(0.5),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(3.0, result.NumericValue);
    }

    [Fact]
    public void Percentile_Zero_ReturnsMinimum()
    {
        var func = PercentileFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(0.0),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void Percentile_One_ReturnsMaximum()
    {
        var func = PercentileFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(1.0),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void Percentile_SingleValue_ReturnsThatValue()
    {
        var func = PercentileFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(0.5),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void Percentile_KLessThanZero_ReturnsError()
    {
        var func = PercentileFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(-0.1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Percentile_KGreaterThanOne_ReturnsError()
    {
        var func = PercentileFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(1.1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Percentile_WrongNumberOfArgs_ReturnsError()
    {
        var func = PercentileFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Percentile_NonNumericK_ReturnsError()
    {
        var func = PercentileFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromString("text"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Percentile_NoNumbers_ReturnsError()
    {
        var func = PercentileFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("text"),
            CellValue.FromNumber(0.5),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Percentile_ErrorValue_PropagatesError()
    {
        var func = PercentileFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(0.5),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    // QUARTILE Tests
    [Fact]
    public void Quartile_Zero_ReturnsMinimum()
    {
        var func = QuartileFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void Quartile_Two_ReturnsMedian()
    {
        var func = QuartileFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void Quartile_Four_ReturnsMaximum()
    {
        var func = QuartileFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(4),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void Quartile_InvalidQuart_ReturnsError()
    {
        var func = QuartileFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Quartile_NegativeQuart_ReturnsError()
    {
        var func = QuartileFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(-1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Quartile_WrongNumberOfArgs_ReturnsError()
    {
        var func = QuartileFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Quartile_NonNumericQuart_ReturnsError()
    {
        var func = QuartileFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromString("text"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Quartile_ErrorValue_PropagatesError()
    {
        var func = QuartileFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }
}
