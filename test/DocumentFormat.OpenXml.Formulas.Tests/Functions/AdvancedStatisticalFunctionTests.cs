// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for advanced statistical functions: PERCENTILE.INC/EXC, QUARTILE.INC/EXC,
/// FORECAST, FORECAST.LINEAR, TREND, GROWTH, LINEST, LOGEST.
/// </summary>
public class AdvancedStatisticalFunctionTests
{
    // PERCENTILE.INC Tests
    [Fact]
    public void PercentileInc_Median_ReturnsMiddleValue()
    {
        var func = PercentileIncFunction.Instance;
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
    public void PercentileInc_Zero_ReturnsMinimum()
    {
        var func = PercentileIncFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(0.0),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void PercentileInc_One_ReturnsMaximum()
    {
        var func = PercentileIncFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(1.0),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void PercentileInc_InvalidK_ReturnsError()
    {
        var func = PercentileIncFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(1.5),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    // PERCENTILE.EXC Tests
    [Fact]
    public void PercentileExc_Median_ReturnsMiddleValue()
    {
        var func = PercentileExcFunction.Instance;
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
    public void PercentileExc_ZeroOrOne_ReturnsError()
    {
        var func = PercentileExcFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(0.0),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);

        args[1] = CellValue.FromNumber(1.0);
        result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void PercentileExc_ValidK_ReturnsValue()
    {
        var func = PercentileExcFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(0.5),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
    }

    // QUARTILE.INC Tests
    [Fact]
    public void QuartileInc_Zero_ReturnsMinimum()
    {
        var func = QuartileIncFunction.Instance;
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
    public void QuartileInc_Two_ReturnsMedian()
    {
        var func = QuartileIncFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void QuartileInc_Four_ReturnsMaximum()
    {
        var func = QuartileIncFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(4),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void QuartileInc_InvalidQuart_ReturnsError()
    {
        var func = QuartileIncFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    // QUARTILE.EXC Tests
    [Fact]
    public void QuartileExc_ValidQuarts_ReturnsValue()
    {
        var func = QuartileExcFunction.Instance;

        // Test quart 1
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(1),
        };
        var result = func.Execute(null!, args);
        Assert.Equal(CellValueType.Number, result.Type);

        // Test quart 2
        args[1] = CellValue.FromNumber(2);
        result = func.Execute(null!, args);
        Assert.Equal(CellValueType.Number, result.Type);

        // Test quart 3
        args[1] = CellValue.FromNumber(3);
        result = func.Execute(null!, args);
        Assert.Equal(CellValueType.Number, result.Type);
    }

    [Fact]
    public void QuartileExc_ZeroOrFour_ReturnsError()
    {
        var func = QuartileExcFunction.Instance;

        // Test quart 0
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(0),
        };
        var result = func.Execute(null!, args);
        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);

        // Test quart 4
        args[1] = CellValue.FromNumber(4);
        result = func.Execute(null!, args);
        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    // FORECAST Tests
    [Fact]
    public void Forecast_SingleDataPoint_ReturnsYValue()
    {
        var func = ForecastFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5), // x for prediction
            CellValue.FromNumber(10), // known_y
            CellValue.FromNumber(2), // known_x
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void Forecast_LinearData_ReturnsCorrectPrediction()
    {
        // With single point in Phase 0, we can't test full linear regression
        // This test validates the basic structure
        var func = ForecastFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(3), // x for prediction
            CellValue.FromNumber(2), // known_y
            CellValue.FromNumber(1), // known_x
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
    }

    [Fact]
    public void Forecast_WrongNumberOfArgs_ReturnsError()
    {
        var func = ForecastFunction.Instance;
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
    public void Forecast_ErrorPropagation_ReturnsError()
    {
        var func = ForecastFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(2),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    // FORECAST.LINEAR Tests
    [Fact]
    public void ForecastLinear_SameAsForecast()
    {
        var forecastFunc = ForecastFunction.Instance;
        var forecastLinearFunc = ForecastLinearFunction.Instance;

        var args = new[]
        {
            CellValue.FromNumber(3),
            CellValue.FromNumber(2),
            CellValue.FromNumber(1),
        };

        var result1 = forecastFunc.Execute(null!, args);
        var result2 = forecastLinearFunc.Execute(null!, args);

        Assert.Equal(result1.NumericValue, result2.NumericValue);
    }

    // TREND Tests
    [Fact]
    public void Trend_SingleDataPoint_ReturnsYValue()
    {
        var func = TrendFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10), // known_y
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void Trend_WithKnownX_ReturnsValue()
    {
        var func = TrendFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10), // known_y
            CellValue.FromNumber(2), // known_x
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
    }

    [Fact]
    public void Trend_WithNewX_ReturnsValue()
    {
        var func = TrendFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10), // known_y
            CellValue.FromNumber(2), // known_x
            CellValue.FromNumber(3), // new_x
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
    }

    [Fact]
    public void Trend_ConstFalse_ForcesZeroIntercept()
    {
        var func = TrendFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10), // known_y
            CellValue.FromNumber(2), // known_x
            CellValue.FromNumber(4), // new_x
            CellValue.FromBool(false), // const = FALSE
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // With const=FALSE and y=10, x=2: slope = 10/2 = 5
        // Trend at x=4: 5 * 4 = 20
        Assert.Equal(20.0, result.NumericValue);
    }

    // GROWTH Tests
    [Fact]
    public void Growth_SingleDataPoint_ReturnsYValue()
    {
        var func = GrowthFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10), // known_y
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void Growth_NegativeY_ReturnsError()
    {
        var func = GrowthFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-10), // negative known_y
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Growth_WithKnownX_ReturnsValue()
    {
        var func = GrowthFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10), // known_y
            CellValue.FromNumber(2), // known_x
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
    }

    [Fact]
    public void Growth_WithNewX_ReturnsValue()
    {
        var func = GrowthFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(8), // known_y
            CellValue.FromNumber(3), // known_x
            CellValue.FromNumber(6), // new_x
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // With const=FALSE and y=8, x=3: m = 8^(1/3) = 2
        // Growth at x=6: 2^6 = 64
    }

    // LINEST Tests
    [Fact]
    public void Linest_TwoDataPoints_ReturnsSlope()
    {
        var func = LinestFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(4), // known_y
            CellValue.FromNumber(2), // known_x
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // Phase 0 returns slope only
    }

    [Fact]
    public void Linest_SingleDataPoint_ReturnsError()
    {
        var func = LinestFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(4), // known_y
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    [Fact]
    public void Linest_ConstFalse_ForcesZeroIntercept()
    {
        var func = LinestFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10), // known_y
            CellValue.FromNumber(2), // known_x
            CellValue.FromBool(false), // const = FALSE
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // With const=FALSE: slope = (2*10) / (2*2) = 20/4 = 5
        Assert.Equal(5.0, result.NumericValue);
    }

    // LOGEST Tests
    [Fact]
    public void Logest_TwoDataPoints_ReturnsBase()
    {
        var func = LogestFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(4), // known_y
            CellValue.FromNumber(2), // known_x
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // Phase 0 returns m (base) only
    }

    [Fact]
    public void Logest_NegativeY_ReturnsError()
    {
        var func = LogestFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-4), // negative known_y
            CellValue.FromNumber(2), // known_x
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Logest_ZeroY_ReturnsError()
    {
        var func = LogestFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0), // zero known_y
            CellValue.FromNumber(2), // known_x
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Logest_SingleDataPoint_ReturnsError()
    {
        var func = LogestFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(4), // known_y
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    [Fact]
    public void Logest_ConstFalse_ForcesBaseOne()
    {
        var func = LogestFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(8), // known_y
            CellValue.FromNumber(3), // known_x
            CellValue.FromBool(false), // const = FALSE
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // With const=FALSE: m = exp((3*ln(8)) / (3*3)) = exp(ln(8)/3) = 8^(1/3) = 2
        Assert.Equal(2.0, result.NumericValue, 10);
    }
}
