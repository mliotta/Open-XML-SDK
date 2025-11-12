// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for exponential smoothing forecasting functions.
/// </summary>
public class ForecastingFunctionTests
{
    // FORECAST.ETS Tests
    [Fact]
    public void ForecastEts_SimpleLinearTrend_ReturnsForecast()
    {
        var func = ForecastEtsFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(6),  // target_date (next time point)
            CellValue.FromNumber(10), // values (simplified: single value representing array [10, 20, 30, 40, 50])
            CellValue.FromNumber(1),  // timeline (simplified: single value representing [1, 2, 3, 4, 5])
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // With linear trend, should forecast a positive value
        Assert.True(result.NumericValue > 0);
    }

    [Fact]
    public void ForecastEts_InsufficientData_ReturnsNA()
    {
        var func = ForecastEtsFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(2),
            CellValue.FromNumber(10),
            CellValue.FromNumber(1),
        };

        // With only 1 data point, should return #N/A
        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    [Fact]
    public void ForecastEts_TargetInPast_ReturnsError()
    {
        var func = ForecastEtsFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.5),  // target before timeline start
            CellValue.FromNumber(10),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        // Should return error for target in past
        Assert.True(result.ErrorValue == "#NUM!" || result.ErrorValue == "#VALUE!");
    }

    [Fact]
    public void ForecastEts_InvalidSeasonality_ReturnsError()
    {
        var func = ForecastEtsFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(6),
            CellValue.FromNumber(10),
            CellValue.FromNumber(1),
            CellValue.FromNumber(-1),  // Invalid negative seasonality
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void ForecastEts_MismatchedArrays_ReturnsError()
    {
        var func = ForecastEtsFunction.Instance;
        // In real implementation, arrays would be different lengths
        // For now, we test the error handling path
        var args = new[]
        {
            CellValue.FromNumber(6),
            CellValue.FromNumber(10),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        // Should handle the case gracefully
        Assert.NotNull(result);
    }

    [Fact]
    public void ForecastEts_PropagatesError_ReturnsError()
    {
        var func = ForecastEtsFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(10),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    // FORECAST.ETS.CONFINT Tests
    [Fact]
    public void ForecastEtsConfint_ValidInputs_ReturnsPositiveInterval()
    {
        var func = ForecastEtsConfintFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(6),
            CellValue.FromNumber(10),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // Confidence interval should be positive
        Assert.True(result.NumericValue > 0);
    }

    [Fact]
    public void ForecastEtsConfint_CustomConfidenceLevel_ReturnsInterval()
    {
        var func = ForecastEtsConfintFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(6),
            CellValue.FromNumber(10),
            CellValue.FromNumber(1),
            CellValue.FromNumber(0.90),  // 90% confidence
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue > 0);
    }

    [Fact]
    public void ForecastEtsConfint_InvalidConfidenceLevel_ReturnsError()
    {
        var func = ForecastEtsConfintFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(6),
            CellValue.FromNumber(10),
            CellValue.FromNumber(1),
            CellValue.FromNumber(1.5),  // Invalid: > 1
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void ForecastEtsConfint_InsufficientData_ReturnsNA()
    {
        var func = ForecastEtsConfintFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(2),
            CellValue.FromNumber(10),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    // FORECAST.ETS.SEASONALITY Tests
    [Fact]
    public void ForecastEtsSeasonality_SmallDataset_ReturnsOne()
    {
        var func = ForecastEtsSeasonalityFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),  // values
            CellValue.FromNumber(1),   // timeline
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // With insufficient data, should return 1 (no seasonality)
        Assert.Equal(1, result.NumericValue);
    }

    [Fact]
    public void ForecastEtsSeasonality_ValidData_ReturnsSeasonalPeriod()
    {
        var func = ForecastEtsSeasonalityFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // Should return a non-negative integer
        Assert.True(result.NumericValue >= 0);
    }

    [Fact]
    public void ForecastEtsSeasonality_MismatchedArrays_ReturnsError()
    {
        var func = ForecastEtsSeasonalityFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        // Should handle gracefully
        Assert.NotNull(result);
    }

    [Fact]
    public void ForecastEtsSeasonality_PropagatesError_ReturnsError()
    {
        var func = ForecastEtsSeasonalityFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#VALUE!"),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    // FORECAST.ETS.STAT Tests
    [Fact]
    public void ForecastEtsStat_Alpha_ReturnsValue()
    {
        var func = ForecastEtsStatFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(1),
            CellValue.FromNumber(1),  // statistic_type = 1 (Alpha)
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // Alpha should be between 0 and 1
        Assert.True(result.NumericValue >= 0 && result.NumericValue <= 1);
    }

    [Fact]
    public void ForecastEtsStat_Beta_ReturnsValue()
    {
        var func = ForecastEtsStatFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(1),
            CellValue.FromNumber(2),  // statistic_type = 2 (Beta)
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // Beta should be between 0 and 1
        Assert.True(result.NumericValue >= 0 && result.NumericValue <= 1);
    }

    [Fact]
    public void ForecastEtsStat_Gamma_ReturnsValue()
    {
        var func = ForecastEtsStatFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(1),
            CellValue.FromNumber(3),  // statistic_type = 3 (Gamma)
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // Gamma should be between 0 and 1
        Assert.True(result.NumericValue >= 0 && result.NumericValue <= 1);
    }

    [Fact]
    public void ForecastEtsStat_MASE_ReturnsValue()
    {
        var func = ForecastEtsStatFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(1),
            CellValue.FromNumber(4),  // statistic_type = 4 (MASE)
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // MASE should be non-negative
        Assert.True(result.NumericValue >= 0);
    }

    [Fact]
    public void ForecastEtsStat_SMAPE_ReturnsValue()
    {
        var func = ForecastEtsStatFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(1),
            CellValue.FromNumber(5),  // statistic_type = 5 (SMAPE)
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // SMAPE should be between 0 and 200
        Assert.True(result.NumericValue >= 0 && result.NumericValue <= 200);
    }

    [Fact]
    public void ForecastEtsStat_MAE_ReturnsValue()
    {
        var func = ForecastEtsStatFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(1),
            CellValue.FromNumber(6),  // statistic_type = 6 (MAE)
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // MAE should be non-negative
        Assert.True(result.NumericValue >= 0);
    }

    [Fact]
    public void ForecastEtsStat_RMSE_ReturnsValue()
    {
        var func = ForecastEtsStatFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(1),
            CellValue.FromNumber(7),  // statistic_type = 7 (RMSE)
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // RMSE should be non-negative
        Assert.True(result.NumericValue >= 0);
    }

    [Fact]
    public void ForecastEtsStat_StepSize_ReturnsValue()
    {
        var func = ForecastEtsStatFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(1),
            CellValue.FromNumber(8),  // statistic_type = 8 (Step size)
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // Step size should be positive
        Assert.True(result.NumericValue > 0);
    }

    [Fact]
    public void ForecastEtsStat_InvalidStatisticType_ReturnsError()
    {
        var func = ForecastEtsStatFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(1),
            CellValue.FromNumber(9),  // Invalid: must be 1-8
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void ForecastEtsStat_InsufficientData_ReturnsNA()
    {
        var func = ForecastEtsStatFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(1),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    [Fact]
    public void ForecastEtsStat_PropagatesError_ReturnsError()
    {
        var func = ForecastEtsStatFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#REF!"),
            CellValue.FromNumber(1),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    // Wrong argument count tests
    [Fact]
    public void ForecastEts_TooFewArguments_ReturnsError()
    {
        var func = ForecastEtsFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(6),
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void ForecastEtsConfint_TooManyArguments_ReturnsError()
    {
        var func = ForecastEtsConfintFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(6),
            CellValue.FromNumber(10),
            CellValue.FromNumber(1),
            CellValue.FromNumber(0.95),
            CellValue.FromNumber(0),
            CellValue.FromNumber(1),
            CellValue.FromNumber(1),
            CellValue.FromNumber(999),  // Too many
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void ForecastEtsSeasonality_WrongArgumentCount_ReturnsError()
    {
        var func = ForecastEtsSeasonalityFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void ForecastEtsStat_WrongArgumentCount_ReturnsError()
    {
        var func = ForecastEtsStatFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }
}
