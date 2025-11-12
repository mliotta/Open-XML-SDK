// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for statistical distribution functions.
/// </summary>
public class DistributionFunctionTests
{
    // T.DIST Tests
    [Fact]
    public void TDist_CDF_AtZero_ReturnsHalf()
    {
        var func = TDistFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0),
            CellValue.FromNumber(10),
            CellValue.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.5, result.NumericValue, 5);
    }

    [Fact]
    public void TDist_PDF_AtZero_ReturnsCorrectValue()
    {
        var func = TDistFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0),
            CellValue.FromNumber(10),
            CellValue.FromBool(false),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // PDF at 0 with df=10 should be approximately 0.3891
        Assert.InRange(result.NumericValue, 0.38, 0.40);
    }

    [Fact]
    public void TDistRT_PositiveValue_ReturnsRightTail()
    {
        var func = TDistRTFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1.812),
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // Right tail for t=1.812, df=10 should be approximately 0.05
        Assert.InRange(result.NumericValue, 0.04, 0.06);
    }

    [Fact]
    public void TDist2T_PositiveValue_ReturnsTwoTail()
    {
        var func = TDist2TFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1.812),
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // Two-tailed for t=1.812, df=10 should be approximately 0.10
        Assert.InRange(result.NumericValue, 0.08, 0.12);
    }

    [Fact]
    public void TInv_Median_ReturnsZero()
    {
        var func = TInvFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.5),
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.0, result.NumericValue, 5);
    }

    [Fact]
    public void TInv2T_Alpha05_ReturnsCorrectValue()
    {
        var func = TInv2TFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.05),
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // Two-tailed critical value for alpha=0.05, df=10 is approximately 2.228
        Assert.InRange(result.NumericValue, 2.1, 2.3);
    }

    // Legacy TDIST Tests
    [Fact]
    public void TDistLegacy_OneTailed_ReturnsRightTail()
    {
        var func = TDistLegacyFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1.5),
            CellValue.FromNumber(10),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.InRange(result.NumericValue, 0.05, 0.10);
    }

    [Fact]
    public void TInvLegacy_ReturnsPositiveValue()
    {
        var func = TInvLegacyFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.05),
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.InRange(result.NumericValue, 2.1, 2.3);
    }

    // CHISQ.DIST Tests
    [Fact]
    public void ChiSqDist_CDF_ReturnsCorrectValue()
    {
        var func = ChiSqDistFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(5),
            CellValue.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // CDF at x=5, df=5 should be approximately 0.584
        Assert.InRange(result.NumericValue, 0.55, 0.62);
    }

    [Fact]
    public void ChiSqDistRT_ReturnsRightTail()
    {
        var func = ChiSqDistRTFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(11.07),
            CellValue.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // Right tail for chi-square=11.07, df=5 should be approximately 0.05
        Assert.InRange(result.NumericValue, 0.04, 0.06);
    }

    [Fact]
    public void ChiSqInv_MedianProbability_ReturnsMedian()
    {
        var func = ChiSqInvFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.5),
            CellValue.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // Median of chi-square with df=5 is approximately 4.35
        Assert.InRange(result.NumericValue, 4.0, 4.7);
    }

    [Fact]
    public void ChiSqInvRT_Alpha05_ReturnsCorrectValue()
    {
        var func = ChiSqInvRTFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.05),
            CellValue.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // Critical value for right-tail 0.05, df=5 is approximately 11.07
        Assert.InRange(result.NumericValue, 10.5, 11.5);
    }

    // F.DIST Tests
    [Fact]
    public void FDist_CDF_ReturnsCorrectValue()
    {
        var func = FDistFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(2),
            CellValue.FromNumber(5),
            CellValue.FromNumber(10),
            CellValue.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // CDF should be between 0 and 1
        Assert.InRange(result.NumericValue, 0.7, 0.9);
    }

    [Fact]
    public void FDistRT_ReturnsRightTail()
    {
        var func = FDistRTFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(3.33),
            CellValue.FromNumber(5),
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // Right tail should be approximately 0.05 for critical value
        Assert.InRange(result.NumericValue, 0.04, 0.07);
    }

    [Fact]
    public void FInv_MedianProbability_ReturnsMedian()
    {
        var func = FInvFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.5),
            CellValue.FromNumber(5),
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // Median should be close to 1
        Assert.InRange(result.NumericValue, 0.8, 1.2);
    }

    [Fact]
    public void FInvRT_Alpha05_ReturnsCorrectValue()
    {
        var func = FInvRTFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.05),
            CellValue.FromNumber(5),
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // Critical value for alpha=0.05, df1=5, df2=10 is approximately 3.33
        Assert.InRange(result.NumericValue, 3.0, 3.6);
    }

    // BETA.DIST Tests
    [Fact]
    public void BetaDist_CDF_AtHalf_ReturnsHalf()
    {
        var func = BetaDistFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.5),
            CellValue.FromNumber(2),
            CellValue.FromNumber(2),
            CellValue.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // For symmetric beta(2,2), CDF at 0.5 should be 0.5
        Assert.Equal(0.5, result.NumericValue, 2);
    }

    [Fact]
    public void BetaDist_PDF_AtHalf_ReturnsMaximum()
    {
        var func = BetaDistFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.5),
            CellValue.FromNumber(2),
            CellValue.FromNumber(2),
            CellValue.FromBool(false),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // PDF for beta(2,2) at 0.5 is 1.5
        Assert.Equal(1.5, result.NumericValue, 2);
    }

    [Fact]
    public void BetaDist_WithBounds_ReturnsScaledValue()
    {
        var func = BetaDistFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(2),
            CellValue.FromNumber(2),
            CellValue.FromBool(true),
            CellValue.FromNumber(0),
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // x=5 in [0,10] maps to 0.5 in [0,1]
        Assert.Equal(0.5, result.NumericValue, 2);
    }

    [Fact]
    public void BetaInv_MedianProbability_ReturnsMedian()
    {
        var func = BetaInvFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.5),
            CellValue.FromNumber(2),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // For symmetric beta(2,2), inverse at 0.5 should be 0.5
        Assert.Equal(0.5, result.NumericValue, 2);
    }

    // LOGNORM.DIST Tests
    [Fact]
    public void LogNormDist_CDF_ReturnsCorrectValue()
    {
        var func = LogNormDistFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(0),
            CellValue.FromNumber(1),
            CellValue.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // Lognormal CDF at x=1, mean=0, sd=1 should be 0.5
        Assert.Equal(0.5, result.NumericValue, 5);
    }

    [Fact]
    public void LogNormDist_PDF_ReturnsCorrectValue()
    {
        var func = LogNormDistFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(0),
            CellValue.FromNumber(1),
            CellValue.FromBool(false),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // PDF at x=1, mean=0, sd=1 is 1/sqrt(2*pi) ≈ 0.3989
        Assert.InRange(result.NumericValue, 0.35, 0.42);
    }

    [Fact]
    public void LogNormInv_MedianProbability_ReturnsExpMean()
    {
        var func = LogNormInvFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.5),
            CellValue.FromNumber(2),
            CellValue.FromNumber(0.5),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // At median probability, should return exp(mean) = exp(2) ≈ 7.389
        Assert.InRange(result.NumericValue, 7.0, 7.8);
    }

    // Error handling tests
    [Fact]
    public void TDist_NegativeDF_ReturnsError()
    {
        var func = TDistFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(-1),
            CellValue.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void ChiSqInv_InvalidProbability_ReturnsError()
    {
        var func = ChiSqInvFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1.5),
            CellValue.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void BetaDist_InvalidBounds_ReturnsError()
    {
        var func = BetaDistFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(2),
            CellValue.FromNumber(2),
            CellValue.FromBool(true),
            CellValue.FromNumber(10),
            CellValue.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void LogNormDist_NegativeX_ReturnsError()
    {
        var func = LogNormDistFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-1),
            CellValue.FromNumber(0),
            CellValue.FromNumber(1),
            CellValue.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }
}
