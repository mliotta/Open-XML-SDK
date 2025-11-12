// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for distribution and statistical compatibility functions.
/// </summary>
public class NewDistributionFunctionTests
{
    // STDEV.S Tests
    [Fact]
    public void StDevS_FiveNumbers_ReturnsSampleStandardDeviation()
    {
        var func = StDevSFunction.Instance;
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
        // Sample stdev of 1,2,3,4,5 is sqrt(2.5) ≈ 1.5811388300841898
        Assert.Equal(1.5811388300841898, result.NumericValue, 10);
    }

    // STDEV.P Tests
    [Fact]
    public void StDevP_FiveNumbers_ReturnsPopulationStandardDeviation()
    {
        var func = StDevPFunction2.Instance;
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

    // VAR.S Tests
    [Fact]
    public void VarS_FiveNumbers_ReturnsSampleVariance()
    {
        var func = VarSFunction.Instance;
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
        // Sample variance of 1,2,3,4,5 is 2.5
        Assert.Equal(2.5, result.NumericValue);
    }

    // VAR.P Tests
    [Fact]
    public void VarP_FiveNumbers_ReturnsPopulationVariance()
    {
        var func = VarPFunction2.Instance;
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

    // MODE.SNGL Tests
    [Fact]
    public void ModeSngl_RepeatedValue_ReturnsMode()
    {
        var func = ModeSnglFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(2),
            CellValue.FromNumber(2),
            CellValue.FromNumber(3),
            CellValue.FromNumber(3),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(3.0, result.NumericValue);
    }

    // MODE.MULT Tests
    [Fact]
    public void ModeMult_MultipleMode_ReturnsSmallest()
    {
        var func = ModeMultFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(1),
            CellValue.FromNumber(2),
            CellValue.FromNumber(2),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // Both 1 and 2 appear twice, should return smallest (1)
        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void ModeMult_NoRepeats_ReturnsError()
    {
        var func = ModeMultFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(2),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    // NA Tests
    [Fact]
    public void Na_NoArgs_ReturnsNAError()
    {
        var func = NaFunction.Instance;
        var args = System.Array.Empty<CellValue>();

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    // AREAS Tests
    [Fact]
    public void Areas_SingleReference_ReturnsOne()
    {
        var func = AreasFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void Areas_WrongNumberOfArgs_ReturnsError()
    {
        var func = AreasFunction.Instance;
        var args = System.Array.Empty<CellValue>();

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    // NORM.DIST Tests
    [Fact]
    public void NormDist_Cumulative_ReturnsCorrectValue()
    {
        var func = NormDistFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0),
            CellValue.FromNumber(0),
            CellValue.FromNumber(1),
            CellValue.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // NORM.DIST(0, 0, 1, TRUE) = 0.5
        Assert.Equal(0.5, result.NumericValue, 10);
    }

    [Fact]
    public void NormDist_PDF_ReturnsCorrectValue()
    {
        var func = NormDistFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0),
            CellValue.FromNumber(0),
            CellValue.FromNumber(1),
            CellValue.FromBool(false),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // NORM.DIST(0, 0, 1, FALSE) = 0.3989422804014327 (1/sqrt(2*pi))
        Assert.Equal(0.3989422804014327, result.NumericValue, 10);
    }

    [Fact]
    public void NormDist_NegativeStdDev_ReturnsError()
    {
        var func = NormDistFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0),
            CellValue.FromNumber(0),
            CellValue.FromNumber(-1),
            CellValue.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    // NORM.INV Tests
    [Fact]
    public void NormInv_MedianProbability_ReturnsMean()
    {
        var func = NormInvFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.5),
            CellValue.FromNumber(10),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // NORM.INV(0.5, 10, 2) = 10
        Assert.Equal(10.0, result.NumericValue, 10);
    }

    [Fact]
    public void NormInv_InvalidProbability_ReturnsError()
    {
        var func = NormInvFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1.5),
            CellValue.FromNumber(0),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    // NORM.S.DIST Tests
    [Fact]
    public void NormSDist_Zero_ReturnsHalf()
    {
        var func = NormSDistFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0),
            CellValue.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // NORM.S.DIST(0, TRUE) = 0.5
        Assert.Equal(0.5, result.NumericValue, 10);
    }

    [Fact]
    public void NormSDist_PDF_AtZero_ReturnsCorrectValue()
    {
        var func = NormSDistFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0),
            CellValue.FromBool(false),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // NORM.S.DIST(0, FALSE) = 0.3989422804014327
        Assert.Equal(0.3989422804014327, result.NumericValue, 10);
    }

    // NORM.S.INV Tests
    [Fact]
    public void NormSInv_MedianProbability_ReturnsZero()
    {
        var func = NormSInvFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.5),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // NORM.S.INV(0.5) = 0
        Assert.Equal(0.0, result.NumericValue, 10);
    }

    [Fact]
    public void NormSInv_InvalidProbability_ReturnsError()
    {
        var func = NormSInvFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    // CONFIDENCE Tests
    [Fact]
    public void Confidence_ValidArgs_ReturnsInterval()
    {
        var func = ConfidenceFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.05),
            CellValue.FromNumber(2.5),
            CellValue.FromNumber(100),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // CONFIDENCE(0.05, 2.5, 100) ≈ 0.489939 (1.96 * 2.5 / sqrt(100))
        Assert.InRange(result.NumericValue, 0.48, 0.50);
    }

    [Fact]
    public void Confidence_InvalidAlpha_ReturnsError()
    {
        var func = ConfidenceFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1.5),
            CellValue.FromNumber(2.5),
            CellValue.FromNumber(100),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    // CONFIDENCE.NORM Tests
    [Fact]
    public void ConfidenceNorm_SameAsConfidence()
    {
        var func = ConfidenceNormFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.05),
            CellValue.FromNumber(2.5),
            CellValue.FromNumber(100),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // CONFIDENCE.NORM should give same result as CONFIDENCE
        Assert.InRange(result.NumericValue, 0.48, 0.50);
    }

    // CONFIDENCE.T Tests
    [Fact]
    public void ConfidenceT_ValidArgs_ReturnsInterval()
    {
        var func = ConfidenceTFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.05),
            CellValue.FromNumber(2.5),
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // CONFIDENCE.T(0.05, 2.5, 10) should be larger than CONFIDENCE due to t-distribution
        Assert.InRange(result.NumericValue, 1.5, 2.0);
    }

    [Fact]
    public void ConfidenceT_LargeSample_SimilarToNorm()
    {
        var func = ConfidenceTFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.05),
            CellValue.FromNumber(2.5),
            CellValue.FromNumber(1000),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // For large samples, t-distribution approaches normal
        Assert.InRange(result.NumericValue, 0.48, 0.50);
    }
}
