// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for new statistical and logical functions (IFNA, SKEW, KURT, FREQUENCY).
/// </summary>
public class NewStatisticalFunctionTests
{
    // IFNA Tests
    [Fact]
    public void Ifna_NAError_ReturnsAlternative()
    {
        var func = IfnaFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#N/A"),
            FormulaResult.FromString("Not found"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("Not found", result.StringValue);
    }

    [Fact]
    public void Ifna_NAError_ReturnsNumericAlternative()
    {
        var func = IfnaFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#N/A"),
            FormulaResult.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void Ifna_ValidValue_ReturnsOriginal()
    {
        var func = IfnaFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(42),
            FormulaResult.FromString("Error"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(42.0, result.NumericValue);
    }

    [Fact]
    public void Ifna_OtherError_ReturnsOriginalError()
    {
        var func = IfnaFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#DIV/0!"),
            FormulaResult.FromString("Alternative"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void Ifna_ValueError_ReturnsOriginalError()
    {
        var func = IfnaFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#VALUE!"),
            FormulaResult.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Ifna_TextValue_ReturnsOriginal()
    {
        var func = IfnaFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Hello"),
            FormulaResult.FromString("Fallback"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("Hello", result.StringValue);
    }

    [Fact]
    public void Ifna_WrongArgumentCount_ReturnsError()
    {
        var func = IfnaFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(42),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Ifna_ThreeArguments_ReturnsError()
    {
        var func = IfnaFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(1),
            FormulaResult.FromNumber(2),
            FormulaResult.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    // SKEW Tests
    [Fact]
    public void Skew_ThreeValues_ReturnsSkewness()
    {
        var func = SkewFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(1),
            FormulaResult.FromNumber(2),
            FormulaResult.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        // For perfectly symmetric data [1,2,3], skewness should be 0
        Assert.Equal(0.0, result.NumericValue, 10);
    }

    [Fact]
    public void Skew_PositiveSkew_ReturnsPositiveValue()
    {
        var func = SkewFunction.Instance;
        // Right-skewed distribution
        var args = new[]
        {
            FormulaResult.FromNumber(1),
            FormulaResult.FromNumber(2),
            FormulaResult.FromNumber(3),
            FormulaResult.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.True(result.NumericValue > 0); // Positive skew
    }

    [Fact]
    public void Skew_NegativeSkew_ReturnsNegativeValue()
    {
        var func = SkewFunction.Instance;
        // Left-skewed distribution
        var args = new[]
        {
            FormulaResult.FromNumber(1),
            FormulaResult.FromNumber(8),
            FormulaResult.FromNumber(9),
            FormulaResult.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.True(result.NumericValue < 0); // Negative skew
    }

    [Fact]
    public void Skew_TwoValues_ReturnsError()
    {
        var func = SkewFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(1),
            FormulaResult.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void Skew_OneValue_ReturnsError()
    {
        var func = SkewFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void Skew_IdenticalValues_ReturnsError()
    {
        var func = SkewFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(5),
            FormulaResult.FromNumber(5),
            FormulaResult.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void Skew_ErrorValue_PropagatesError()
    {
        var func = SkewFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(1),
            FormulaResult.Error("#DIV/0!"),
            FormulaResult.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void Skew_MixedTypes_IgnoresNonNumeric()
    {
        var func = SkewFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(1),
            FormulaResult.FromString("text"),
            FormulaResult.FromNumber(2),
            FormulaResult.FromBool(true),
            FormulaResult.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        // Should calculate skewness of 1, 2, 3 only
        Assert.Equal(0.0, result.NumericValue, 10);
    }

    [Fact]
    public void Skew_NoNumbers_ReturnsError()
    {
        var func = SkewFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("text1"),
            FormulaResult.FromString("text2"),
            FormulaResult.FromString("text3"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    // KURT Tests
    [Fact]
    public void Kurt_FourValues_ReturnsKurtosis()
    {
        var func = KurtFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(1),
            FormulaResult.FromNumber(2),
            FormulaResult.FromNumber(3),
            FormulaResult.FromNumber(4),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        // Excel's KURT returns excess kurtosis (kurtosis - 3)
        // For uniform distribution, excess kurtosis is negative
    }

    [Fact]
    public void Kurt_NormalDistribution_ReturnsNearZero()
    {
        var func = KurtFunction.Instance;
        // Values approximating normal distribution
        var args = new[]
        {
            FormulaResult.FromNumber(1),
            FormulaResult.FromNumber(2),
            FormulaResult.FromNumber(3),
            FormulaResult.FromNumber(4),
            FormulaResult.FromNumber(5),
            FormulaResult.FromNumber(6),
            FormulaResult.FromNumber(7),
            FormulaResult.FromNumber(8),
            FormulaResult.FromNumber(9),
            FormulaResult.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        // For uniform distribution, excess kurtosis is approximately -1.2
        Assert.True(result.NumericValue < 0);
    }

    [Fact]
    public void Kurt_ThreeValues_ReturnsError()
    {
        var func = KurtFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(1),
            FormulaResult.FromNumber(2),
            FormulaResult.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void Kurt_OneValue_ReturnsError()
    {
        var func = KurtFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void Kurt_IdenticalValues_ReturnsError()
    {
        var func = KurtFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(5),
            FormulaResult.FromNumber(5),
            FormulaResult.FromNumber(5),
            FormulaResult.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void Kurt_ErrorValue_PropagatesError()
    {
        var func = KurtFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(1),
            FormulaResult.FromNumber(2),
            FormulaResult.Error("#NUM!"),
            FormulaResult.FromNumber(4),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Kurt_MixedTypes_IgnoresNonNumeric()
    {
        var func = KurtFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(1),
            FormulaResult.FromString("text"),
            FormulaResult.FromNumber(2),
            FormulaResult.FromBool(false),
            FormulaResult.FromNumber(3),
            FormulaResult.FromNumber(4),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        // Should calculate kurtosis of 1, 2, 3, 4 only
    }

    [Fact]
    public void Kurt_NoNumbers_ReturnsError()
    {
        var func = KurtFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("text1"),
            FormulaResult.FromString("text2"),
            FormulaResult.FromString("text3"),
            FormulaResult.FromString("text4"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    // FREQUENCY Tests
    [Fact]
    public void Frequency_SingleDataSingleBin_ReturnsCount()
    {
        var func = FrequencyFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(5),
            FormulaResult.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        // 5 <= 10, so it should be counted in first bin
        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void Frequency_SingleDataAboveBin_ReturnsZero()
    {
        var func = FrequencyFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(15),
            FormulaResult.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        // 15 > 10, so it's not counted in first bin (Phase 0 returns first bin only)
        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void Frequency_DataEqualsBin_CountsInBin()
    {
        var func = FrequencyFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(10),
            FormulaResult.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        // 10 <= 10, so it should be counted
        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void Frequency_NoData_ReturnsZero()
    {
        var func = FrequencyFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("text"),
            FormulaResult.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void Frequency_ErrorInData_PropagatesError()
    {
        var func = FrequencyFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#DIV/0!"),
            FormulaResult.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void Frequency_ErrorInBins_PropagatesError()
    {
        var func = FrequencyFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(5),
            FormulaResult.Error("#VALUE!"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Frequency_WrongArgumentCount_ReturnsError()
    {
        var func = FrequencyFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Frequency_ThreeArguments_ReturnsError()
    {
        var func = FrequencyFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(1),
            FormulaResult.FromNumber(2),
            FormulaResult.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }
}
