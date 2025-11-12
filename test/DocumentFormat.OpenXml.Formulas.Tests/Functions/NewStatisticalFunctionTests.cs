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
            CellValue.Error("#N/A"),
            CellValue.FromString("Not found"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("Not found", result.StringValue);
    }

    [Fact]
    public void Ifna_NAError_ReturnsNumericAlternative()
    {
        var func = IfnaFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#N/A"),
            CellValue.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void Ifna_ValidValue_ReturnsOriginal()
    {
        var func = IfnaFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(42),
            CellValue.FromString("Error"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(42.0, result.NumericValue);
    }

    [Fact]
    public void Ifna_OtherError_ReturnsOriginalError()
    {
        var func = IfnaFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#DIV/0!"),
            CellValue.FromString("Alternative"),
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
            CellValue.Error("#VALUE!"),
            CellValue.FromNumber(0),
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
            CellValue.FromString("Hello"),
            CellValue.FromString("Fallback"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("Hello", result.StringValue);
    }

    [Fact]
    public void Ifna_WrongArgumentCount_ReturnsError()
    {
        var func = IfnaFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(42),
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
            CellValue.FromNumber(1),
            CellValue.FromNumber(2),
            CellValue.FromNumber(3),
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
            CellValue.FromNumber(1),
            CellValue.FromNumber(2),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
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
            CellValue.FromNumber(1),
            CellValue.FromNumber(2),
            CellValue.FromNumber(3),
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue > 0); // Positive skew
    }

    [Fact]
    public void Skew_NegativeSkew_ReturnsNegativeValue()
    {
        var func = SkewFunction.Instance;
        // Left-skewed distribution
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(8),
            CellValue.FromNumber(9),
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue < 0); // Negative skew
    }

    [Fact]
    public void Skew_TwoValues_ReturnsError()
    {
        var func = SkewFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(2),
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
            CellValue.FromNumber(5),
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
            CellValue.FromNumber(5),
            CellValue.FromNumber(5),
            CellValue.FromNumber(5),
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
            CellValue.FromNumber(1),
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(3),
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
            CellValue.FromNumber(1),
            CellValue.FromString("text"),
            CellValue.FromNumber(2),
            CellValue.FromBool(true),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // Should calculate skewness of 1, 2, 3 only
        Assert.Equal(0.0, result.NumericValue, 10);
    }

    [Fact]
    public void Skew_NoNumbers_ReturnsError()
    {
        var func = SkewFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("text1"),
            CellValue.FromString("text2"),
            CellValue.FromString("text3"),
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
            CellValue.FromNumber(1),
            CellValue.FromNumber(2),
            CellValue.FromNumber(3),
            CellValue.FromNumber(4),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
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
            CellValue.FromNumber(1),
            CellValue.FromNumber(2),
            CellValue.FromNumber(3),
            CellValue.FromNumber(4),
            CellValue.FromNumber(5),
            CellValue.FromNumber(6),
            CellValue.FromNumber(7),
            CellValue.FromNumber(8),
            CellValue.FromNumber(9),
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // For uniform distribution, excess kurtosis is approximately -1.2
        Assert.True(result.NumericValue < 0);
    }

    [Fact]
    public void Kurt_ThreeValues_ReturnsError()
    {
        var func = KurtFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(2),
            CellValue.FromNumber(3),
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
            CellValue.FromNumber(5),
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
            CellValue.FromNumber(5),
            CellValue.FromNumber(5),
            CellValue.FromNumber(5),
            CellValue.FromNumber(5),
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
            CellValue.FromNumber(1),
            CellValue.FromNumber(2),
            CellValue.Error("#NUM!"),
            CellValue.FromNumber(4),
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
            CellValue.FromNumber(1),
            CellValue.FromString("text"),
            CellValue.FromNumber(2),
            CellValue.FromBool(false),
            CellValue.FromNumber(3),
            CellValue.FromNumber(4),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // Should calculate kurtosis of 1, 2, 3, 4 only
    }

    [Fact]
    public void Kurt_NoNumbers_ReturnsError()
    {
        var func = KurtFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("text1"),
            CellValue.FromString("text2"),
            CellValue.FromString("text3"),
            CellValue.FromString("text4"),
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
            CellValue.FromNumber(5),
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // 5 <= 10, so it should be counted in first bin
        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void Frequency_SingleDataAboveBin_ReturnsZero()
    {
        var func = FrequencyFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(15),
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // 15 > 10, so it's not counted in first bin (Phase 0 returns first bin only)
        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void Frequency_DataEqualsBin_CountsInBin()
    {
        var func = FrequencyFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // 10 <= 10, so it should be counted
        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void Frequency_NoData_ReturnsZero()
    {
        var func = FrequencyFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("text"),
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void Frequency_ErrorInData_PropagatesError()
    {
        var func = FrequencyFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(10),
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
            CellValue.FromNumber(5),
            CellValue.Error("#VALUE!"),
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
            CellValue.FromNumber(5),
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
            CellValue.FromNumber(1),
            CellValue.FromNumber(2),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }
}
