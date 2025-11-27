// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for multi-criteria functions (SUMIFS, COUNTIFS, AVERAGEIFS, MAXIFS, MINIFS).
/// </summary>
public class MultiCriteriaFunctionTests
{
    #region SUMIFS Tests

    [Fact]
    public void SumIfs_SingleCriteria_ReturnsSum()
    {
        var func = SumIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(100),
            FormulaResult.FromNumber(10),
            FormulaResult.FromString(">5"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(100.0, result.NumericValue);
    }

    [Fact]
    public void SumIfs_SingleCriteriaNotMet_ReturnsZero()
    {
        var func = SumIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(100),
            FormulaResult.FromNumber(3),
            FormulaResult.FromString(">5"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void SumIfs_MultipleCriteriaBothMet_ReturnsSum()
    {
        var func = SumIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(100),
            FormulaResult.FromNumber(10),
            FormulaResult.FromString(">5"),
            FormulaResult.FromString("Yes"),
            FormulaResult.FromString("Yes"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(100.0, result.NumericValue);
    }

    [Fact]
    public void SumIfs_MultipleCriteriaOneMet_ReturnsZero()
    {
        var func = SumIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(100),
            FormulaResult.FromNumber(10),
            FormulaResult.FromString(">5"),
            FormulaResult.FromString("No"),
            FormulaResult.FromString("Yes"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void SumIfs_ThreeCriteriaAllMet_ReturnsSum()
    {
        var func = SumIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(250),
            FormulaResult.FromNumber(15),
            FormulaResult.FromString(">=10"),
            FormulaResult.FromString("North"),
            FormulaResult.FromString("North"),
            FormulaResult.FromNumber(100),
            FormulaResult.FromString(">50"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(250.0, result.NumericValue);
    }

    [Fact]
    public void SumIfs_NonNumericSumRange_ReturnsZero()
    {
        var func = SumIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("text"),
            FormulaResult.FromNumber(10),
            FormulaResult.FromString(">5"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void SumIfs_InsufficientArguments_ReturnsError()
    {
        var func = SumIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(100),
            FormulaResult.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void SumIfs_EvenNumberOfArguments_ReturnsError()
    {
        var func = SumIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(100),
            FormulaResult.FromNumber(10),
            FormulaResult.FromString(">5"),
            FormulaResult.FromString("Yes"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void SumIfs_ErrorValue_PropagatesError()
    {
        var func = SumIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#DIV/0!"),
            FormulaResult.FromNumber(10),
            FormulaResult.FromString(">5"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    #endregion

    #region COUNTIFS Tests

    [Fact]
    public void CountIfs_SingleCriteria_ReturnsCount()
    {
        var func = CountIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(10),
            FormulaResult.FromString(">5"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void CountIfs_SingleCriteriaNotMet_ReturnsZero()
    {
        var func = CountIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(3),
            FormulaResult.FromString(">5"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void CountIfs_MultipleCriteriaBothMet_ReturnsCount()
    {
        var func = CountIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(10),
            FormulaResult.FromString(">5"),
            FormulaResult.FromNumber(8),
            FormulaResult.FromString("<10"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void CountIfs_MultipleCriteriaOneMet_ReturnsZero()
    {
        var func = CountIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(10),
            FormulaResult.FromString(">5"),
            FormulaResult.FromNumber(12),
            FormulaResult.FromString("<10"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void CountIfs_ThreeCriteriaAllMet_ReturnsCount()
    {
        var func = CountIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(15),
            FormulaResult.FromString(">=10"),
            FormulaResult.FromString("North"),
            FormulaResult.FromString("North"),
            FormulaResult.FromNumber(100),
            FormulaResult.FromString(">50"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void CountIfs_TextCriteria_Matches()
    {
        var func = CountIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Apple"),
            FormulaResult.FromString("Apple"),
            FormulaResult.FromNumber(5),
            FormulaResult.FromString(">3"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void CountIfs_InsufficientArguments_ReturnsError()
    {
        var func = CountIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void CountIfs_OddNumberOfArguments_ReturnsError()
    {
        var func = CountIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(10),
            FormulaResult.FromString(">5"),
            FormulaResult.FromNumber(8),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void CountIfs_ErrorValue_PropagatesError()
    {
        var func = CountIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#REF!"),
            FormulaResult.FromString(">5"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    #endregion

    #region AVERAGEIFS Tests

    [Fact]
    public void AverageIfs_SingleCriteria_ReturnsAverage()
    {
        var func = AverageIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(100),
            FormulaResult.FromNumber(10),
            FormulaResult.FromString(">=10"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(100.0, result.NumericValue);
    }

    [Fact]
    public void AverageIfs_SingleCriteriaNotMet_ReturnsDivError()
    {
        var func = AverageIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(100),
            FormulaResult.FromNumber(3),
            FormulaResult.FromString(">5"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void AverageIfs_MultipleCriteriaBothMet_ReturnsAverage()
    {
        var func = AverageIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(75),
            FormulaResult.FromNumber(10),
            FormulaResult.FromString(">5"),
            FormulaResult.FromString("Yes"),
            FormulaResult.FromString("Yes"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(75.0, result.NumericValue);
    }

    [Fact]
    public void AverageIfs_MultipleCriteriaOneMet_ReturnsDivError()
    {
        var func = AverageIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(75),
            FormulaResult.FromNumber(10),
            FormulaResult.FromString(">5"),
            FormulaResult.FromString("No"),
            FormulaResult.FromString("Yes"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void AverageIfs_ThreeCriteriaAllMet_ReturnsAverage()
    {
        var func = AverageIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(150),
            FormulaResult.FromNumber(15),
            FormulaResult.FromString(">=10"),
            FormulaResult.FromString("North"),
            FormulaResult.FromString("North"),
            FormulaResult.FromNumber(100),
            FormulaResult.FromString(">50"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(150.0, result.NumericValue);
    }

    [Fact]
    public void AverageIfs_NonNumericAverageRange_ReturnsDivError()
    {
        var func = AverageIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("text"),
            FormulaResult.FromNumber(10),
            FormulaResult.FromString(">5"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void AverageIfs_InsufficientArguments_ReturnsError()
    {
        var func = AverageIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(100),
            FormulaResult.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void AverageIfs_EvenNumberOfArguments_ReturnsError()
    {
        var func = AverageIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(100),
            FormulaResult.FromNumber(10),
            FormulaResult.FromString(">5"),
            FormulaResult.FromString("Yes"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void AverageIfs_ErrorValue_PropagatesError()
    {
        var func = AverageIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#N/A"),
            FormulaResult.FromNumber(10),
            FormulaResult.FromString(">5"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    #endregion

    #region MAXIFS Tests

    [Fact]
    public void MaxIfs_SingleCriteria_ReturnsMax()
    {
        var func = MaxIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(100),
            FormulaResult.FromString("North"),
            FormulaResult.FromString("North"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(100.0, result.NumericValue);
    }

    [Fact]
    public void MaxIfs_SingleCriteriaNotMet_ReturnsZero()
    {
        var func = MaxIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(100),
            FormulaResult.FromString("South"),
            FormulaResult.FromString("North"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void MaxIfs_MultipleCriteriaBothMet_ReturnsMax()
    {
        var func = MaxIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(250),
            FormulaResult.FromString("North"),
            FormulaResult.FromString("North"),
            FormulaResult.FromNumber(120),
            FormulaResult.FromString(">100"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(250.0, result.NumericValue);
    }

    [Fact]
    public void MaxIfs_MultipleCriteriaOneMet_ReturnsZero()
    {
        var func = MaxIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(250),
            FormulaResult.FromString("North"),
            FormulaResult.FromString("North"),
            FormulaResult.FromNumber(80),
            FormulaResult.FromString(">100"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void MaxIfs_ThreeCriteriaAllMet_ReturnsMax()
    {
        var func = MaxIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(500),
            FormulaResult.FromString("North"),
            FormulaResult.FromString("North"),
            FormulaResult.FromNumber(120),
            FormulaResult.FromString(">100"),
            FormulaResult.FromString("Q1"),
            FormulaResult.FromString("Q1"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(500.0, result.NumericValue);
    }

    [Fact]
    public void MaxIfs_NonNumericMaxRange_ReturnsZero()
    {
        var func = MaxIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("text"),
            FormulaResult.FromString("North"),
            FormulaResult.FromString("North"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void MaxIfs_InsufficientArguments_ReturnsError()
    {
        var func = MaxIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(100),
            FormulaResult.FromString("North"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void MaxIfs_EvenNumberOfArguments_ReturnsError()
    {
        var func = MaxIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(100),
            FormulaResult.FromString("North"),
            FormulaResult.FromString("North"),
            FormulaResult.FromNumber(120),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void MaxIfs_ErrorValue_PropagatesError()
    {
        var func = MaxIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#DIV/0!"),
            FormulaResult.FromString("North"),
            FormulaResult.FromString("North"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    #endregion

    #region MINIFS Tests

    [Fact]
    public void MinIfs_SingleCriteria_ReturnsMin()
    {
        var func = MinIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(50),
            FormulaResult.FromString("South"),
            FormulaResult.FromString("South"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(50.0, result.NumericValue);
    }

    [Fact]
    public void MinIfs_SingleCriteriaNotMet_ReturnsZero()
    {
        var func = MinIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(50),
            FormulaResult.FromString("North"),
            FormulaResult.FromString("South"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void MinIfs_MultipleCriteriaBothMet_ReturnsMin()
    {
        var func = MinIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(30),
            FormulaResult.FromString("South"),
            FormulaResult.FromString("South"),
            FormulaResult.FromNumber(25),
            FormulaResult.FromString("<50"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(30.0, result.NumericValue);
    }

    [Fact]
    public void MinIfs_MultipleCriteriaOneMet_ReturnsZero()
    {
        var func = MinIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(30),
            FormulaResult.FromString("South"),
            FormulaResult.FromString("South"),
            FormulaResult.FromNumber(60),
            FormulaResult.FromString("<50"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void MinIfs_ThreeCriteriaAllMet_ReturnsMin()
    {
        var func = MinIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(25),
            FormulaResult.FromString("South"),
            FormulaResult.FromString("South"),
            FormulaResult.FromNumber(30),
            FormulaResult.FromString("<50"),
            FormulaResult.FromString("Q2"),
            FormulaResult.FromString("Q2"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(25.0, result.NumericValue);
    }

    [Fact]
    public void MinIfs_NonNumericMinRange_ReturnsZero()
    {
        var func = MinIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("text"),
            FormulaResult.FromString("South"),
            FormulaResult.FromString("South"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void MinIfs_InsufficientArguments_ReturnsError()
    {
        var func = MinIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(50),
            FormulaResult.FromString("South"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void MinIfs_EvenNumberOfArguments_ReturnsError()
    {
        var func = MinIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(50),
            FormulaResult.FromString("South"),
            FormulaResult.FromString("South"),
            FormulaResult.FromNumber(30),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void MinIfs_ErrorValue_PropagatesError()
    {
        var func = MinIfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#REF!"),
            FormulaResult.FromString("South"),
            FormulaResult.FromString("South"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    #endregion
}
