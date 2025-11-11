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
            CellValue.FromNumber(100),
            CellValue.FromNumber(10),
            CellValue.FromString(">5"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(100.0, result.NumericValue);
    }

    [Fact]
    public void SumIfs_SingleCriteriaNotMet_ReturnsZero()
    {
        var func = SumIfsFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(100),
            CellValue.FromNumber(3),
            CellValue.FromString(">5"),
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
            CellValue.FromNumber(100),
            CellValue.FromNumber(10),
            CellValue.FromString(">5"),
            CellValue.FromString("Yes"),
            CellValue.FromString("Yes"),
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
            CellValue.FromNumber(100),
            CellValue.FromNumber(10),
            CellValue.FromString(">5"),
            CellValue.FromString("No"),
            CellValue.FromString("Yes"),
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
            CellValue.FromNumber(250),
            CellValue.FromNumber(15),
            CellValue.FromString(">=10"),
            CellValue.FromString("North"),
            CellValue.FromString("North"),
            CellValue.FromNumber(100),
            CellValue.FromString(">50"),
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
            CellValue.FromString("text"),
            CellValue.FromNumber(10),
            CellValue.FromString(">5"),
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
            CellValue.FromNumber(100),
            CellValue.FromNumber(10),
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
            CellValue.FromNumber(100),
            CellValue.FromNumber(10),
            CellValue.FromString(">5"),
            CellValue.FromString("Yes"),
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
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(10),
            CellValue.FromString(">5"),
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
            CellValue.FromNumber(10),
            CellValue.FromString(">5"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void CountIfs_SingleCriteriaNotMet_ReturnsZero()
    {
        var func = CountIfsFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(3),
            CellValue.FromString(">5"),
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
            CellValue.FromNumber(10),
            CellValue.FromString(">5"),
            CellValue.FromNumber(8),
            CellValue.FromString("<10"),
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
            CellValue.FromNumber(10),
            CellValue.FromString(">5"),
            CellValue.FromNumber(12),
            CellValue.FromString("<10"),
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
            CellValue.FromNumber(15),
            CellValue.FromString(">=10"),
            CellValue.FromString("North"),
            CellValue.FromString("North"),
            CellValue.FromNumber(100),
            CellValue.FromString(">50"),
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
            CellValue.FromString("Apple"),
            CellValue.FromString("Apple"),
            CellValue.FromNumber(5),
            CellValue.FromString(">3"),
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
            CellValue.FromNumber(10),
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
            CellValue.FromNumber(10),
            CellValue.FromString(">5"),
            CellValue.FromNumber(8),
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
            CellValue.Error("#REF!"),
            CellValue.FromString(">5"),
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
            CellValue.FromNumber(100),
            CellValue.FromNumber(10),
            CellValue.FromString(">=10"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(100.0, result.NumericValue);
    }

    [Fact]
    public void AverageIfs_SingleCriteriaNotMet_ReturnsDivError()
    {
        var func = AverageIfsFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(100),
            CellValue.FromNumber(3),
            CellValue.FromString(">5"),
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
            CellValue.FromNumber(75),
            CellValue.FromNumber(10),
            CellValue.FromString(">5"),
            CellValue.FromString("Yes"),
            CellValue.FromString("Yes"),
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
            CellValue.FromNumber(75),
            CellValue.FromNumber(10),
            CellValue.FromString(">5"),
            CellValue.FromString("No"),
            CellValue.FromString("Yes"),
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
            CellValue.FromNumber(150),
            CellValue.FromNumber(15),
            CellValue.FromString(">=10"),
            CellValue.FromString("North"),
            CellValue.FromString("North"),
            CellValue.FromNumber(100),
            CellValue.FromString(">50"),
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
            CellValue.FromString("text"),
            CellValue.FromNumber(10),
            CellValue.FromString(">5"),
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
            CellValue.FromNumber(100),
            CellValue.FromNumber(10),
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
            CellValue.FromNumber(100),
            CellValue.FromNumber(10),
            CellValue.FromString(">5"),
            CellValue.FromString("Yes"),
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
            CellValue.Error("#N/A"),
            CellValue.FromNumber(10),
            CellValue.FromString(">5"),
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
            CellValue.FromNumber(100),
            CellValue.FromString("North"),
            CellValue.FromString("North"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(100.0, result.NumericValue);
    }

    [Fact]
    public void MaxIfs_SingleCriteriaNotMet_ReturnsZero()
    {
        var func = MaxIfsFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(100),
            CellValue.FromString("South"),
            CellValue.FromString("North"),
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
            CellValue.FromNumber(250),
            CellValue.FromString("North"),
            CellValue.FromString("North"),
            CellValue.FromNumber(120),
            CellValue.FromString(">100"),
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
            CellValue.FromNumber(250),
            CellValue.FromString("North"),
            CellValue.FromString("North"),
            CellValue.FromNumber(80),
            CellValue.FromString(">100"),
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
            CellValue.FromNumber(500),
            CellValue.FromString("North"),
            CellValue.FromString("North"),
            CellValue.FromNumber(120),
            CellValue.FromString(">100"),
            CellValue.FromString("Q1"),
            CellValue.FromString("Q1"),
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
            CellValue.FromString("text"),
            CellValue.FromString("North"),
            CellValue.FromString("North"),
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
            CellValue.FromNumber(100),
            CellValue.FromString("North"),
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
            CellValue.FromNumber(100),
            CellValue.FromString("North"),
            CellValue.FromString("North"),
            CellValue.FromNumber(120),
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
            CellValue.Error("#DIV/0!"),
            CellValue.FromString("North"),
            CellValue.FromString("North"),
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
            CellValue.FromNumber(50),
            CellValue.FromString("South"),
            CellValue.FromString("South"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(50.0, result.NumericValue);
    }

    [Fact]
    public void MinIfs_SingleCriteriaNotMet_ReturnsZero()
    {
        var func = MinIfsFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(50),
            CellValue.FromString("North"),
            CellValue.FromString("South"),
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
            CellValue.FromNumber(30),
            CellValue.FromString("South"),
            CellValue.FromString("South"),
            CellValue.FromNumber(25),
            CellValue.FromString("<50"),
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
            CellValue.FromNumber(30),
            CellValue.FromString("South"),
            CellValue.FromString("South"),
            CellValue.FromNumber(60),
            CellValue.FromString("<50"),
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
            CellValue.FromNumber(25),
            CellValue.FromString("South"),
            CellValue.FromString("South"),
            CellValue.FromNumber(30),
            CellValue.FromString("<50"),
            CellValue.FromString("Q2"),
            CellValue.FromString("Q2"),
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
            CellValue.FromString("text"),
            CellValue.FromString("South"),
            CellValue.FromString("South"),
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
            CellValue.FromNumber(50),
            CellValue.FromString("South"),
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
            CellValue.FromNumber(50),
            CellValue.FromString("South"),
            CellValue.FromString("South"),
            CellValue.FromNumber(30),
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
            CellValue.Error("#REF!"),
            CellValue.FromString("South"),
            CellValue.FromString("South"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    #endregion
}
