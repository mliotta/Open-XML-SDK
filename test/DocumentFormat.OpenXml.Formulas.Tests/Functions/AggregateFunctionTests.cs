// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for aggregate functions including *A variants, SUBTOTAL, and AGGREGATE.
/// </summary>
public class AggregateFunctionTests
{
    // AVERAGEA Tests
    [Fact]
    public void AverageA_NumbersOnly_ReturnsAverage()
    {
        var func = AverageAFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(2),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(2.0, result.NumericValue);
    }

    [Fact]
    public void AverageA_WithText_CountsTextAsZero()
    {
        var func = AverageAFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromString("text"),
            CellValue.FromNumber(20),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(10.0, result.NumericValue); // (10 + 0 + 20) / 3 = 10
    }

    [Fact]
    public void AverageA_WithBoolean_CountsTrueAsOne()
    {
        var func = AverageAFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(2),
            CellValue.FromBool(true),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(2.0, result.NumericValue); // (2 + 1 + 3) / 3 = 2
    }

    [Fact]
    public void AverageA_WithFalse_CountsFalseAsZero()
    {
        var func = AverageAFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(4),
            CellValue.FromBool(false),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(2.0, result.NumericValue); // (4 + 0 + 2) / 3 = 2
    }

    [Fact]
    public void AverageA_NoValues_ReturnsError()
    {
        var func = AverageAFunction.Instance;
        var args = new[]
        {
            CellValue.Empty,
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    // MINA Tests
    [Fact]
    public void MinA_NumbersOnly_ReturnsMinimum()
    {
        var func = MinAFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(2),
            CellValue.FromNumber(8),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(2.0, result.NumericValue);
    }

    [Fact]
    public void MinA_WithText_TextIsZero()
    {
        var func = MinAFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromString("text"),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void MinA_WithBoolean_TrueIsOne()
    {
        var func = MinAFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromBool(true),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue);
    }

    // MAXA Tests
    [Fact]
    public void MaxA_NumbersOnly_ReturnsMaximum()
    {
        var func = MaxAFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(2),
            CellValue.FromNumber(8),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(8.0, result.NumericValue);
    }

    [Fact]
    public void MaxA_WithText_TextIsZero()
    {
        var func = MaxAFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-5),
            CellValue.FromString("text"),
            CellValue.FromNumber(-2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.0, result.NumericValue);
    }

    // STDEVA Tests
    [Fact]
    public void StDevA_NumbersOnly_ReturnsStdDev()
    {
        var func = StDevAFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(2),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue, 10);
    }

    [Fact]
    public void StDevA_WithText_TextIsZero()
    {
        var func = StDevAFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(2),
            CellValue.FromString("text"),
            CellValue.FromNumber(4),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // Values: 2, 0, 4 -> mean=2, stdev=2.0
        Assert.Equal(2.0, result.NumericValue, 10);
    }

    [Fact]
    public void StDevA_LessThanTwo_ReturnsError()
    {
        var func = StDevAFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    // STDEVPA Tests
    [Fact]
    public void StDevPA_NumbersOnly_ReturnsPopStdDev()
    {
        var func = StDevPAFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(2),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.816496580927726, result.NumericValue, 10);
    }

    [Fact]
    public void StDevPA_SingleValue_ReturnsZero()
    {
        var func = StDevPAFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(0.0, result.NumericValue);
    }

    // VARA Tests
    [Fact]
    public void VarA_NumbersOnly_ReturnsVariance()
    {
        var func = VarAFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(2),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void VarA_WithBoolean_IncludesBoolean()
    {
        var func = VarAFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(2),
            CellValue.FromBool(true),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // Values: 2, 1, 3 -> variance = 1.0
        Assert.Equal(1.0, result.NumericValue, 10);
    }

    // VARPA Tests
    [Fact]
    public void VarPA_NumbersOnly_ReturnsPopVariance()
    {
        var func = VarPAFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(2),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.6666666666666666, result.NumericValue, 10);
    }

    // SUBTOTAL Tests
    [Fact]
    public void Subtotal_Average_ReturnsAverage()
    {
        var func = SubtotalFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1), // Function code 1 = AVERAGE
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(20.0, result.NumericValue);
    }

    [Fact]
    public void Subtotal_Sum_ReturnsSum()
    {
        var func = SubtotalFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(9), // Function code 9 = SUM
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(60.0, result.NumericValue);
    }

    [Fact]
    public void Subtotal_Max_ReturnsMax()
    {
        var func = SubtotalFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(4), // Function code 4 = MAX
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(30.0, result.NumericValue);
    }

    [Fact]
    public void Subtotal_IgnoreHidden_Average()
    {
        var func = SubtotalFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(101), // Function code 101 = AVERAGE (ignore hidden)
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(20.0, result.NumericValue);
    }

    [Fact]
    public void Subtotal_InvalidFunction_ReturnsError()
    {
        var func = SubtotalFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(99), // Invalid function code
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Subtotal_TooFewArgs_ReturnsError()
    {
        var func = SubtotalFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    // AGGREGATE Tests
    [Fact]
    public void Aggregate_Average_ReturnsAverage()
    {
        var func = AggregateFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1), // Function code 1 = AVERAGE
            CellValue.FromNumber(0), // Options
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(20.0, result.NumericValue);
    }

    [Fact]
    public void Aggregate_Sum_ReturnsSum()
    {
        var func = AggregateFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(9), // Function code 9 = SUM
            CellValue.FromNumber(0), // Options
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(60.0, result.NumericValue);
    }

    [Fact]
    public void Aggregate_WithErrors_IgnoreErrors()
    {
        var func = AggregateFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(9), // Function code 9 = SUM
            CellValue.FromNumber(2), // Options 2 = ignore errors
            CellValue.FromNumber(10),
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(20),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(30.0, result.NumericValue);
    }

    [Fact]
    public void Aggregate_WithErrors_NoIgnore_PropagatesError()
    {
        var func = AggregateFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(9), // Function code 9 = SUM
            CellValue.FromNumber(0), // Options 0 = don't ignore errors
            CellValue.FromNumber(10),
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(20),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void Aggregate_Median_ReturnsMedian()
    {
        var func = AggregateFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(12), // Function code 12 = MEDIAN
            CellValue.FromNumber(0), // Options
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(20.0, result.NumericValue);
    }

    [Fact]
    public void Aggregate_InvalidFunction_ReturnsError()
    {
        var func = AggregateFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(99), // Invalid function code
            CellValue.FromNumber(0),
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Aggregate_TooFewArgs_ReturnsError()
    {
        var func = AggregateFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Aggregate_InvalidOptions_ReturnsError()
    {
        var func = AggregateFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(99), // Invalid options
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }
}
