// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Sample tests for mathematical functions.
/// This demonstrates the testing pattern for all 50 implemented functions.
/// Additional function tests should follow this pattern.
/// </summary>
public class MathFunctionTests
{
    [Fact]
    public void Count_Numbers_ReturnsCount()
    {
        var func = CountFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(2),
            CellValue.FromString("text"),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(3.0, result.NumericValue);
    }

    [Fact]
    public void Count_ErrorValue_PropagatesError()
    {
        var func = CountFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void Max_MultipleValues_ReturnsMaximum()
    {
        var func = MaxFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(10),
            CellValue.FromNumber(3),
            CellValue.FromNumber(8),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void Max_NoNumericValues_ReturnsZero()
    {
        var func = MaxFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("text"),
            CellValue.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void Max_NegativeNumbers_ReturnsCorrectMax()
    {
        var func = MaxFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-5),
            CellValue.FromNumber(-10),
            CellValue.FromNumber(-3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(-3.0, result.NumericValue);
    }

    [Fact]
    public void Min_MultipleValues_ReturnsMinimum()
    {
        var func = MinFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(10),
            CellValue.FromNumber(3),
            CellValue.FromNumber(8),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(3.0, result.NumericValue);
    }

    [Fact]
    public void Min_NoNumericValues_ReturnsZero()
    {
        var func = MinFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("text"),
            CellValue.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void Round_ExcelRounding_RoundsCorrectly()
    {
        var func = RoundFunction.Instance;

        // 2.5 rounds to 3 (Excel's rounding - rounds half away from zero)
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(2.5),
            CellValue.FromNumber(0),
        });

        Assert.Equal(3.0, result1.NumericValue);

        // 3.5 rounds to 4 (Excel's rounding - rounds half away from zero)
        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(3.5),
            CellValue.FromNumber(0),
        });

        Assert.Equal(4.0, result2.NumericValue);

        // -2.5 rounds to -3 (Excel's rounding - rounds half away from zero)
        var result3 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-2.5),
            CellValue.FromNumber(0),
        });

        Assert.Equal(-3.0, result3.NumericValue);
    }

    [Fact]
    public void Round_DecimalPlaces_RoundsCorrectly()
    {
        var func = RoundFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(3.14159),
            CellValue.FromNumber(2),
        });

        Assert.Equal(3.14, result.NumericValue);
    }

    [Fact]
    public void Round_InvalidArguments_ReturnsError()
    {
        var func = RoundFunction.Instance;

        // Wrong number of arguments
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(2.5),
        });

        Assert.True(result1.IsError);
        Assert.Equal("#VALUE!", result1.ErrorValue);

        // Non-numeric first argument
        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromString("text"),
            CellValue.FromNumber(0),
        });

        Assert.True(result2.IsError);
        Assert.Equal("#VALUE!", result2.ErrorValue);
    }

    [Fact]
    public void Abs_PositiveNumber_ReturnsSameValue()
    {
        var func = AbsFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(5.5),
        });

        Assert.Equal(5.5, result.NumericValue);
    }

    [Fact]
    public void Abs_NegativeNumber_ReturnsPositiveValue()
    {
        var func = AbsFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-5.5),
        });

        Assert.Equal(5.5, result.NumericValue);
    }

    [Fact]
    public void Abs_Zero_ReturnsZero()
    {
        var func = AbsFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
        });

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void Abs_InvalidArguments_ReturnsError()
    {
        var func = AbsFunction.Instance;

        // Wrong number of arguments
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(10),
        });

        Assert.True(result1.IsError);

        // Non-numeric argument
        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromString("text"),
        });

        Assert.True(result2.IsError);
    }

    [Fact]
    public void Sum_MultipleNumbers_ReturnsSum()
    {
        var func = SumFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(60.0, result.NumericValue);
    }

    [Fact]
    public void Sum_MixedTypes_SumsOnlyNumbers()
    {
        var func = SumFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromString("text"),
            CellValue.FromNumber(20),
            CellValue.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(30.0, result.NumericValue);
    }

    [Fact]
    public void Average_Numbers_ReturnsAverage()
    {
        var func = AverageFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(20.0, result.NumericValue);
    }

    [Fact]
    public void Average_NoNumbers_ReturnsError()
    {
        var func = AverageFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("text"),
            CellValue.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }
}
