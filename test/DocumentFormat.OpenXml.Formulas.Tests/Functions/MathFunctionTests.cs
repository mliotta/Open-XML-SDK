// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
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

    [Fact]
    public void Sqrt_PositiveNumber_ReturnsSquareRoot()
    {
        var func = SqrtFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(16),
        });

        Assert.Equal(4.0, result.NumericValue);
    }

    [Fact]
    public void Sqrt_Zero_ReturnsZero()
    {
        var func = SqrtFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
        });

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void Sqrt_NegativeNumber_ReturnsError()
    {
        var func = SqrtFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-4),
        });

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Sqrt_InvalidArguments_ReturnsError()
    {
        var func = SqrtFunction.Instance;

        // Wrong number of arguments
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(16),
            CellValue.FromNumber(2),
        });

        Assert.True(result1.IsError);
        Assert.Equal("#VALUE!", result1.ErrorValue);

        // Non-numeric argument
        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromString("text"),
        });

        Assert.True(result2.IsError);
        Assert.Equal("#VALUE!", result2.ErrorValue);
    }

    [Fact]
    public void Mod_PositiveNumbers_ReturnsRemainder()
    {
        var func = ModFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(3),
        });

        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void Mod_NegativeNumber_ReturnsCorrectRemainder()
    {
        var func = ModFunction.Instance;

        // Excel MOD behavior: MOD(-10, 3) = 2
        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-10),
            CellValue.FromNumber(3),
        });

        Assert.Equal(2.0, result.NumericValue);
    }

    [Fact]
    public void Mod_ZeroDivisor_ReturnsError()
    {
        var func = ModFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(0),
        });

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void Mod_InvalidArguments_ReturnsError()
    {
        var func = ModFunction.Instance;

        // Wrong number of arguments
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(10),
        });

        Assert.True(result1.IsError);
        Assert.Equal("#VALUE!", result1.ErrorValue);

        // Non-numeric argument
        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromString("text"),
            CellValue.FromNumber(3),
        });

        Assert.True(result2.IsError);
        Assert.Equal("#VALUE!", result2.ErrorValue);
    }

    [Fact]
    public void Int_PositiveNumber_RoundsDown()
    {
        var func = IntFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(8.9),
        });

        Assert.Equal(8.0, result.NumericValue);
    }

    [Fact]
    public void Int_NegativeNumber_RoundsDown()
    {
        var func = IntFunction.Instance;

        // INT(-8.9) = -9 (rounds down, not toward zero)
        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-8.9),
        });

        Assert.Equal(-9.0, result.NumericValue);
    }

    [Fact]
    public void Int_Zero_ReturnsZero()
    {
        var func = IntFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
        });

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void Int_InvalidArguments_ReturnsError()
    {
        var func = IntFunction.Instance;

        // Wrong number of arguments
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(8.9),
            CellValue.FromNumber(1),
        });

        Assert.True(result1.IsError);
        Assert.Equal("#VALUE!", result1.ErrorValue);

        // Non-numeric argument
        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromString("text"),
        });

        Assert.True(result2.IsError);
        Assert.Equal("#VALUE!", result2.ErrorValue);
    }

    [Fact]
    public void Ceiling_PositiveNumbers_RoundsUp()
    {
        var func = CeilingFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(4.3),
            CellValue.FromNumber(1),
        });

        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void Ceiling_SignificanceGreaterThanOne_RoundsCorrectly()
    {
        var func = CeilingFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(22),
            CellValue.FromNumber(10),
        });

        Assert.Equal(30.0, result.NumericValue);
    }

    [Fact]
    public void Ceiling_NegativeNumbers_RoundsCorrectly()
    {
        var func = CeilingFunction.Instance;

        // CEILING(-4.3, -1) = -4 (rounds toward zero when both are negative)
        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-4.3),
            CellValue.FromNumber(-1),
        });

        Assert.Equal(-4.0, result.NumericValue);
    }

    [Fact]
    public void Ceiling_MixedSigns_ReturnsError()
    {
        var func = CeilingFunction.Instance;

        // Different signs should return #NUM!
        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(4.3),
            CellValue.FromNumber(-1),
        });

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Ceiling_ZeroSignificance_ReturnsZero()
    {
        var func = CeilingFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(4.3),
            CellValue.FromNumber(0),
        });

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void Ceiling_InvalidArguments_ReturnsError()
    {
        var func = CeilingFunction.Instance;

        // Wrong number of arguments
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(4.3),
        });

        Assert.True(result1.IsError);
        Assert.Equal("#VALUE!", result1.ErrorValue);
    }

    [Fact]
    public void Floor_PositiveNumbers_RoundsDown()
    {
        var func = FloorFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(4.7),
            CellValue.FromNumber(1),
        });

        Assert.Equal(4.0, result.NumericValue);
    }

    [Fact]
    public void Floor_SignificanceGreaterThanOne_RoundsCorrectly()
    {
        var func = FloorFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(22),
            CellValue.FromNumber(10),
        });

        Assert.Equal(20.0, result.NumericValue);
    }

    [Fact]
    public void Floor_NegativeNumbers_RoundsCorrectly()
    {
        var func = FloorFunction.Instance;

        // FLOOR(-4.7, -1) = -5 (rounds away from zero when both are negative)
        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-4.7),
            CellValue.FromNumber(-1),
        });

        Assert.Equal(-5.0, result.NumericValue);
    }

    [Fact]
    public void Floor_MixedSigns_ReturnsError()
    {
        var func = FloorFunction.Instance;

        // Different signs should return #NUM!
        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(4.7),
            CellValue.FromNumber(-1),
        });

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Floor_ZeroSignificance_ReturnsZero()
    {
        var func = FloorFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(4.7),
            CellValue.FromNumber(0),
        });

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void Floor_InvalidArguments_ReturnsError()
    {
        var func = FloorFunction.Instance;

        // Wrong number of arguments
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(4.7),
        });

        Assert.True(result1.IsError);
        Assert.Equal("#VALUE!", result1.ErrorValue);
    }

    [Fact]
    public void Trunc_NoDigits_TruncatesToInteger()
    {
        var func = TruncFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(8.9),
        });

        Assert.Equal(8.0, result.NumericValue);
    }

    [Fact]
    public void Trunc_WithDigits_TruncatesToPrecision()
    {
        var func = TruncFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(8.987654),
            CellValue.FromNumber(2),
        });

        Assert.Equal(8.98, result.NumericValue);
    }

    [Fact]
    public void Trunc_NegativeNumber_TruncatesCorrectly()
    {
        var func = TruncFunction.Instance;

        // TRUNC(-8.9) = -8 (truncates toward zero, not down)
        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-8.9),
        });

        Assert.Equal(-8.0, result.NumericValue);
    }

    [Fact]
    public void Trunc_NegativeDigits_TruncatesLeftOfDecimal()
    {
        var func = TruncFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1234.567),
            CellValue.FromNumber(-2),
        });

        Assert.Equal(1200.0, result.NumericValue);
    }

    [Fact]
    public void Trunc_Zero_ReturnsZero()
    {
        var func = TruncFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
        });

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void Trunc_InvalidArguments_ReturnsError()
    {
        var func = TruncFunction.Instance;

        // Too many arguments
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(8.9),
            CellValue.FromNumber(2),
            CellValue.FromNumber(3),
        });

        Assert.True(result1.IsError);
        Assert.Equal("#VALUE!", result1.ErrorValue);

        // Non-numeric first argument
        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromString("text"),
        });

        Assert.True(result2.IsError);
        Assert.Equal("#VALUE!", result2.ErrorValue);

        // Non-numeric second argument
        var result3 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(8.9),
            CellValue.FromString("text"),
        });

        Assert.True(result3.IsError);
        Assert.Equal("#VALUE!", result3.ErrorValue);
    }

    [Fact]
    public void Sign_PositiveNumber_ReturnsOne()
    {
        var func = SignFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(5.5),
        });

        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void Sign_NegativeNumber_ReturnsNegativeOne()
    {
        var func = SignFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-5.5),
        });

        Assert.Equal(-1.0, result.NumericValue);
    }

    [Fact]
    public void Sign_Zero_ReturnsZero()
    {
        var func = SignFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
        });

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void Sign_InvalidArguments_ReturnsError()
    {
        var func = SignFunction.Instance;

        // Wrong number of arguments
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(10),
        });

        Assert.True(result1.IsError);
        Assert.Equal("#VALUE!", result1.ErrorValue);

        // Non-numeric argument
        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromString("text"),
        });

        Assert.True(result2.IsError);
        Assert.Equal("#VALUE!", result2.ErrorValue);
    }

    [Fact]
    public void Exp_ReturnsCorrectValue()
    {
        var func = ExpFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1),
        });

        Assert.Equal(System.Math.E, result.NumericValue, 10);
    }

    [Fact]
    public void Exp_Zero_ReturnsOne()
    {
        var func = ExpFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
        });

        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void Exp_NegativeNumber_ReturnsCorrectValue()
    {
        var func = ExpFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-1),
        });

        Assert.Equal(1.0 / System.Math.E, result.NumericValue, 10);
    }

    [Fact]
    public void Exp_InvalidArguments_ReturnsError()
    {
        var func = ExpFunction.Instance;

        // Wrong number of arguments
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(2),
        });

        Assert.True(result1.IsError);
        Assert.Equal("#VALUE!", result1.ErrorValue);

        // Non-numeric argument
        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromString("text"),
        });

        Assert.True(result2.IsError);
        Assert.Equal("#VALUE!", result2.ErrorValue);
    }

    [Fact]
    public void Ln_ReturnsCorrectValue()
    {
        var func = LnFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(System.Math.E),
        });

        Assert.Equal(1.0, result.NumericValue, 10);
    }

    [Fact]
    public void Ln_One_ReturnsZero()
    {
        var func = LnFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1),
        });

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void Ln_NegativeNumber_ReturnsError()
    {
        var func = LnFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-1),
        });

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Ln_Zero_ReturnsError()
    {
        var func = LnFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
        });

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Ln_InvalidArguments_ReturnsError()
    {
        var func = LnFunction.Instance;

        // Wrong number of arguments
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(2),
        });

        Assert.True(result1.IsError);
        Assert.Equal("#VALUE!", result1.ErrorValue);

        // Non-numeric argument
        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromString("text"),
        });

        Assert.True(result2.IsError);
        Assert.Equal("#VALUE!", result2.ErrorValue);
    }

    [Fact]
    public void Log_DefaultBase10_ReturnsCorrectValue()
    {
        var func = LogFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(100),
        });

        Assert.Equal(2.0, result.NumericValue);
    }

    [Fact]
    public void Log_CustomBase_ReturnsCorrectValue()
    {
        var func = LogFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(8),
            CellValue.FromNumber(2),
        });

        Assert.Equal(3.0, result.NumericValue);
    }

    [Fact]
    public void Log_Base10_1000_Returns3()
    {
        var func = LogFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1000),
            CellValue.FromNumber(10),
        });

        Assert.Equal(3.0, result.NumericValue, 10);
    }

    [Fact]
    public void Log_NegativeNumber_ReturnsError()
    {
        var func = LogFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-10),
        });

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Log_InvalidBase_ReturnsError()
    {
        var func = LogFunction.Instance;

        // Base <= 0
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(-1),
        });

        Assert.True(result1.IsError);
        Assert.Equal("#NUM!", result1.ErrorValue);

        // Base = 1
        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(1),
        });

        Assert.True(result2.IsError);
        Assert.Equal("#NUM!", result2.ErrorValue);
    }

    [Fact]
    public void Log_InvalidArguments_ReturnsError()
    {
        var func = LogFunction.Instance;

        // No arguments
        var result1 = func.Execute(null!, Array.Empty<CellValue>());

        Assert.True(result1.IsError);
        Assert.Equal("#VALUE!", result1.ErrorValue);

        // Non-numeric argument
        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromString("text"),
        });

        Assert.True(result2.IsError);
        Assert.Equal("#VALUE!", result2.ErrorValue);
    }

    [Fact]
    public void Log10_ReturnsCorrectValue()
    {
        var func = Log10Function.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1000),
        });

        Assert.Equal(3.0, result.NumericValue);
    }

    [Fact]
    public void Log10_100_Returns2()
    {
        var func = Log10Function.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(100),
        });

        Assert.Equal(2.0, result.NumericValue);
    }

    [Fact]
    public void Log10_One_ReturnsZero()
    {
        var func = Log10Function.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1),
        });

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void Log10_NegativeNumber_ReturnsError()
    {
        var func = Log10Function.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-10),
        });

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Log10_Zero_ReturnsError()
    {
        var func = Log10Function.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
        });

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Log10_InvalidArguments_ReturnsError()
    {
        var func = Log10Function.Instance;

        // Wrong number of arguments
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(100),
            CellValue.FromNumber(10),
        });

        Assert.True(result1.IsError);
        Assert.Equal("#VALUE!", result1.ErrorValue);

        // Non-numeric argument
        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromString("text"),
        });

        Assert.True(result2.IsError);
        Assert.Equal("#VALUE!", result2.ErrorValue);
    }

    [Fact]
    public void Pi_ReturnsCorrectValue()
    {
        var func = PiFunction.Instance;

        var result = func.Execute(null!, Array.Empty<CellValue>());

        Assert.Equal(System.Math.PI, result.NumericValue);
    }

    [Fact]
    public void Pi_InvalidArguments_ReturnsError()
    {
        var func = PiFunction.Instance;

        // With arguments
        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1),
        });

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Radians_180_ReturnsPi()
    {
        var func = RadiansFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(180),
        });

        Assert.Equal(System.Math.PI, result.NumericValue, 10);
    }

    [Fact]
    public void Radians_90_ReturnsPiOver2()
    {
        var func = RadiansFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(90),
        });

        Assert.Equal(System.Math.PI / 2, result.NumericValue, 10);
    }

    [Fact]
    public void Radians_Zero_ReturnsZero()
    {
        var func = RadiansFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
        });

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void Radians_InvalidArguments_ReturnsError()
    {
        var func = RadiansFunction.Instance;

        // Wrong number of arguments
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(180),
            CellValue.FromNumber(2),
        });

        Assert.True(result1.IsError);
        Assert.Equal("#VALUE!", result1.ErrorValue);

        // Non-numeric argument
        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromString("text"),
        });

        Assert.True(result2.IsError);
        Assert.Equal("#VALUE!", result2.ErrorValue);
    }

    [Fact]
    public void Degrees_Pi_Returns180()
    {
        var func = DegreesFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(System.Math.PI),
        });

        Assert.Equal(180.0, result.NumericValue, 10);
    }

    [Fact]
    public void Degrees_PiOver2_Returns90()
    {
        var func = DegreesFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(System.Math.PI / 2),
        });

        Assert.Equal(90.0, result.NumericValue, 10);
    }

    [Fact]
    public void Degrees_Zero_ReturnsZero()
    {
        var func = DegreesFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
        });

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void Degrees_InvalidArguments_ReturnsError()
    {
        var func = DegreesFunction.Instance;

        // Wrong number of arguments
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(System.Math.PI),
            CellValue.FromNumber(2),
        });

        Assert.True(result1.IsError);
        Assert.Equal("#VALUE!", result1.ErrorValue);

        // Non-numeric argument
        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromString("text"),
        });

        Assert.True(result2.IsError);
        Assert.Equal("#VALUE!", result2.ErrorValue);
    }

    [Fact]
    public void Sin_PiOver2_ReturnsOne()
    {
        var func = SinFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(System.Math.PI / 2),
        });

        Assert.Equal(1.0, result.NumericValue, 10);
    }

    [Fact]
    public void Sin_Zero_ReturnsZero()
    {
        var func = SinFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
        });

        Assert.Equal(0.0, result.NumericValue, 10);
    }

    [Fact]
    public void Sin_Pi_ReturnsZero()
    {
        var func = SinFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(System.Math.PI),
        });

        Assert.Equal(0.0, result.NumericValue, 10);
    }

    [Fact]
    public void Sin_InvalidArguments_ReturnsError()
    {
        var func = SinFunction.Instance;

        // Wrong number of arguments
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(System.Math.PI),
            CellValue.FromNumber(2),
        });

        Assert.True(result1.IsError);
        Assert.Equal("#VALUE!", result1.ErrorValue);

        // Non-numeric argument
        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromString("text"),
        });

        Assert.True(result2.IsError);
        Assert.Equal("#VALUE!", result2.ErrorValue);
    }

    [Fact]
    public void Cos_Zero_ReturnsOne()
    {
        var func = CosFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
        });

        Assert.Equal(1.0, result.NumericValue, 10);
    }

    [Fact]
    public void Cos_Pi_ReturnsNegativeOne()
    {
        var func = CosFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(System.Math.PI),
        });

        Assert.Equal(-1.0, result.NumericValue, 10);
    }

    [Fact]
    public void Cos_PiOver2_ReturnsZero()
    {
        var func = CosFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(System.Math.PI / 2),
        });

        Assert.Equal(0.0, result.NumericValue, 10);
    }

    [Fact]
    public void Cos_InvalidArguments_ReturnsError()
    {
        var func = CosFunction.Instance;

        // Wrong number of arguments
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(System.Math.PI),
            CellValue.FromNumber(2),
        });

        Assert.True(result1.IsError);
        Assert.Equal("#VALUE!", result1.ErrorValue);

        // Non-numeric argument
        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromString("text"),
        });

        Assert.True(result2.IsError);
        Assert.Equal("#VALUE!", result2.ErrorValue);
    }

    [Fact]
    public void Tan_PiOver4_ReturnsOne()
    {
        var func = TanFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(System.Math.PI / 4),
        });

        Assert.Equal(1.0, result.NumericValue, 10);
    }

    [Fact]
    public void Tan_Zero_ReturnsZero()
    {
        var func = TanFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
        });

        Assert.Equal(0.0, result.NumericValue, 10);
    }

    [Fact]
    public void Tan_InvalidArguments_ReturnsError()
    {
        var func = TanFunction.Instance;

        // Wrong number of arguments
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(System.Math.PI),
            CellValue.FromNumber(2),
        });

        Assert.True(result1.IsError);
        Assert.Equal("#VALUE!", result1.ErrorValue);

        // Non-numeric argument
        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromString("text"),
        });

        Assert.True(result2.IsError);
        Assert.Equal("#VALUE!", result2.ErrorValue);
    }
}
