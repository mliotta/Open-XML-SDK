// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for newly implemented mathematical functions:
/// SUMPRODUCT, RAND, RANDBETWEEN, FACT, GCD, LCM, EVEN, ODD.
/// </summary>
public class NewMathFunctionTests
{
    #region SUMPRODUCT Tests

    [Fact]
    public void SumProduct_TwoNumbers_ReturnsProduct()
    {
        var func = SumProductFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(2),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(6.0, result.NumericValue);
    }

    [Fact]
    public void SumProduct_ThreeNumbers_ReturnsProduct()
    {
        var func = SumProductFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(2),
            CellValue.FromNumber(3),
            CellValue.FromNumber(4),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(24.0, result.NumericValue); // 2*3*4 = 24
    }

    [Fact]
    public void SumProduct_NoArguments_ReturnsError()
    {
        var func = SumProductFunction.Instance;

        var result = func.Execute(null!, Array.Empty<CellValue>());

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void SumProduct_ErrorValue_PropagatesError()
    {
        var func = SumProductFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(2),
            CellValue.Error("#DIV/0!"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void SumProduct_NonNumericValue_ReturnsError()
    {
        var func = SumProductFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(2),
            CellValue.FromString("text"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region RAND Tests

    [Fact]
    public void Rand_NoArguments_ReturnsBetweenZeroAndOne()
    {
        var func = RandFunction.Instance;

        // Call multiple times to check range
        for (int i = 0; i < 10; i++)
        {
            var result = func.Execute(null!, Array.Empty<CellValue>());

            Assert.Equal(CellValueType.Number, result.Type);
            Assert.True(result.NumericValue >= 0.0);
            Assert.True(result.NumericValue < 1.0);
        }
    }

    [Fact]
    public void Rand_WithArguments_ReturnsError()
    {
        var func = RandFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1),
        });

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region RANDBETWEEN Tests

    [Fact]
    public void RandBetween_ValidRange_ReturnsIntegerInRange()
    {
        var func = RandBetweenFunction.Instance;

        // Call multiple times to check range
        for (int i = 0; i < 10; i++)
        {
            var result = func.Execute(null!, new[]
            {
                CellValue.FromNumber(1),
                CellValue.FromNumber(10),
            });

            Assert.Equal(CellValueType.Number, result.Type);
            Assert.True(result.NumericValue >= 1.0);
            Assert.True(result.NumericValue <= 10.0);
            Assert.Equal(System.Math.Floor(result.NumericValue), result.NumericValue); // Should be integer
        }
    }

    [Fact]
    public void RandBetween_SameValues_ReturnsThatValue()
    {
        var func = RandBetweenFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(5),
        });

        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void RandBetween_BottomGreaterThanTop_ReturnsError()
    {
        var func = RandBetweenFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(1),
        });

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void RandBetween_WrongNumberOfArguments_ReturnsError()
    {
        var func = RandBetweenFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1),
        });

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void RandBetween_NonNumericArgument_ReturnsError()
    {
        var func = RandBetweenFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromString("text"),
            CellValue.FromNumber(10),
        });

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void RandBetween_ErrorValue_PropagatesError()
    {
        var func = RandBetweenFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(10),
        });

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    #endregion

    #region FACT Tests

    [Fact]
    public void Fact_Zero_ReturnsOne()
    {
        var func = FactFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
        });

        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void Fact_Five_Returns120()
    {
        var func = FactFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(5),
        });

        Assert.Equal(120.0, result.NumericValue); // 5! = 5*4*3*2*1 = 120
    }

    [Fact]
    public void Fact_Ten_ReturnsCorrectValue()
    {
        var func = FactFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(10),
        });

        Assert.Equal(3628800.0, result.NumericValue); // 10!
    }

    [Fact]
    public void Fact_NegativeNumber_ReturnsError()
    {
        var func = FactFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-5),
        });

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Fact_DecimalNumber_TruncatesToInteger()
    {
        var func = FactFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(5.9),
        });

        Assert.Equal(120.0, result.NumericValue); // 5! = 120
    }

    [Fact]
    public void Fact_WrongNumberOfArguments_ReturnsError()
    {
        var func = FactFunction.Instance;

        var result = func.Execute(null!, Array.Empty<CellValue>());

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Fact_NonNumericArgument_ReturnsError()
    {
        var func = FactFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromString("text"),
        });

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Fact_ErrorValue_PropagatesError()
    {
        var func = FactFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.Error("#DIV/0!"),
        });

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    #endregion

    #region GCD Tests

    [Fact]
    public void Gcd_TwoNumbers_ReturnsGcd()
    {
        var func = GcdFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(24),
            CellValue.FromNumber(36),
        });

        Assert.Equal(12.0, result.NumericValue);
    }

    [Fact]
    public void Gcd_ThreeNumbers_ReturnsGcd()
    {
        var func = GcdFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(12),
            CellValue.FromNumber(18),
            CellValue.FromNumber(24),
        });

        Assert.Equal(6.0, result.NumericValue);
    }

    [Fact]
    public void Gcd_CoprimesNumbers_ReturnsOne()
    {
        var func = GcdFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(7),
            CellValue.FromNumber(13),
        });

        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void Gcd_WithZero_ReturnsOtherNumber()
    {
        var func = GcdFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
            CellValue.FromNumber(12),
        });

        Assert.Equal(12.0, result.NumericValue);
    }

    [Fact]
    public void Gcd_NoArguments_ReturnsError()
    {
        var func = GcdFunction.Instance;

        var result = func.Execute(null!, Array.Empty<CellValue>());

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Gcd_NonNumericArgument_ReturnsError()
    {
        var func = GcdFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(24),
            CellValue.FromString("text"),
        });

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Gcd_ErrorValue_PropagatesError()
    {
        var func = GcdFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(24),
            CellValue.Error("#DIV/0!"),
        });

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    #endregion

    #region LCM Tests

    [Fact]
    public void Lcm_TwoNumbers_ReturnsLcm()
    {
        var func = LcmFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(4),
            CellValue.FromNumber(6),
        });

        Assert.Equal(12.0, result.NumericValue);
    }

    [Fact]
    public void Lcm_ThreeNumbers_ReturnsLcm()
    {
        var func = LcmFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(2),
            CellValue.FromNumber(3),
            CellValue.FromNumber(4),
        });

        Assert.Equal(12.0, result.NumericValue);
    }

    [Fact]
    public void Lcm_SameNumbers_ReturnsThatNumber()
    {
        var func = LcmFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(5),
        });

        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void Lcm_WithOne_ReturnsOtherNumber()
    {
        var func = LcmFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(12),
        });

        Assert.Equal(12.0, result.NumericValue);
    }

    [Fact]
    public void Lcm_WithZero_ReturnsZero()
    {
        var func = LcmFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
            CellValue.FromNumber(12),
        });

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void Lcm_NoArguments_ReturnsError()
    {
        var func = LcmFunction.Instance;

        var result = func.Execute(null!, Array.Empty<CellValue>());

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Lcm_NonNumericArgument_ReturnsError()
    {
        var func = LcmFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(4),
            CellValue.FromString("text"),
        });

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Lcm_ErrorValue_PropagatesError()
    {
        var func = LcmFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(4),
            CellValue.Error("#DIV/0!"),
        });

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    #endregion

    #region EVEN Tests

    [Fact]
    public void Even_PositiveOddNumber_RoundsUp()
    {
        var func = EvenFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(3),
        });

        Assert.Equal(4.0, result.NumericValue);
    }

    [Fact]
    public void Even_PositiveEvenNumber_ReturnsSame()
    {
        var func = EvenFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(4),
        });

        Assert.Equal(4.0, result.NumericValue);
    }

    [Fact]
    public void Even_PositiveDecimal_RoundsUp()
    {
        var func = EvenFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(3.1),
        });

        Assert.Equal(4.0, result.NumericValue);
    }

    [Fact]
    public void Even_NegativeOddNumber_RoundsDown()
    {
        var func = EvenFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-3),
        });

        Assert.Equal(-4.0, result.NumericValue);
    }

    [Fact]
    public void Even_NegativeEvenNumber_ReturnsSame()
    {
        var func = EvenFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-4),
        });

        Assert.Equal(-4.0, result.NumericValue);
    }

    [Fact]
    public void Even_Zero_ReturnsZero()
    {
        var func = EvenFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
        });

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void Even_WrongNumberOfArguments_ReturnsError()
    {
        var func = EvenFunction.Instance;

        var result = func.Execute(null!, Array.Empty<CellValue>());

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Even_NonNumericArgument_ReturnsError()
    {
        var func = EvenFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromString("text"),
        });

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Even_ErrorValue_PropagatesError()
    {
        var func = EvenFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.Error("#DIV/0!"),
        });

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    #endregion

    #region ODD Tests

    [Fact]
    public void Odd_PositiveEvenNumber_RoundsUp()
    {
        var func = OddFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(2),
        });

        Assert.Equal(3.0, result.NumericValue);
    }

    [Fact]
    public void Odd_PositiveOddNumber_ReturnsSame()
    {
        var func = OddFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(3),
        });

        Assert.Equal(3.0, result.NumericValue);
    }

    [Fact]
    public void Odd_PositiveDecimal_RoundsUp()
    {
        var func = OddFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(2.1),
        });

        Assert.Equal(3.0, result.NumericValue);
    }

    [Fact]
    public void Odd_NegativeEvenNumber_RoundsDown()
    {
        var func = OddFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-2),
        });

        Assert.Equal(-3.0, result.NumericValue);
    }

    [Fact]
    public void Odd_NegativeOddNumber_ReturnsSame()
    {
        var func = OddFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-3),
        });

        Assert.Equal(-3.0, result.NumericValue);
    }

    [Fact]
    public void Odd_Zero_ReturnsOne()
    {
        var func = OddFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
        });

        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void Odd_WrongNumberOfArguments_ReturnsError()
    {
        var func = OddFunction.Instance;

        var result = func.Execute(null!, Array.Empty<CellValue>());

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Odd_NonNumericArgument_ReturnsError()
    {
        var func = OddFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromString("text"),
        });

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Odd_ErrorValue_PropagatesError()
    {
        var func = OddFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.Error("#DIV/0!"),
        });

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    #endregion
}
