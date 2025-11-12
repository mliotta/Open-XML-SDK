// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;
using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for advanced trigonometric and hyperbolic functions.
/// </summary>
public class TrigonometricFunctionTests
{
    // ASIN Tests
    [Fact]
    public void Asin_ValidInput_ReturnsCorrectValue()
    {
        var func = AsinFunction.Instance;

        // ASIN(0.5) ≈ 0.5236 (30 degrees in radians)
        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0.5),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(System.Math.PI / 6, result.NumericValue, 10);
    }

    [Fact]
    public void Asin_Zero_ReturnsZero()
    {
        var func = AsinFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
        });

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void Asin_One_ReturnsPiOver2()
    {
        var func = AsinFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1),
        });

        Assert.Equal(System.Math.PI / 2, result.NumericValue, 10);
    }

    [Fact]
    public void Asin_NegativeOne_ReturnsNegativePiOver2()
    {
        var func = AsinFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-1),
        });

        Assert.Equal(-System.Math.PI / 2, result.NumericValue, 10);
    }

    [Fact]
    public void Asin_OutOfRange_ReturnsError()
    {
        var func = AsinFunction.Instance;

        // Greater than 1
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1.5),
        });

        Assert.True(result1.IsError);
        Assert.Equal("#NUM!", result1.ErrorValue);

        // Less than -1
        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-1.5),
        });

        Assert.True(result2.IsError);
        Assert.Equal("#NUM!", result2.ErrorValue);
    }

    [Fact]
    public void Asin_InvalidArguments_ReturnsError()
    {
        var func = AsinFunction.Instance;

        // Wrong number of arguments
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0.5),
            CellValue.FromNumber(0.5),
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
    public void Asin_ErrorInput_PropagatesError()
    {
        var func = AsinFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.Error("#DIV/0!"),
        });

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    // ACOS Tests
    [Fact]
    public void Acos_ValidInput_ReturnsCorrectValue()
    {
        var func = AcosFunction.Instance;

        // ACOS(0.5) ≈ 1.0472 (60 degrees in radians)
        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0.5),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(System.Math.PI / 3, result.NumericValue, 10);
    }

    [Fact]
    public void Acos_One_ReturnsZero()
    {
        var func = AcosFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1),
        });

        Assert.Equal(0.0, result.NumericValue, 10);
    }

    [Fact]
    public void Acos_Zero_ReturnsPiOver2()
    {
        var func = AcosFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
        });

        Assert.Equal(System.Math.PI / 2, result.NumericValue, 10);
    }

    [Fact]
    public void Acos_NegativeOne_ReturnsPi()
    {
        var func = AcosFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-1),
        });

        Assert.Equal(System.Math.PI, result.NumericValue, 10);
    }

    [Fact]
    public void Acos_OutOfRange_ReturnsError()
    {
        var func = AcosFunction.Instance;

        // Greater than 1
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1.5),
        });

        Assert.True(result1.IsError);
        Assert.Equal("#NUM!", result1.ErrorValue);

        // Less than -1
        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-1.5),
        });

        Assert.True(result2.IsError);
        Assert.Equal("#NUM!", result2.ErrorValue);
    }

    [Fact]
    public void Acos_InvalidArguments_ReturnsError()
    {
        var func = AcosFunction.Instance;

        // Wrong number of arguments
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0.5),
            CellValue.FromNumber(0.5),
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

    // ATAN Tests
    [Fact]
    public void Atan_ValidInput_ReturnsCorrectValue()
    {
        var func = AtanFunction.Instance;

        // ATAN(1) ≈ 0.7854 (45 degrees in radians)
        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(System.Math.PI / 4, result.NumericValue, 10);
    }

    [Fact]
    public void Atan_Zero_ReturnsZero()
    {
        var func = AtanFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
        });

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void Atan_NegativeValue_ReturnsNegativeResult()
    {
        var func = AtanFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-1),
        });

        Assert.Equal(-System.Math.PI / 4, result.NumericValue, 10);
    }

    [Fact]
    public void Atan_LargeValue_ReturnsCorrectValue()
    {
        var func = AtanFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1000),
        });

        // ATAN approaches PI/2 for large positive values
        Assert.True(result.NumericValue > 1.5 && result.NumericValue < System.Math.PI / 2);
    }

    [Fact]
    public void Atan_InvalidArguments_ReturnsError()
    {
        var func = AtanFunction.Instance;

        // Wrong number of arguments
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1),
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

    // ATAN2 Tests
    [Fact]
    public void Atan2_ValidInput_ReturnsCorrectValue()
    {
        var func = Atan2Function.Instance;

        // ATAN2(1, 1) ≈ 0.7854 (45 degrees in radians)
        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(1),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(System.Math.PI / 4, result.NumericValue, 10);
    }

    [Fact]
    public void Atan2_DifferentQuadrants_ReturnsCorrectAngles()
    {
        var func = Atan2Function.Instance;

        // Quadrant I
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(1),
        });
        Assert.True(result1.NumericValue > 0 && result1.NumericValue < System.Math.PI / 2);

        // Quadrant II
        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-1),
            CellValue.FromNumber(1),
        });
        Assert.True(result2.NumericValue > System.Math.PI / 2);

        // Quadrant III
        var result3 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-1),
            CellValue.FromNumber(-1),
        });
        Assert.True(result3.NumericValue < -System.Math.PI / 2);

        // Quadrant IV
        var result4 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(-1),
        });
        Assert.True(result4.NumericValue < 0 && result4.NumericValue > -System.Math.PI / 2);
    }

    [Fact]
    public void Atan2_BothZero_ReturnsError()
    {
        var func = Atan2Function.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
            CellValue.FromNumber(0),
        });

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void Atan2_InvalidArguments_ReturnsError()
    {
        var func = Atan2Function.Instance;

        // Wrong number of arguments
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1),
        });

        Assert.True(result1.IsError);
        Assert.Equal("#VALUE!", result1.ErrorValue);

        // Non-numeric argument
        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromString("text"),
            CellValue.FromNumber(1),
        });

        Assert.True(result2.IsError);
        Assert.Equal("#VALUE!", result2.ErrorValue);
    }

    // SINH Tests
    [Fact]
    public void Sinh_ValidInput_ReturnsCorrectValue()
    {
        var func = SinhFunction.Instance;

        // SINH(1) ≈ 1.1752
        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(System.Math.Sinh(1), result.NumericValue, 10);
    }

    [Fact]
    public void Sinh_Zero_ReturnsZero()
    {
        var func = SinhFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
        });

        Assert.Equal(0.0, result.NumericValue, 10);
    }

    [Fact]
    public void Sinh_NegativeValue_ReturnsNegativeResult()
    {
        var func = SinhFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-1),
        });

        Assert.Equal(-System.Math.Sinh(1), result.NumericValue, 10);
    }

    [Fact]
    public void Sinh_InvalidArguments_ReturnsError()
    {
        var func = SinhFunction.Instance;

        // Wrong number of arguments
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1),
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

    // COSH Tests
    [Fact]
    public void Cosh_ValidInput_ReturnsCorrectValue()
    {
        var func = CoshFunction.Instance;

        // COSH(1)
        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(System.Math.Cosh(1), result.NumericValue, 10);
    }

    [Fact]
    public void Cosh_Zero_ReturnsOne()
    {
        var func = CoshFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
        });

        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void Cosh_NegativeValue_ReturnsSameAsPositive()
    {
        var func = CoshFunction.Instance;

        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(2),
        });

        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-2),
        });

        Assert.Equal(result1.NumericValue, result2.NumericValue, 10);
    }

    [Fact]
    public void Cosh_InvalidArguments_ReturnsError()
    {
        var func = CoshFunction.Instance;

        // Wrong number of arguments
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1),
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

    // TANH Tests
    [Fact]
    public void Tanh_ValidInput_ReturnsCorrectValue()
    {
        var func = TanhFunction.Instance;

        // TANH(1)
        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(System.Math.Tanh(1), result.NumericValue, 10);
    }

    [Fact]
    public void Tanh_Zero_ReturnsZero()
    {
        var func = TanhFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
        });

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void Tanh_LargeValue_ApproachesOne()
    {
        var func = TanhFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(10),
        });

        // TANH approaches 1 for large positive values
        Assert.True(result.NumericValue > 0.99);
    }

    [Fact]
    public void Tanh_InvalidArguments_ReturnsError()
    {
        var func = TanhFunction.Instance;

        // Wrong number of arguments
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1),
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

    // ASINH Tests
    [Fact]
    public void Asinh_ValidInput_ReturnsCorrectValue()
    {
        var func = AsinhFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(System.Math.Asinh(1), result.NumericValue, 10);
    }

    [Fact]
    public void Asinh_Zero_ReturnsZero()
    {
        var func = AsinhFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
        });

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void Asinh_NegativeValue_ReturnsNegativeResult()
    {
        var func = AsinhFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-1),
        });

        Assert.Equal(-System.Math.Asinh(1), result.NumericValue, 10);
    }

    [Fact]
    public void Asinh_InvalidArguments_ReturnsError()
    {
        var func = AsinhFunction.Instance;

        // Wrong number of arguments
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1),
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

    // ACOSH Tests
    [Fact]
    public void Acosh_ValidInput_ReturnsCorrectValue()
    {
        var func = AcoshFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(2),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(System.Math.Acosh(2), result.NumericValue, 10);
    }

    [Fact]
    public void Acosh_One_ReturnsZero()
    {
        var func = AcoshFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1),
        });

        Assert.Equal(0.0, result.NumericValue, 10);
    }

    [Fact]
    public void Acosh_LessThanOne_ReturnsError()
    {
        var func = AcoshFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0.5),
        });

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Acosh_Zero_ReturnsError()
    {
        var func = AcoshFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
        });

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Acosh_NegativeValue_ReturnsError()
    {
        var func = AcoshFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-1),
        });

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Acosh_InvalidArguments_ReturnsError()
    {
        var func = AcoshFunction.Instance;

        // Wrong number of arguments
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(2),
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

    // ATANH Tests
    [Fact]
    public void Atanh_ValidInput_ReturnsCorrectValue()
    {
        var func = AtanhFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0.5),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(System.Math.Atanh(0.5), result.NumericValue, 10);
    }

    [Fact]
    public void Atanh_Zero_ReturnsZero()
    {
        var func = AtanhFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
        });

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void Atanh_NegativeValue_ReturnsNegativeResult()
    {
        var func = AtanhFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-0.5),
        });

        Assert.Equal(-System.Math.Atanh(0.5), result.NumericValue, 10);
    }

    [Fact]
    public void Atanh_OutOfRange_ReturnsError()
    {
        var func = AtanhFunction.Instance;

        // Greater than or equal to 1
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1),
        });

        Assert.True(result1.IsError);
        Assert.Equal("#NUM!", result1.ErrorValue);

        // Less than or equal to -1
        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-1),
        });

        Assert.True(result2.IsError);
        Assert.Equal("#NUM!", result2.ErrorValue);

        // Greater than 1
        var result3 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1.5),
        });

        Assert.True(result3.IsError);
        Assert.Equal("#NUM!", result3.ErrorValue);

        // Less than -1
        var result4 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-1.5),
        });

        Assert.True(result4.IsError);
        Assert.Equal("#NUM!", result4.ErrorValue);
    }

    [Fact]
    public void Atanh_InvalidArguments_ReturnsError()
    {
        var func = AtanhFunction.Instance;

        // Wrong number of arguments
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0.5),
            CellValue.FromNumber(0.5),
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
