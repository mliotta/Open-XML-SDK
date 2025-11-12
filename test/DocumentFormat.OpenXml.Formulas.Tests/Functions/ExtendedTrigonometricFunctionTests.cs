// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;
using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for extended trigonometric functions (SEC, CSC, COT, SECH, CSCH, COTH, ACOT, ACOTH).
/// </summary>
public class ExtendedTrigonometricFunctionTests
{
    // SEC Tests
    [Fact]
    public void Sec_ValidInput_ReturnsCorrectValue()
    {
        var func = SecFunction.Instance;

        // SEC(PI/3) = 1/COS(PI/3) = 1/0.5 = 2
        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(System.Math.PI / 3),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(2.0, result.NumericValue, 10);
    }

    [Fact]
    public void Sec_Zero_ReturnsOne()
    {
        var func = SecFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
        });

        Assert.Equal(1.0, result.NumericValue, 10);
    }

    [Fact]
    public void Sec_PiOver2_ReturnsDivByZeroError()
    {
        var func = SecFunction.Instance;

        // SEC(PI/2) is undefined (COS(PI/2) = 0)
        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(System.Math.PI / 2),
        });

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void Sec_InvalidArguments_ReturnsError()
    {
        var func = SecFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromString("text"),
        });

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    // SECH Tests
    [Fact]
    public void Sech_ValidInput_ReturnsCorrectValue()
    {
        var func = SechFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.0 / System.Math.Cosh(1), result.NumericValue, 10);
    }

    [Fact]
    public void Sech_Zero_ReturnsOne()
    {
        var func = SechFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
        });

        Assert.Equal(1.0, result.NumericValue);
    }

    // CSC Tests
    [Fact]
    public void Csc_ValidInput_ReturnsCorrectValue()
    {
        var func = CscFunction.Instance;

        // CSC(PI/6) = 1/SIN(PI/6) = 1/0.5 = 2
        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(System.Math.PI / 6),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(2.0, result.NumericValue, 10);
    }

    [Fact]
    public void Csc_Zero_ReturnsDivByZeroError()
    {
        var func = CscFunction.Instance;

        // CSC(0) is undefined (SIN(0) = 0)
        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
        });

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    // CSCH Tests
    [Fact]
    public void Csch_ValidInput_ReturnsCorrectValue()
    {
        var func = CschFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.0 / System.Math.Sinh(1), result.NumericValue, 10);
    }

    [Fact]
    public void Csch_Zero_ReturnsDivByZeroError()
    {
        var func = CschFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
        });

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    // COT Tests
    [Fact]
    public void Cot_ValidInput_ReturnsCorrectValue()
    {
        var func = CotFunction.Instance;

        // COT(PI/4) = 1/TAN(PI/4) = 1/1 = 1
        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(System.Math.PI / 4),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue, 10);
    }

    [Fact]
    public void Cot_Zero_ReturnsDivByZeroError()
    {
        var func = CotFunction.Instance;

        // COT(0) is undefined (TAN(0) = 0)
        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
        });

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    // COTH Tests
    [Fact]
    public void Coth_ValidInput_ReturnsCorrectValue()
    {
        var func = CothFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.0 / System.Math.Tanh(1), result.NumericValue, 10);
    }

    [Fact]
    public void Coth_Zero_ReturnsDivByZeroError()
    {
        var func = CothFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
        });

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    // ACOT Tests
    [Fact]
    public void Acot_ValidInput_ReturnsCorrectValue()
    {
        var func = AcotFunction.Instance;

        // ACOT(1) = PI/4
        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(System.Math.PI / 4, result.NumericValue, 10);
    }

    [Fact]
    public void Acot_Zero_ReturnsPiOver2()
    {
        var func = AcotFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
        });

        Assert.Equal(System.Math.PI / 2, result.NumericValue, 10);
    }

    [Fact]
    public void Acot_NegativeValue_ReturnsCorrectValue()
    {
        var func = AcotFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-1),
        });

        // ACOT(-1) = 3*PI/4
        Assert.Equal(3 * System.Math.PI / 4, result.NumericValue, 10);
    }

    // ACOTH Tests
    [Fact]
    public void Acoth_ValidInput_ReturnsCorrectValue()
    {
        var func = AcothFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(2),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        // ACOTH(2) = 0.5 * ln(3)
        var expected = 0.5 * System.Math.Log(3.0);
        Assert.Equal(expected, result.NumericValue, 10);
    }

    [Fact]
    public void Acoth_OutOfRange_ReturnsError()
    {
        var func = AcothFunction.Instance;

        // |x| must be > 1
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0.5),
        });

        Assert.True(result1.IsError);
        Assert.Equal("#NUM!", result1.ErrorValue);

        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1),
        });

        Assert.True(result2.IsError);
        Assert.Equal("#NUM!", result2.ErrorValue);

        var result3 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-0.5),
        });

        Assert.True(result3.IsError);
        Assert.Equal("#NUM!", result3.ErrorValue);
    }

    [Fact]
    public void Acoth_InvalidArguments_ReturnsError()
    {
        var func = AcothFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromString("text"),
        });

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }
}
