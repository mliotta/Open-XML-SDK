// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for error handling and information functions.
/// </summary>
public class ErrorHandlingFunctionTests
{
    [Fact]
    public void IFError_ErrorValue_ReturnsAlternative()
    {
        var func = IFErrorFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void IFError_ValidValue_ReturnsOriginal()
    {
        var func = IFErrorFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(42),
            CellValue.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(42.0, result.NumericValue);
    }

    [Fact]
    public void IFError_TextValue_ReturnsOriginal()
    {
        var func = IFErrorFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("Hello"),
            CellValue.FromString("Error"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("Hello", result.StringValue);
    }

    [Fact]
    public void IFError_WrongArgumentCount_ReturnsError()
    {
        var func = IFErrorFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(42),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void IsError_ErrorValue_ReturnsTrue()
    {
        var func = IsErrorFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#DIV/0!"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void IsError_ValidValue_ReturnsFalse()
    {
        var func = IsErrorFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(42),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void IsError_NAError_ReturnsTrue()
    {
        var func = IsErrorFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#N/A"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void IsNA_NAError_ReturnsTrue()
    {
        var func = IsNaFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#N/A"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void IsNA_OtherError_ReturnsFalse()
    {
        var func = IsNaFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#DIV/0!"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void IsNA_ValidValue_ReturnsFalse()
    {
        var func = IsNaFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(42),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void IsErr_DivError_ReturnsTrue()
    {
        var func = IsErrFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#DIV/0!"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void IsErr_NAError_ReturnsFalse()
    {
        var func = IsErrFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#N/A"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void IsErr_ValidValue_ReturnsFalse()
    {
        var func = IsErrFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(42),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void IsErr_ValueError_ReturnsTrue()
    {
        var func = IsErrFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#VALUE!"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void IsBlank_EmptyValue_ReturnsTrue()
    {
        var func = IsBlankFunction.Instance;
        var args = new[]
        {
            CellValue.Empty,
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void IsBlank_NumberValue_ReturnsFalse()
    {
        var func = IsBlankFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void IsBlank_EmptyString_ReturnsFalse()
    {
        var func = IsBlankFunction.Instance;
        var args = new[]
        {
            CellValue.FromString(""),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void IsBlank_TextValue_ReturnsFalse()
    {
        var func = IsBlankFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("Hello"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void IsBlank_WrongArgumentCount_ReturnsError()
    {
        var func = IsBlankFunction.Instance;
        var args = new[]
        {
            CellValue.Empty,
            CellValue.Empty,
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void IsError_WrongArgumentCount_ReturnsError()
    {
        var func = IsErrorFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }
}
