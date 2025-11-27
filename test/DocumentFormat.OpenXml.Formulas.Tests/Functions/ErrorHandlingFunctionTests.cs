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
            FormulaResult.Error("#DIV/0!"),
            FormulaResult.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void IFError_ValidValue_ReturnsOriginal()
    {
        var func = IFErrorFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(42),
            FormulaResult.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(42.0, result.NumericValue);
    }

    [Fact]
    public void IFError_TextValue_ReturnsOriginal()
    {
        var func = IFErrorFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Hello"),
            FormulaResult.FromString("Error"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("Hello", result.StringValue);
    }

    [Fact]
    public void IFError_WrongArgumentCount_ReturnsError()
    {
        var func = IFErrorFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(42),
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
            FormulaResult.Error("#DIV/0!"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void IsError_ValidValue_ReturnsFalse()
    {
        var func = IsErrorFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(42),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void IsError_NAError_ReturnsTrue()
    {
        var func = IsErrorFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#N/A"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void IsNA_NAError_ReturnsTrue()
    {
        var func = IsNaFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#N/A"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void IsNA_OtherError_ReturnsFalse()
    {
        var func = IsNaFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#DIV/0!"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void IsNA_ValidValue_ReturnsFalse()
    {
        var func = IsNaFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(42),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void IsErr_DivError_ReturnsTrue()
    {
        var func = IsErrFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#DIV/0!"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void IsErr_NAError_ReturnsFalse()
    {
        var func = IsErrFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#N/A"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void IsErr_ValidValue_ReturnsFalse()
    {
        var func = IsErrFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(42),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void IsErr_ValueError_ReturnsTrue()
    {
        var func = IsErrFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#VALUE!"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void IsBlank_EmptyValue_ReturnsTrue()
    {
        var func = IsBlankFunction.Instance;
        var args = new[]
        {
            FormulaResult.Empty,
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void IsBlank_NumberValue_ReturnsFalse()
    {
        var func = IsBlankFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void IsBlank_EmptyString_ReturnsFalse()
    {
        var func = IsBlankFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString(""),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void IsBlank_TextValue_ReturnsFalse()
    {
        var func = IsBlankFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Hello"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void IsBlank_WrongArgumentCount_ReturnsError()
    {
        var func = IsBlankFunction.Instance;
        var args = new[]
        {
            FormulaResult.Empty,
            FormulaResult.Empty,
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
            FormulaResult.FromNumber(1),
            FormulaResult.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }
}
