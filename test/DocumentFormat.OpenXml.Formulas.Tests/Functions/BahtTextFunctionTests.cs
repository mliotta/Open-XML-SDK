// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;
using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for BAHTTEXT function.
/// Note: This is a simplified implementation that returns a placeholder format.
/// </summary>
public class BahtTextFunctionTests
{
    [Fact]
    public void BahtText_PositiveNumber_ReturnsBahtSuffix()
    {
        var func = BahtTextFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1234.56),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Contains("1234.56", result.StringValue);
        Assert.Contains("บาท", result.StringValue); // Thai Baht symbol
    }

    [Fact]
    public void BahtText_Zero_ReturnsZeroBaht()
    {
        var func = BahtTextFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.Contains("0.00", result.StringValue);
        Assert.Contains("บาท", result.StringValue);
    }

    [Fact]
    public void BahtText_NegativeNumber_ReturnsNegativeBaht()
    {
        var func = BahtTextFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-500.25),
        };

        var result = func.Execute(null!, args);

        Assert.Contains("-500.25", result.StringValue);
        Assert.Contains("บาท", result.StringValue);
    }

    [Fact]
    public void BahtText_LargeNumber_ReturnsBahtSuffix()
    {
        var func = BahtTextFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1000000),
        };

        var result = func.Execute(null!, args);

        Assert.Contains("1000000", result.StringValue);
        Assert.Contains("บาท", result.StringValue);
    }

    [Fact]
    public void BahtText_NonNumericArgument_ReturnsError()
    {
        var func = BahtTextFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("abc"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void BahtText_WrongNumberOfArgs_ReturnsError()
    {
        var func = BahtTextFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(100),
            CellValue.FromNumber(200),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void BahtText_ErrorPropagation_ReturnsError()
    {
        var func = BahtTextFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#DIV/0!"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void BahtText_DecimalFormatting_UsesCorrectFormat()
    {
        var func = BahtTextFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(123.4),
        };

        var result = func.Execute(null!, args);

        // Should format with 2 decimal places
        Assert.Contains("123.40", result.StringValue);
        Assert.Contains("บาท", result.StringValue);
    }
}
