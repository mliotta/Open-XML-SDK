// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;
using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for ASC and DBCS functions (Japanese text conversion).
/// </summary>
public class AscDbcsFunctionTests
{
    #region ASC Function Tests

    [Fact]
    public void Asc_FullWidthToHalfWidth_ReturnsHalfWidth()
    {
        var func = AscFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("ＡＢＣ"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("ABC", result.StringValue);
    }

    [Fact]
    public void Asc_FullWidthSpace_ReturnsHalfWidthSpace()
    {
        var func = AscFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("　"), // Full-width space U+3000
        };

        var result = func.Execute(null!, args);

        Assert.Equal(" ", result.StringValue); // Half-width space U+0020
    }

    [Fact]
    public void Asc_MixedWidthText_ConvertsOnlyFullWidth()
    {
        var func = AscFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("ＡBＣ"), // Full A, half B, full C
        };

        var result = func.Execute(null!, args);

        Assert.Equal("ABC", result.StringValue);
    }

    [Fact]
    public void Asc_FullWidthNumbers_ReturnsHalfWidthNumbers()
    {
        var func = AscFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("１２３"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("123", result.StringValue);
    }

    [Fact]
    public void Asc_EmptyString_ReturnsEmptyString()
    {
        var func = AscFunction.Instance;
        var args = new[]
        {
            CellValue.FromString(string.Empty),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(string.Empty, result.StringValue);
    }

    [Fact]
    public void Asc_WrongNumberOfArgs_ReturnsError()
    {
        var func = AscFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("ABC"),
            CellValue.FromString("DEF"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Asc_ErrorPropagation_ReturnsError()
    {
        var func = AscFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#DIV/0!"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void Asc_FullWidthSymbols_ConvertsToHalfWidth()
    {
        var func = AscFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("！＠＃"), // Full-width !, @, #
        };

        var result = func.Execute(null!, args);

        Assert.Equal("!@#", result.StringValue);
    }

    #endregion

    #region DBCS Function Tests

    [Fact]
    public void Dbcs_HalfWidthToFullWidth_ReturnsFullWidth()
    {
        var func = DbcsFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("ABC"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("ＡＢＣ", result.StringValue);
    }

    [Fact]
    public void Dbcs_HalfWidthSpace_ReturnsFullWidthSpace()
    {
        var func = DbcsFunction.Instance;
        var args = new[]
        {
            CellValue.FromString(" "), // Half-width space U+0020
        };

        var result = func.Execute(null!, args);

        Assert.Equal("　", result.StringValue); // Full-width space U+3000
    }

    [Fact]
    public void Dbcs_MixedWidthText_ConvertsOnlyHalfWidth()
    {
        var func = DbcsFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("AＢC"), // Half A, full B, half C
        };

        var result = func.Execute(null!, args);

        Assert.Equal("ＡＢＣ", result.StringValue);
    }

    [Fact]
    public void Dbcs_HalfWidthNumbers_ReturnsFullWidthNumbers()
    {
        var func = DbcsFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("123"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("１２３", result.StringValue);
    }

    [Fact]
    public void Dbcs_EmptyString_ReturnsEmptyString()
    {
        var func = DbcsFunction.Instance;
        var args = new[]
        {
            CellValue.FromString(string.Empty),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(string.Empty, result.StringValue);
    }

    [Fact]
    public void Dbcs_WrongNumberOfArgs_ReturnsError()
    {
        var func = DbcsFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("ABC"),
            CellValue.FromString("DEF"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Dbcs_ErrorPropagation_ReturnsError()
    {
        var func = DbcsFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#DIV/0!"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void Dbcs_HalfWidthSymbols_ConvertsToFullWidth()
    {
        var func = DbcsFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("!@#"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("！＠＃", result.StringValue);
    }

    #endregion

    #region Round-trip Tests

    [Fact]
    public void AscDbcs_RoundTrip_ReturnsOriginal()
    {
        var ascFunc = AscFunction.Instance;
        var dbcsFunc = DbcsFunction.Instance;

        string original = "ＡＢＣ１２３";

        // Full-width -> half-width
        var halfWidth = ascFunc.Execute(null!, new[] { CellValue.FromString(original) });
        Assert.Equal("ABC123", halfWidth.StringValue);

        // Half-width -> full-width (should restore original)
        var fullWidth = dbcsFunc.Execute(null!, new[] { halfWidth });
        Assert.Equal(original, fullWidth.StringValue);
    }

    #endregion
}
