// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;
using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for regex functions (Excel 2025).
/// </summary>
public class RegexFunctionTests
{
    #region REGEXTEST Function Tests

    [Fact]
    public void RegexTest_BasicMatch_ReturnsTrue()
    {
        var func = RegexTestFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("abc123"),
            CellValue.FromString(@"\d+"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void RegexTest_NoMatch_ReturnsFalse()
    {
        var func = RegexTestFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("abcdef"),
            CellValue.FromString(@"\d+"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void RegexTest_CaseSensitiveDefault_ReturnsFalse()
    {
        var func = RegexTestFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("Hello"),
            CellValue.FromString("hello"),
        };

        var result = func.Execute(null!, args);

        Assert.False(result.BoolValue);
    }

    [Fact]
    public void RegexTest_CaseInsensitiveMode_ReturnsTrue()
    {
        var func = RegexTestFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("Hello"),
            CellValue.FromString("hello"),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.BoolValue);
    }

    [Fact]
    public void RegexTest_MultilineMode_MatchesAcrossLines()
    {
        var func = RegexTestFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("line1\nline2"),
            CellValue.FromString("^line2"),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.BoolValue);
    }

    [Fact]
    public void RegexTest_SinglelineMode_DotMatchesNewline()
    {
        var func = RegexTestFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("line1\nline2"),
            CellValue.FromString("line1.line2"),
            CellValue.FromNumber(4),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.BoolValue);
    }

    [Fact]
    public void RegexTest_CombinedModes_Works()
    {
        var func = RegexTestFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("HELLO\nWORLD"),
            CellValue.FromString("hello.world"),
            CellValue.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.BoolValue);
    }

    [Fact]
    public void RegexTest_EmailPattern_ReturnsTrue()
    {
        var func = RegexTestFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("user@example.com"),
            CellValue.FromString(@"^[\w\.-]+@[\w\.-]+\.\w+$"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.BoolValue);
    }

    [Fact]
    public void RegexTest_InvalidPattern_ReturnsError()
    {
        var func = RegexTestFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("test"),
            CellValue.FromString("[invalid"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void RegexTest_NegativeMode_ReturnsError()
    {
        var func = RegexTestFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("test"),
            CellValue.FromString("test"),
            CellValue.FromNumber(-1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void RegexTest_ErrorValue_PropagatesError()
    {
        var func = RegexTestFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#DIV/0!"),
            CellValue.FromString("test"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void RegexTest_WrongArgumentCount_ReturnsError()
    {
        var func = RegexTestFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("test"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region REGEXEXTRACT Function Tests

    [Fact]
    public void RegexExtract_BasicExtraction_ReturnsMatch()
    {
        var func = RegexExtractFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("abc123def"),
            CellValue.FromString(@"\d+"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("123", result.StringValue);
    }

    [Fact]
    public void RegexExtract_NoMatch_ReturnsNA()
    {
        var func = RegexExtractFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("abcdef"),
            CellValue.FromString(@"\d+"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    [Fact]
    public void RegexExtract_CaptureGroup_ReturnsGroup()
    {
        var func = RegexExtractFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("user@example.com"),
            CellValue.FromString(@"^([\w\.-]+)@([\w\.-]+)\.(\w+)$"),
            CellValue.FromNumber(0),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("user", result.StringValue);
    }

    [Fact]
    public void RegexExtract_CaptureGroup2_ReturnsGroup()
    {
        var func = RegexExtractFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("user@example.com"),
            CellValue.FromString(@"^([\w\.-]+)@([\w\.-]+)\.(\w+)$"),
            CellValue.FromNumber(0),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("example", result.StringValue);
    }

    [Fact]
    public void RegexExtract_Group0_ReturnsFullMatch()
    {
        var func = RegexExtractFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("Price: $99.99"),
            CellValue.FromString(@"\$(\d+\.\d+)"),
            CellValue.FromNumber(0),
            CellValue.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("$99.99", result.StringValue);
    }

    [Fact]
    public void RegexExtract_CaseInsensitive_ReturnsMatch()
    {
        var func = RegexExtractFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("HELLO123"),
            CellValue.FromString("hello"),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("HELLO", result.StringValue);
    }

    [Fact]
    public void RegexExtract_InvalidGroupNumber_ReturnsError()
    {
        var func = RegexExtractFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("test123"),
            CellValue.FromString(@"\d+"),
            CellValue.FromNumber(0),
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void RegexExtract_NegativeGroupNumber_ReturnsError()
    {
        var func = RegexExtractFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("test123"),
            CellValue.FromString(@"\d+"),
            CellValue.FromNumber(0),
            CellValue.FromNumber(-1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void RegexExtract_InvalidPattern_ReturnsError()
    {
        var func = RegexExtractFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("test"),
            CellValue.FromString("[invalid"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void RegexExtract_ErrorValue_PropagatesError()
    {
        var func = RegexExtractFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("test"),
            CellValue.Error("#REF!"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    [Fact]
    public void RegexExtract_WrongArgumentCount_ReturnsError()
    {
        var func = RegexExtractFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("test"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void RegexExtract_GroupOutOfRange_ReturnsError()
    {
        var func = RegexExtractFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("test123"),
            CellValue.FromString(@"(\d+)"),
            CellValue.FromNumber(0),
            CellValue.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region REGEXREPLACE Function Tests

    [Fact]
    public void RegexReplace_BasicReplacement_ReturnsExpected()
    {
        var func = RegexReplaceFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("abc123def456"),
            CellValue.FromString(@"\d+"),
            CellValue.FromString("X"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("abcXdefX", result.StringValue);
    }

    [Fact]
    public void RegexReplace_ReplaceAll_ReturnsExpected()
    {
        var func = RegexReplaceFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("Hello World Hello"),
            CellValue.FromString("Hello"),
            CellValue.FromString("Hi"),
            CellValue.FromNumber(0),
            CellValue.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("Hi World Hi", result.StringValue);
    }

    [Fact]
    public void RegexReplace_ReplaceFirstOccurrence_ReturnsExpected()
    {
        var func = RegexReplaceFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("abc123def456"),
            CellValue.FromString(@"\d+"),
            CellValue.FromString("X"),
            CellValue.FromNumber(0),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("abcXdef456", result.StringValue);
    }

    [Fact]
    public void RegexReplace_ReplaceSecondOccurrence_ReturnsExpected()
    {
        var func = RegexReplaceFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("abc123def456ghi789"),
            CellValue.FromString(@"\d+"),
            CellValue.FromString("X"),
            CellValue.FromNumber(0),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("abc123defXghi789", result.StringValue);
    }

    [Fact]
    public void RegexReplace_NoMatch_ReturnsUnchanged()
    {
        var func = RegexReplaceFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("abcdef"),
            CellValue.FromString(@"\d+"),
            CellValue.FromString("X"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("abcdef", result.StringValue);
    }

    [Fact]
    public void RegexReplace_CaseInsensitive_ReturnsExpected()
    {
        var func = RegexReplaceFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("Hello HELLO hello"),
            CellValue.FromString("hello"),
            CellValue.FromString("Hi"),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("Hi Hi Hi", result.StringValue);
    }

    [Fact]
    public void RegexReplace_EmptyReplacement_RemovesMatches()
    {
        var func = RegexReplaceFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("abc123def456"),
            CellValue.FromString(@"\d+"),
            CellValue.FromString(string.Empty),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("abcdef", result.StringValue);
    }

    [Fact]
    public void RegexReplace_PhoneNumberFormat_ReturnsExpected()
    {
        var func = RegexReplaceFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("555-123-4567"),
            CellValue.FromString(@"(\d{3})-(\d{3})-(\d{4})"),
            CellValue.FromString("($1) $2-$3"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("(555) 123-4567", result.StringValue);
    }

    [Fact]
    public void RegexReplace_OccurrenceBeyondMatches_ReturnsUnchanged()
    {
        var func = RegexReplaceFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("abc123def"),
            CellValue.FromString(@"\d+"),
            CellValue.FromString("X"),
            CellValue.FromNumber(0),
            CellValue.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("abc123def", result.StringValue);
    }

    [Fact]
    public void RegexReplace_InvalidPattern_ReturnsError()
    {
        var func = RegexReplaceFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("test"),
            CellValue.FromString("[invalid"),
            CellValue.FromString("X"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void RegexReplace_NegativeOccurrence_ReturnsError()
    {
        var func = RegexReplaceFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("test"),
            CellValue.FromString("test"),
            CellValue.FromString("X"),
            CellValue.FromNumber(0),
            CellValue.FromNumber(-1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void RegexReplace_ErrorValue_PropagatesError()
    {
        var func = RegexReplaceFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#N/A"),
            CellValue.FromString("test"),
            CellValue.FromString("X"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    [Fact]
    public void RegexReplace_WrongArgumentCount_ReturnsError()
    {
        var func = RegexReplaceFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("test"),
            CellValue.FromString("test"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void RegexReplace_MultilineMode_ReplacesAcrossLines()
    {
        var func = RegexReplaceFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("line1\nline2\nline3"),
            CellValue.FromString("^line"),
            CellValue.FromString("LINE"),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("LINE1\nLINE2\nLINE3", result.StringValue);
    }

    #endregion
}
