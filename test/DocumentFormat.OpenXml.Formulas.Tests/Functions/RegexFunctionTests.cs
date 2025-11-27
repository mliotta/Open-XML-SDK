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
            FormulaResult.FromString("abc123"),
            FormulaResult.FromString(@"\d+"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void RegexTest_NoMatch_ReturnsFalse()
    {
        var func = RegexTestFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("abcdef"),
            FormulaResult.FromString(@"\d+"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void RegexTest_CaseSensitiveDefault_ReturnsFalse()
    {
        var func = RegexTestFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Hello"),
            FormulaResult.FromString("hello"),
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
            FormulaResult.FromString("Hello"),
            FormulaResult.FromString("hello"),
            FormulaResult.FromNumber(1),
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
            FormulaResult.FromString("line1\nline2"),
            FormulaResult.FromString("^line2"),
            FormulaResult.FromNumber(2),
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
            FormulaResult.FromString("line1\nline2"),
            FormulaResult.FromString("line1.line2"),
            FormulaResult.FromNumber(4),
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
            FormulaResult.FromString("HELLO\nWORLD"),
            FormulaResult.FromString("hello.world"),
            FormulaResult.FromNumber(5),
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
            FormulaResult.FromString("user@example.com"),
            FormulaResult.FromString(@"^[\w\.-]+@[\w\.-]+\.\w+$"),
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
            FormulaResult.FromString("test"),
            FormulaResult.FromString("[invalid"),
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
            FormulaResult.FromString("test"),
            FormulaResult.FromString("test"),
            FormulaResult.FromNumber(-1),
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
            FormulaResult.Error("#DIV/0!"),
            FormulaResult.FromString("test"),
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
            FormulaResult.FromString("test"),
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
            FormulaResult.FromString("abc123def"),
            FormulaResult.FromString(@"\d+"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("123", result.StringValue);
    }

    [Fact]
    public void RegexExtract_NoMatch_ReturnsNA()
    {
        var func = RegexExtractFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("abcdef"),
            FormulaResult.FromString(@"\d+"),
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
            FormulaResult.FromString("user@example.com"),
            FormulaResult.FromString(@"^([\w\.-]+)@([\w\.-]+)\.(\w+)$"),
            FormulaResult.FromNumber(0),
            FormulaResult.FromNumber(1),
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
            FormulaResult.FromString("user@example.com"),
            FormulaResult.FromString(@"^([\w\.-]+)@([\w\.-]+)\.(\w+)$"),
            FormulaResult.FromNumber(0),
            FormulaResult.FromNumber(2),
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
            FormulaResult.FromString("Price: $99.99"),
            FormulaResult.FromString(@"\$(\d+\.\d+)"),
            FormulaResult.FromNumber(0),
            FormulaResult.FromNumber(0),
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
            FormulaResult.FromString("HELLO123"),
            FormulaResult.FromString("hello"),
            FormulaResult.FromNumber(1),
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
            FormulaResult.FromString("test123"),
            FormulaResult.FromString(@"\d+"),
            FormulaResult.FromNumber(0),
            FormulaResult.FromNumber(10),
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
            FormulaResult.FromString("test123"),
            FormulaResult.FromString(@"\d+"),
            FormulaResult.FromNumber(0),
            FormulaResult.FromNumber(-1),
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
            FormulaResult.FromString("test"),
            FormulaResult.FromString("[invalid"),
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
            FormulaResult.FromString("test"),
            FormulaResult.Error("#REF!"),
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
            FormulaResult.FromString("test"),
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
            FormulaResult.FromString("test123"),
            FormulaResult.FromString(@"(\d+)"),
            FormulaResult.FromNumber(0),
            FormulaResult.FromNumber(5),
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
            FormulaResult.FromString("abc123def456"),
            FormulaResult.FromString(@"\d+"),
            FormulaResult.FromString("X"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("abcXdefX", result.StringValue);
    }

    [Fact]
    public void RegexReplace_ReplaceAll_ReturnsExpected()
    {
        var func = RegexReplaceFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Hello World Hello"),
            FormulaResult.FromString("Hello"),
            FormulaResult.FromString("Hi"),
            FormulaResult.FromNumber(0),
            FormulaResult.FromNumber(0),
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
            FormulaResult.FromString("abc123def456"),
            FormulaResult.FromString(@"\d+"),
            FormulaResult.FromString("X"),
            FormulaResult.FromNumber(0),
            FormulaResult.FromNumber(1),
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
            FormulaResult.FromString("abc123def456ghi789"),
            FormulaResult.FromString(@"\d+"),
            FormulaResult.FromString("X"),
            FormulaResult.FromNumber(0),
            FormulaResult.FromNumber(2),
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
            FormulaResult.FromString("abcdef"),
            FormulaResult.FromString(@"\d+"),
            FormulaResult.FromString("X"),
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
            FormulaResult.FromString("Hello HELLO hello"),
            FormulaResult.FromString("hello"),
            FormulaResult.FromString("Hi"),
            FormulaResult.FromNumber(1),
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
            FormulaResult.FromString("abc123def456"),
            FormulaResult.FromString(@"\d+"),
            FormulaResult.FromString(string.Empty),
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
            FormulaResult.FromString("555-123-4567"),
            FormulaResult.FromString(@"(\d{3})-(\d{3})-(\d{4})"),
            FormulaResult.FromString("($1) $2-$3"),
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
            FormulaResult.FromString("abc123def"),
            FormulaResult.FromString(@"\d+"),
            FormulaResult.FromString("X"),
            FormulaResult.FromNumber(0),
            FormulaResult.FromNumber(5),
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
            FormulaResult.FromString("test"),
            FormulaResult.FromString("[invalid"),
            FormulaResult.FromString("X"),
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
            FormulaResult.FromString("test"),
            FormulaResult.FromString("test"),
            FormulaResult.FromString("X"),
            FormulaResult.FromNumber(0),
            FormulaResult.FromNumber(-1),
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
            FormulaResult.Error("#N/A"),
            FormulaResult.FromString("test"),
            FormulaResult.FromString("X"),
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
            FormulaResult.FromString("test"),
            FormulaResult.FromString("test"),
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
            FormulaResult.FromString("line1\nline2\nline3"),
            FormulaResult.FromString("^line"),
            FormulaResult.FromString("LINE"),
            FormulaResult.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("LINE1\nLINE2\nLINE3", result.StringValue);
    }

    #endregion
}
