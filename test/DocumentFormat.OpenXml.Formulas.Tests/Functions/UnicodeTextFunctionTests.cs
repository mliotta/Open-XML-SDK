// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;
using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for Unicode and additional text manipulation functions.
/// </summary>
public class UnicodeTextFunctionTests
{
    #region TRIMALL Function Tests

    [Fact]
    public void TrimAll_RemovesAllSpaces_ReturnsExpected()
    {
        var func = TrimAllFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("  hello  world  "),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("helloworld", result.StringValue);
    }

    [Fact]
    public void TrimAll_MultipleSpaces_RemovesAll()
    {
        var func = TrimAllFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("  a  b  c  "),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("abc", result.StringValue);
    }

    [Fact]
    public void TrimAll_NoSpaces_ReturnsUnchanged()
    {
        var func = TrimAllFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("helloworld"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("helloworld", result.StringValue);
    }

    [Fact]
    public void TrimAll_OnlySpaces_ReturnsEmpty()
    {
        var func = TrimAllFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("     "),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(string.Empty, result.StringValue);
    }

    [Fact]
    public void TrimAll_EmptyString_ReturnsEmpty()
    {
        var func = TrimAllFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString(string.Empty),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(string.Empty, result.StringValue);
    }

    [Fact]
    public void TrimAll_WithTabs_KeepsTabs()
    {
        var func = TrimAllFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("hello\tworld"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("hello\tworld", result.StringValue);
    }

    [Fact]
    public void TrimAll_ErrorValue_PropagatesError()
    {
        var func = TrimAllFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#N/A"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    [Fact]
    public void TrimAll_WrongArgumentCount_ReturnsError()
    {
        var func = TrimAllFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("test"),
            FormulaResult.FromString("extra"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region UNICHAR Function Tests

    [Fact]
    public void Unichar_LetterA_ReturnsA()
    {
        var func = UnicharFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(65),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("A", result.StringValue);
    }

    [Fact]
    public void Unichar_Snowman_ReturnsSnowman()
    {
        var func = UnicharFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(9731),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("☃", result.StringValue);
    }

    [Fact]
    public void Unichar_EuroSign_ReturnsEuro()
    {
        var func = UnicharFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(8364),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("€", result.StringValue);
    }

    [Fact]
    public void Unichar_SupplementaryPlane_ReturnsCorrectCharacter()
    {
        var func = UnicharFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(128512), // 😀 (Grinning Face)
        };

        var result = func.Execute(null!, args);

        Assert.Equal("😀", result.StringValue);
    }

    [Fact]
    public void Unichar_ChineseCharacter_ReturnsCharacter()
    {
        var func = UnicharFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(20013), // 中
        };

        var result = func.Execute(null!, args);

        Assert.Equal("中", result.StringValue);
    }

    [Fact]
    public void Unichar_BelowRange_ReturnsError()
    {
        var func = UnicharFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Unichar_AboveRange_ReturnsError()
    {
        var func = UnicharFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(1114112),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Unichar_SurrogateRange_ReturnsError()
    {
        var func = UnicharFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(0xD800), // Start of surrogate range
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Unichar_NonNumericValue_ReturnsError()
    {
        var func = UnicharFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("not a number"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Unichar_ErrorValue_PropagatesError()
    {
        var func = UnicharFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#REF!"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    [Fact]
    public void Unichar_WrongArgumentCount_ReturnsError()
    {
        var func = UnicharFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(65),
            FormulaResult.FromNumber(66),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region UNICODE Function Tests

    [Fact]
    public void Unicode_LetterA_Returns65()
    {
        var func = UnicodeFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("A"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(65.0, result.NumericValue);
    }

    [Fact]
    public void Unicode_Snowman_Returns9731()
    {
        var func = UnicodeFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("☃"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(9731.0, result.NumericValue);
    }

    [Fact]
    public void Unicode_EuroSign_Returns8364()
    {
        var func = UnicodeFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("€"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(8364.0, result.NumericValue);
    }

    [Fact]
    public void Unicode_ChineseCharacter_ReturnsCodePoint()
    {
        var func = UnicodeFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("中"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(20013.0, result.NumericValue);
    }

    [Fact]
    public void Unicode_SupplementaryPlane_ReturnsCorrectCodePoint()
    {
        var func = UnicodeFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("😀"), // Grinning Face
        };

        var result = func.Execute(null!, args);

        Assert.Equal(128512.0, result.NumericValue);
    }

    [Fact]
    public void Unicode_MultipleCharacters_ReturnsFirstCharCodePoint()
    {
        var func = UnicodeFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("ABC"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(65.0, result.NumericValue);
    }

    [Fact]
    public void Unicode_MultipleCharactersWithEmoji_ReturnsFirstCodePoint()
    {
        var func = UnicodeFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("😀ABC"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(128512.0, result.NumericValue);
    }

    [Fact]
    public void Unicode_LowercaseA_Returns97()
    {
        var func = UnicodeFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("a"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(97.0, result.NumericValue);
    }

    [Fact]
    public void Unicode_Space_Returns32()
    {
        var func = UnicodeFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString(" "),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(32.0, result.NumericValue);
    }

    [Fact]
    public void Unicode_EmptyString_ReturnsError()
    {
        var func = UnicodeFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString(string.Empty),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Unicode_ErrorValue_PropagatesError()
    {
        var func = UnicodeFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#NULL!"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NULL!", result.ErrorValue);
    }

    [Fact]
    public void Unicode_WrongArgumentCount_ReturnsError()
    {
        var func = UnicodeFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("A"),
            FormulaResult.FromString("B"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region PHONETIC Function Tests

    [Fact]
    public void Phonetic_SimpleText_ReturnsText()
    {
        var func = PhoneticFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Tokyo"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("Tokyo", result.StringValue);
    }

    [Fact]
    public void Phonetic_JapaneseText_ReturnsText()
    {
        var func = PhoneticFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("東京"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("東京", result.StringValue);
    }

    [Fact]
    public void Phonetic_EmptyString_ReturnsEmpty()
    {
        var func = PhoneticFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString(string.Empty),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(string.Empty, result.StringValue);
    }

    [Fact]
    public void Phonetic_NumericValue_ReturnsAsText()
    {
        var func = PhoneticFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(12345),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("12345", result.StringValue);
    }

    [Fact]
    public void Phonetic_BooleanValue_ReturnsAsText()
    {
        var func = PhoneticFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("TRUE", result.StringValue);
    }

    [Fact]
    public void Phonetic_ErrorValue_PropagatesError()
    {
        var func = PhoneticFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#NAME?"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NAME?", result.ErrorValue);
    }

    [Fact]
    public void Phonetic_WrongArgumentCount_ReturnsError()
    {
        var func = PhoneticFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("test"),
            FormulaResult.FromString("extra"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Phonetic_NoArguments_ReturnsError()
    {
        var func = PhoneticFunction.Instance;
        var args = Array.Empty<FormulaResult>();

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region Round-trip Tests (UNICHAR/UNICODE)

    [Fact]
    public void UnicharUnicode_RoundTrip_BasicLatin()
    {
        // Test UNICHAR(65) = "A" and UNICODE("A") = 65
        var unichar = UnicharFunction.Instance;
        var unicode = UnicodeFunction.Instance;

        var charResult = unichar.Execute(null!, new[] { FormulaResult.FromNumber(65) });
        Assert.Equal("A", charResult.StringValue);

        var codeResult = unicode.Execute(null!, new[] { charResult });
        Assert.Equal(65.0, codeResult.NumericValue);
    }

    [Fact]
    public void UnicharUnicode_RoundTrip_Emoji()
    {
        // Test UNICHAR(128512) = "😀" and UNICODE("😀") = 128512
        var unichar = UnicharFunction.Instance;
        var unicode = UnicodeFunction.Instance;

        var charResult = unichar.Execute(null!, new[] { FormulaResult.FromNumber(128512) });
        Assert.Equal("😀", charResult.StringValue);

        var codeResult = unicode.Execute(null!, new[] { charResult });
        Assert.Equal(128512.0, codeResult.NumericValue);
    }

    [Fact]
    public void UnicharUnicode_RoundTrip_ChineseCharacter()
    {
        // Test UNICHAR(20013) = "中" and UNICODE("中") = 20013
        var unichar = UnicharFunction.Instance;
        var unicode = UnicodeFunction.Instance;

        var charResult = unichar.Execute(null!, new[] { FormulaResult.FromNumber(20013) });
        Assert.Equal("中", charResult.StringValue);

        var codeResult = unicode.Execute(null!, new[] { charResult });
        Assert.Equal(20013.0, codeResult.NumericValue);
    }

    #endregion
}
