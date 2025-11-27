// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;
using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for text manipulation functions.
/// </summary>
public class TextFunctionTests
{
    #region REPLACE Function Tests

    [Fact]
    public void Replace_BasicReplacement_ReturnsExpected()
    {
        var func = ReplaceFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Excel"),
            FormulaResult.FromNumber(2),
            FormulaResult.FromNumber(2),
            FormulaResult.FromString("pert"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("Expert", result.StringValue);
    }

    [Fact]
    public void Replace_ReplaceAtStart_ReturnsExpected()
    {
        var func = ReplaceFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("abcdef"),
            FormulaResult.FromNumber(1),
            FormulaResult.FromNumber(3),
            FormulaResult.FromString("XYZ"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("XYZdef", result.StringValue);
    }

    [Fact]
    public void Replace_ReplaceAtEnd_ReturnsExpected()
    {
        var func = ReplaceFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("abcdef"),
            FormulaResult.FromNumber(4),
            FormulaResult.FromNumber(3),
            FormulaResult.FromString("XYZ"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("abcXYZ", result.StringValue);
    }

    [Fact]
    public void Replace_ZeroCharsReplaced_InsertsText()
    {
        var func = ReplaceFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("abc"),
            FormulaResult.FromNumber(2),
            FormulaResult.FromNumber(0),
            FormulaResult.FromString("XYZ"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("aXYZbc", result.StringValue);
    }

    [Fact]
    public void Replace_StartBeyondLength_AppendsText()
    {
        var func = ReplaceFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("abc"),
            FormulaResult.FromNumber(10),
            FormulaResult.FromNumber(1),
            FormulaResult.FromString("XYZ"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("abcXYZ", result.StringValue);
    }

    [Fact]
    public void Replace_InvalidStartPosition_ReturnsError()
    {
        var func = ReplaceFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("test"),
            FormulaResult.FromNumber(0),
            FormulaResult.FromNumber(1),
            FormulaResult.FromString("X"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Replace_NegativeNumChars_ReturnsError()
    {
        var func = ReplaceFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("test"),
            FormulaResult.FromNumber(1),
            FormulaResult.FromNumber(-1),
            FormulaResult.FromString("X"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Replace_ErrorValue_PropagatesError()
    {
        var func = ReplaceFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#DIV/0!"),
            FormulaResult.FromNumber(1),
            FormulaResult.FromNumber(1),
            FormulaResult.FromString("X"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void Replace_WrongArgumentCount_ReturnsError()
    {
        var func = ReplaceFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("test"),
            FormulaResult.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region REPT Function Tests

    [Fact]
    public void Rept_BasicRepetition_ReturnsExpected()
    {
        var func = ReptFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("*"),
            FormulaResult.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("*****", result.StringValue);
    }

    [Fact]
    public void Rept_MultiCharacterString_ReturnsExpected()
    {
        var func = ReptFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("ab"),
            FormulaResult.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("ababab", result.StringValue);
    }

    [Fact]
    public void Rept_ZeroTimes_ReturnsEmpty()
    {
        var func = ReptFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("test"),
            FormulaResult.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(string.Empty, result.StringValue);
    }

    [Fact]
    public void Rept_OneTimes_ReturnsSameString()
    {
        var func = ReptFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("test"),
            FormulaResult.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("test", result.StringValue);
    }

    [Fact]
    public void Rept_NegativeTimes_ReturnsError()
    {
        var func = ReptFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("test"),
            FormulaResult.FromNumber(-1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Rept_ErrorValue_PropagatesError()
    {
        var func = ReptFunction.Instance;
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
    public void Rept_WrongArgumentCount_ReturnsError()
    {
        var func = ReptFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("test"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region EXACT Function Tests

    [Fact]
    public void Exact_IdenticalStrings_ReturnsTrue()
    {
        var func = ExactFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Excel"),
            FormulaResult.FromString("Excel"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void Exact_DifferentCase_ReturnsFalse()
    {
        var func = ExactFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Excel"),
            FormulaResult.FromString("excel"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void Exact_DifferentStrings_ReturnsFalse()
    {
        var func = ExactFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Excel"),
            FormulaResult.FromString("Word"),
        };

        var result = func.Execute(null!, args);

        Assert.False(result.BoolValue);
    }

    [Fact]
    public void Exact_EmptyStrings_ReturnsTrue()
    {
        var func = ExactFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString(string.Empty),
            FormulaResult.FromString(string.Empty),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.BoolValue);
    }

    [Fact]
    public void Exact_WithWhitespace_IsCaseSensitive()
    {
        var func = ExactFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("test "),
            FormulaResult.FromString("test"),
        };

        var result = func.Execute(null!, args);

        Assert.False(result.BoolValue);
    }

    [Fact]
    public void Exact_ErrorValue_PropagatesError()
    {
        var func = ExactFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#N/A"),
            FormulaResult.FromString("test"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    [Fact]
    public void Exact_WrongArgumentCount_ReturnsError()
    {
        var func = ExactFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("test"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region CHAR Function Tests

    [Fact]
    public void Char_LetterA_ReturnsA()
    {
        var func = CharFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(65),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("A", result.StringValue);
    }

    [Fact]
    public void Char_LetterZ_ReturnsZ()
    {
        var func = CharFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(90),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("Z", result.StringValue);
    }

    [Fact]
    public void Char_Space_ReturnsSpace()
    {
        var func = CharFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(32),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(" ", result.StringValue);
    }

    [Fact]
    public void Char_Digit_ReturnsDigit()
    {
        var func = CharFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(48),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("0", result.StringValue);
    }

    [Fact]
    public void Char_BelowRange_ReturnsError()
    {
        var func = CharFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Char_AboveRange_ReturnsError()
    {
        var func = CharFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(256),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Char_ErrorValue_PropagatesError()
    {
        var func = CharFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#NAME?"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NAME?", result.ErrorValue);
    }

    [Fact]
    public void Char_WrongArgumentCount_ReturnsError()
    {
        var func = CharFunction.Instance;
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

    #region CODE Function Tests

    [Fact]
    public void Code_LetterA_Returns65()
    {
        var func = CodeFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("A"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(65.0, result.NumericValue);
    }

    [Fact]
    public void Code_LetterZ_Returns90()
    {
        var func = CodeFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Z"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(90.0, result.NumericValue);
    }

    [Fact]
    public void Code_LowercaseA_Returns97()
    {
        var func = CodeFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("a"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(97.0, result.NumericValue);
    }

    [Fact]
    public void Code_MultipleCharacters_ReturnsFirstCharCode()
    {
        var func = CodeFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("ABC"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(65.0, result.NumericValue);
    }

    [Fact]
    public void Code_Space_Returns32()
    {
        var func = CodeFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString(" "),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(32.0, result.NumericValue);
    }

    [Fact]
    public void Code_EmptyString_ReturnsError()
    {
        var func = CodeFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString(string.Empty),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Code_ErrorValue_PropagatesError()
    {
        var func = CodeFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#NULL!"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NULL!", result.ErrorValue);
    }

    [Fact]
    public void Code_WrongArgumentCount_ReturnsError()
    {
        var func = CodeFunction.Instance;
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

    #region CLEAN Function Tests

    [Fact]
    public void Clean_RemovesNonPrintableCharacters_ReturnsExpected()
    {
        var func = CleanFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("text\n\r"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("text", result.StringValue);
    }

    [Fact]
    public void Clean_RemovesTab_ReturnsExpected()
    {
        var func = CleanFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("hello\tworld"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("helloworld", result.StringValue);
    }

    [Fact]
    public void Clean_RemovesMultipleNonPrintable_ReturnsExpected()
    {
        var func = CleanFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("a\x01b\x02c\x03d"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("abcd", result.StringValue);
    }

    [Fact]
    public void Clean_NormalText_ReturnsUnchanged()
    {
        var func = CleanFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Hello World"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("Hello World", result.StringValue);
    }

    [Fact]
    public void Clean_EmptyString_ReturnsEmpty()
    {
        var func = CleanFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString(string.Empty),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(string.Empty, result.StringValue);
    }

    [Fact]
    public void Clean_OnlyNonPrintable_ReturnsEmpty()
    {
        var func = CleanFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("\x01\x02\x03"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(string.Empty, result.StringValue);
    }

    [Fact]
    public void Clean_ErrorValue_PropagatesError()
    {
        var func = CleanFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#VALUE!"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Clean_WrongArgumentCount_ReturnsError()
    {
        var func = CleanFunction.Instance;
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

    #region T Function Tests

    [Fact]
    public void T_TextValue_ReturnsText()
    {
        var func = TFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("hello"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("hello", result.StringValue);
    }

    [Fact]
    public void T_NumberValue_ReturnsEmpty()
    {
        var func = TFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(123),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal(string.Empty, result.StringValue);
    }

    [Fact]
    public void T_BooleanValue_ReturnsEmpty()
    {
        var func = TFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(string.Empty, result.StringValue);
    }

    [Fact]
    public void T_EmptyString_ReturnsEmpty()
    {
        var func = TFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString(string.Empty),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(string.Empty, result.StringValue);
    }

    [Fact]
    public void T_ErrorValue_PropagatesError()
    {
        var func = TFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#DIV/0!"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void T_WrongArgumentCount_ReturnsError()
    {
        var func = TFunction.Instance;
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

    #region CONCAT Function Tests

    [Fact]
    public void Concat_TwoStrings_ReturnsExpected()
    {
        var func = ConcatFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Hello"),
            FormulaResult.FromString(" World"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("Hello World", result.StringValue);
    }

    [Fact]
    public void Concat_ThreeStrings_ReturnsExpected()
    {
        var func = ConcatFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Hello"),
            FormulaResult.FromString(" "),
            FormulaResult.FromString("World"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("Hello World", result.StringValue);
    }

    [Fact]
    public void Concat_NumbersAsText_ReturnsExpected()
    {
        var func = ConcatFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(2024),
            FormulaResult.FromString("-"),
            FormulaResult.FromNumber(11),
            FormulaResult.FromString("-"),
            FormulaResult.FromNumber(11),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("2024-11-11", result.StringValue);
    }

    [Fact]
    public void Concat_EmptyStrings_ReturnsEmpty()
    {
        var func = ConcatFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString(string.Empty),
            FormulaResult.FromString(string.Empty),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(string.Empty, result.StringValue);
    }

    [Fact]
    public void Concat_SingleString_ReturnsSame()
    {
        var func = ConcatFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Test"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("Test", result.StringValue);
    }

    [Fact]
    public void Concat_ErrorValue_PropagatesError()
    {
        var func = ConcatFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Hello"),
            FormulaResult.Error("#DIV/0!"),
            FormulaResult.FromString("World"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    #endregion

    #region TEXTJOIN Function Tests

    [Fact]
    public void TextJoin_BasicJoin_ReturnsExpected()
    {
        var func = TextJoinFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString(", "),
            FormulaResult.FromBool(true),
            FormulaResult.FromString("A"),
            FormulaResult.FromString("B"),
            FormulaResult.FromString("C"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("A, B, C", result.StringValue);
    }

    [Fact]
    public void TextJoin_IgnoreEmpty_SkipsEmptyStrings()
    {
        var func = TextJoinFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString(", "),
            FormulaResult.FromBool(true),
            FormulaResult.FromString("A"),
            FormulaResult.FromString(string.Empty),
            FormulaResult.FromString("B"),
            FormulaResult.FromString("C"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("A, B, C", result.StringValue);
    }

    [Fact]
    public void TextJoin_KeepEmpty_IncludesEmptyStrings()
    {
        var func = TextJoinFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString(", "),
            FormulaResult.FromBool(false),
            FormulaResult.FromString("A"),
            FormulaResult.FromString(string.Empty),
            FormulaResult.FromString("B"),
            FormulaResult.FromString("C"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("A, , B, C", result.StringValue);
    }

    [Fact]
    public void TextJoin_DateFormat_ReturnsExpected()
    {
        var func = TextJoinFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("-"),
            FormulaResult.FromBool(false),
            FormulaResult.FromString("2024"),
            FormulaResult.FromString("01"),
            FormulaResult.FromString("15"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("2024-01-15", result.StringValue);
    }

    [Fact]
    public void TextJoin_EmptyDelimiter_ConcatenatesDirectly()
    {
        var func = TextJoinFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString(string.Empty),
            FormulaResult.FromBool(true),
            FormulaResult.FromString("A"),
            FormulaResult.FromString("B"),
            FormulaResult.FromString("C"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("ABC", result.StringValue);
    }

    [Fact]
    public void TextJoin_NumericIgnoreEmpty_AcceptsNumber()
    {
        var func = TextJoinFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString(", "),
            FormulaResult.FromNumber(1),
            FormulaResult.FromString("A"),
            FormulaResult.FromString("B"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("A, B", result.StringValue);
    }

    [Fact]
    public void TextJoin_ZeroIgnoreEmpty_KeepsEmpty()
    {
        var func = TextJoinFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString(", "),
            FormulaResult.FromNumber(0),
            FormulaResult.FromString("A"),
            FormulaResult.FromString(string.Empty),
            FormulaResult.FromString("B"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("A, , B", result.StringValue);
    }

    [Fact]
    public void TextJoin_TextBooleanTrue_AcceptsText()
    {
        var func = TextJoinFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString(", "),
            FormulaResult.FromString("TRUE"),
            FormulaResult.FromString("A"),
            FormulaResult.FromString("B"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("A, B", result.StringValue);
    }

    [Fact]
    public void TextJoin_TextBooleanFalse_AcceptsText()
    {
        var func = TextJoinFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString(", "),
            FormulaResult.FromString("FALSE"),
            FormulaResult.FromString("A"),
            FormulaResult.FromString("B"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("A, B", result.StringValue);
    }

    [Fact]
    public void TextJoin_InvalidBoolean_ReturnsError()
    {
        var func = TextJoinFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString(", "),
            FormulaResult.FromString("invalid"),
            FormulaResult.FromString("A"),
            FormulaResult.FromString("B"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void TextJoin_ErrorInDelimiter_PropagatesError()
    {
        var func = TextJoinFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#REF!"),
            FormulaResult.FromBool(true),
            FormulaResult.FromString("A"),
            FormulaResult.FromString("B"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    [Fact]
    public void TextJoin_ErrorInText_PropagatesError()
    {
        var func = TextJoinFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString(", "),
            FormulaResult.FromBool(true),
            FormulaResult.FromString("A"),
            FormulaResult.Error("#N/A"),
            FormulaResult.FromString("B"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    [Fact]
    public void TextJoin_TooFewArguments_ReturnsError()
    {
        var func = TextJoinFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString(", "),
            FormulaResult.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region REVERSE Function Tests

    [Fact]
    public void Reverse_SimpleString_ReturnsReversed()
    {
        var func = ReverseFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Excel"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("lecxE", result.StringValue);
    }

    [Fact]
    public void Reverse_Palindrome_ReturnsSame()
    {
        var func = ReverseFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("racecar"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("racecar", result.StringValue);
    }

    [Fact]
    public void Reverse_EmptyString_ReturnsEmpty()
    {
        var func = ReverseFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString(string.Empty),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(string.Empty, result.StringValue);
    }

    [Fact]
    public void Reverse_SingleCharacter_ReturnsSame()
    {
        var func = ReverseFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("A"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("A", result.StringValue);
    }

    [Fact]
    public void Reverse_Numbers_ReturnsReversed()
    {
        var func = ReverseFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(12345),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("54321", result.StringValue);
    }

    [Fact]
    public void Reverse_WithSpaces_ReversesIncludingSpaces()
    {
        var func = ReverseFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Hello World"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("dlroW olleH", result.StringValue);
    }

    [Fact]
    public void Reverse_ErrorValue_PropagatesError()
    {
        var func = ReverseFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#NAME?"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NAME?", result.ErrorValue);
    }

    [Fact]
    public void Reverse_TooManyArguments_ReturnsError()
    {
        var func = ReverseFunction.Instance;
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
    public void Reverse_NoArguments_ReturnsError()
    {
        var func = ReverseFunction.Instance;
        var args = Array.Empty<FormulaResult>();

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion
}
