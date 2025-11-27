// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;
using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for new text manipulation functions (Excel 365 and DBCS).
/// </summary>
public class NewTextFunctionTests
{
    #region TEXTBEFORE Function Tests

    [Fact]
    public void TextBefore_BasicDelimiter_ReturnsExpected()
    {
        var func = TextBeforeFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Red-Apple"),
            FormulaResult.FromString("-"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("Red", result.StringValue);
    }

    [Fact]
    public void TextBefore_SecondInstance_ReturnsExpected()
    {
        var func = TextBeforeFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("one-two-three"),
            FormulaResult.FromString("-"),
            FormulaResult.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("one-two", result.StringValue);
    }

    [Fact]
    public void TextBefore_CaseInsensitive_ReturnsExpected()
    {
        var func = TextBeforeFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Hello WORLD"),
            FormulaResult.FromString("world"),
            FormulaResult.FromNumber(1),
            FormulaResult.FromNumber(1), // case-insensitive
        };

        var result = func.Execute(null!, args);

        Assert.Equal("Hello ", result.StringValue);
    }

    [Fact]
    public void TextBefore_NotFound_ReturnsError()
    {
        var func = TextBeforeFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Hello"),
            FormulaResult.FromString("XYZ"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    [Fact]
    public void TextBefore_WithIfNotFound_ReturnsCustomValue()
    {
        var func = TextBeforeFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Hello"),
            FormulaResult.FromString("XYZ"),
            FormulaResult.FromNumber(1),
            FormulaResult.FromNumber(0),
            FormulaResult.FromNumber(0),
            FormulaResult.FromString("Not Found"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("Not Found", result.StringValue);
    }

    #endregion

    #region TEXTAFTER Function Tests

    [Fact]
    public void TextAfter_BasicDelimiter_ReturnsExpected()
    {
        var func = TextAfterFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Red-Apple"),
            FormulaResult.FromString("-"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("Apple", result.StringValue);
    }

    [Fact]
    public void TextAfter_SecondInstance_ReturnsExpected()
    {
        var func = TextAfterFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("one-two-three"),
            FormulaResult.FromString("-"),
            FormulaResult.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("three", result.StringValue);
    }

    [Fact]
    public void TextAfter_NotFound_ReturnsError()
    {
        var func = TextAfterFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Hello"),
            FormulaResult.FromString("XYZ"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    #endregion

    #region TEXTSPLIT Function Tests

    [Fact]
    public void TextSplit_BasicSplit_ReturnsFirstElement()
    {
        var func = TextSplitFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("one,two,three"),
            FormulaResult.FromString(","),
        };

        var result = func.Execute(null!, args);

        // Simplified implementation returns first element
        Assert.Equal("one", result.StringValue);
    }

    [Fact]
    public void TextSplit_EmptyText_ReturnsEmpty()
    {
        var func = TextSplitFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString(string.Empty),
            FormulaResult.FromString(","),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(string.Empty, result.StringValue);
    }

    #endregion

    #region VALUETOTEXT Function Tests

    [Fact]
    public void ValueToText_TextValue_ReturnsText()
    {
        var func = ValueToTextFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Hello"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("Hello", result.StringValue);
    }

    [Fact]
    public void ValueToText_NumberValue_ReturnsNumberText()
    {
        var func = ValueToTextFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(123.45),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("123.45", result.StringValue);
    }

    [Fact]
    public void ValueToText_BooleanTrue_ReturnsTRUE()
    {
        var func = ValueToTextFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("TRUE", result.StringValue);
    }

    [Fact]
    public void ValueToText_BooleanFalse_ReturnsFALSE()
    {
        var func = ValueToTextFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromBool(false),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("FALSE", result.StringValue);
    }

    [Fact]
    public void ValueToText_StrictFormat_QuotesText()
    {
        var func = ValueToTextFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Hello"),
            FormulaResult.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("\"Hello\"", result.StringValue);
    }

    #endregion

    #region ARRAYTOTEXT Function Tests

    [Fact]
    public void ArrayToText_TextValue_ReturnsText()
    {
        var func = ArrayToTextFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Test"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("Test", result.StringValue);
    }

    [Fact]
    public void ArrayToText_NumberValue_ReturnsNumberText()
    {
        var func = ArrayToTextFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(42),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("42", result.StringValue);
    }

    #endregion

    #region LENB Function Tests

    [Fact]
    public void LenB_ASCIIText_ReturnsByteCount()
    {
        var func = LenBFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Hello"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void LenB_UnicodeText_ReturnsByteCount()
    {
        var func = LenBFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Hello世界"),
        };

        var result = func.Execute(null!, args);

        // "Hello" = 5 bytes, "世界" = 6 bytes (3 bytes each in UTF-8)
        Assert.Equal(11.0, result.NumericValue);
    }

    [Fact]
    public void LenB_EmptyString_ReturnsZero()
    {
        var func = LenBFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString(string.Empty),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(0.0, result.NumericValue);
    }

    #endregion

    #region LEFTB Function Tests

    [Fact]
    public void LeftB_ASCIIText_ReturnsLeftBytes()
    {
        var func = LeftBFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Hello"),
            FormulaResult.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("Hel", result.StringValue);
    }

    [Fact]
    public void LeftB_DefaultOneChar_ReturnsFirstChar()
    {
        var func = LeftBFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Hello"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("H", result.StringValue);
    }

    [Fact]
    public void LeftB_ZeroBytes_ReturnsEmpty()
    {
        var func = LeftBFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Hello"),
            FormulaResult.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(string.Empty, result.StringValue);
    }

    #endregion

    #region RIGHTB Function Tests

    [Fact]
    public void RightB_ASCIIText_ReturnsRightBytes()
    {
        var func = RightBFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Hello"),
            FormulaResult.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("llo", result.StringValue);
    }

    [Fact]
    public void RightB_DefaultOneChar_ReturnsLastChar()
    {
        var func = RightBFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Hello"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("o", result.StringValue);
    }

    #endregion

    #region MIDB Function Tests

    [Fact]
    public void MidB_ASCIIText_ReturnsMiddleBytes()
    {
        var func = MidBFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Hello World"),
            FormulaResult.FromNumber(7),
            FormulaResult.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("World", result.StringValue);
    }

    [Fact]
    public void MidB_StartBeyondLength_ReturnsEmpty()
    {
        var func = MidBFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Hello"),
            FormulaResult.FromNumber(20),
            FormulaResult.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(string.Empty, result.StringValue);
    }

    [Fact]
    public void MidB_InvalidStartNum_ReturnsError()
    {
        var func = MidBFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Hello"),
            FormulaResult.FromNumber(0),
            FormulaResult.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region FINDB Function Tests

    [Fact]
    public void FindB_BasicSearch_ReturnsPosition()
    {
        var func = FindBFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("World"),
            FormulaResult.FromString("Hello World"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(7.0, result.NumericValue);
    }

    [Fact]
    public void FindB_NotFound_ReturnsError()
    {
        var func = FindBFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("XYZ"),
            FormulaResult.FromString("Hello World"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void FindB_WithStartPosition_ReturnsPosition()
    {
        var func = FindBFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("o"),
            FormulaResult.FromString("Hello World"),
            FormulaResult.FromNumber(6),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(8.0, result.NumericValue);
    }

    #endregion

    #region SEARCHB Function Tests

    [Fact]
    public void SearchB_BasicSearch_ReturnsPosition()
    {
        var func = SearchBFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("world"),
            FormulaResult.FromString("Hello World"),
        };

        var result = func.Execute(null!, args);

        // SEARCHB is case-insensitive
        Assert.Equal(7.0, result.NumericValue);
    }

    [Fact]
    public void SearchB_WithWildcard_ReturnsPosition()
    {
        var func = SearchBFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("W*d"),
            FormulaResult.FromString("Hello World"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(7.0, result.NumericValue);
    }

    [Fact]
    public void SearchB_NotFound_ReturnsError()
    {
        var func = SearchBFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("XYZ"),
            FormulaResult.FromString("Hello World"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region REPLACEB Function Tests

    [Fact]
    public void ReplaceB_BasicReplacement_ReturnsExpected()
    {
        var func = ReplaceBFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Hello World"),
            FormulaResult.FromNumber(7),
            FormulaResult.FromNumber(5),
            FormulaResult.FromString("Excel"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("Hello Excel", result.StringValue);
    }

    [Fact]
    public void ReplaceB_ZeroBytes_InsertsText()
    {
        var func = ReplaceBFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Hello"),
            FormulaResult.FromNumber(6),
            FormulaResult.FromNumber(0),
            FormulaResult.FromString(" World"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("Hello World", result.StringValue);
    }

    [Fact]
    public void ReplaceB_InvalidStartNum_ReturnsError()
    {
        var func = ReplaceBFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Hello"),
            FormulaResult.FromNumber(0),
            FormulaResult.FromNumber(1),
            FormulaResult.FromString("X"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion
}
