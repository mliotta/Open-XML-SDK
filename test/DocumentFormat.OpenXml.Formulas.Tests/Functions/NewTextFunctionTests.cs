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
            CellValue.FromString("Red-Apple"),
            CellValue.FromString("-"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("Red", result.StringValue);
    }

    [Fact]
    public void TextBefore_SecondInstance_ReturnsExpected()
    {
        var func = TextBeforeFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("one-two-three"),
            CellValue.FromString("-"),
            CellValue.FromNumber(2),
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
            CellValue.FromString("Hello WORLD"),
            CellValue.FromString("world"),
            CellValue.FromNumber(1),
            CellValue.FromNumber(1), // case-insensitive
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
            CellValue.FromString("Hello"),
            CellValue.FromString("XYZ"),
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
            CellValue.FromString("Hello"),
            CellValue.FromString("XYZ"),
            CellValue.FromNumber(1),
            CellValue.FromNumber(0),
            CellValue.FromNumber(0),
            CellValue.FromString("Not Found"),
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
            CellValue.FromString("Red-Apple"),
            CellValue.FromString("-"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("Apple", result.StringValue);
    }

    [Fact]
    public void TextAfter_SecondInstance_ReturnsExpected()
    {
        var func = TextAfterFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("one-two-three"),
            CellValue.FromString("-"),
            CellValue.FromNumber(2),
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
            CellValue.FromString("Hello"),
            CellValue.FromString("XYZ"),
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
            CellValue.FromString("one,two,three"),
            CellValue.FromString(","),
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
            CellValue.FromString(string.Empty),
            CellValue.FromString(","),
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
            CellValue.FromString("Hello"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("Hello", result.StringValue);
    }

    [Fact]
    public void ValueToText_NumberValue_ReturnsNumberText()
    {
        var func = ValueToTextFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(123.45),
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
            CellValue.FromBool(true),
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
            CellValue.FromBool(false),
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
            CellValue.FromString("Hello"),
            CellValue.FromNumber(1),
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
            CellValue.FromString("Test"),
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
            CellValue.FromNumber(42),
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
            CellValue.FromString("Hello"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void LenB_UnicodeText_ReturnsByteCount()
    {
        var func = LenBFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("Hello世界"),
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
            CellValue.FromString(string.Empty),
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
            CellValue.FromString("Hello"),
            CellValue.FromNumber(3),
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
            CellValue.FromString("Hello"),
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
            CellValue.FromString("Hello"),
            CellValue.FromNumber(0),
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
            CellValue.FromString("Hello"),
            CellValue.FromNumber(3),
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
            CellValue.FromString("Hello"),
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
            CellValue.FromString("Hello World"),
            CellValue.FromNumber(7),
            CellValue.FromNumber(5),
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
            CellValue.FromString("Hello"),
            CellValue.FromNumber(20),
            CellValue.FromNumber(5),
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
            CellValue.FromString("Hello"),
            CellValue.FromNumber(0),
            CellValue.FromNumber(5),
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
            CellValue.FromString("World"),
            CellValue.FromString("Hello World"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(7.0, result.NumericValue);
    }

    [Fact]
    public void FindB_NotFound_ReturnsError()
    {
        var func = FindBFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("XYZ"),
            CellValue.FromString("Hello World"),
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
            CellValue.FromString("o"),
            CellValue.FromString("Hello World"),
            CellValue.FromNumber(6),
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
            CellValue.FromString("world"),
            CellValue.FromString("Hello World"),
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
            CellValue.FromString("W*d"),
            CellValue.FromString("Hello World"),
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
            CellValue.FromString("XYZ"),
            CellValue.FromString("Hello World"),
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
            CellValue.FromString("Hello World"),
            CellValue.FromNumber(7),
            CellValue.FromNumber(5),
            CellValue.FromString("Excel"),
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
            CellValue.FromString("Hello"),
            CellValue.FromNumber(6),
            CellValue.FromNumber(0),
            CellValue.FromString(" World"),
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
            CellValue.FromString("Hello"),
            CellValue.FromNumber(0),
            CellValue.FromNumber(1),
            CellValue.FromString("X"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion
}
