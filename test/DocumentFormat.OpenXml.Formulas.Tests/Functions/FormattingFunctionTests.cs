// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;
using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for text formatting and number conversion functions.
/// </summary>
public class FormattingFunctionTests
{
    #region FIXED Function Tests

    [Fact]
    public void Fixed_BasicFormatting_ReturnsExpected()
    {
        var func = FixedFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1234.567),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("1,234.6", result.StringValue);
    }

    [Fact]
    public void Fixed_WithoutCommas_ReturnsExpected()
    {
        var func = FixedFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1234.567),
            CellValue.FromNumber(1),
            CellValue.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("1234.6", result.StringValue);
    }

    [Fact]
    public void Fixed_DefaultDecimals_ReturnsTwoDecimals()
    {
        var func = FixedFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1234.567),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("1,234.57", result.StringValue);
    }

    [Fact]
    public void Fixed_ZeroDecimals_ReturnsRounded()
    {
        var func = FixedFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1234.567),
            CellValue.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("1,235", result.StringValue);
    }

    [Fact]
    public void Fixed_ThreeDecimals_ReturnsExpected()
    {
        var func = FixedFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1234.56789),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("1,234.568", result.StringValue);
    }

    [Fact]
    public void Fixed_NegativeNumber_ReturnsExpected()
    {
        var func = FixedFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-1234.567),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("-1,234.57", result.StringValue);
    }

    [Fact]
    public void Fixed_NegativeDecimals_UsesZero()
    {
        var func = FixedFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1234.567),
            CellValue.FromNumber(-1),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("1,235", result.StringValue);
    }

    [Fact]
    public void Fixed_NoCommasFalse_IncludesCommas()
    {
        var func = FixedFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1234.567),
            CellValue.FromNumber(2),
            CellValue.FromBool(false),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("1,234.57", result.StringValue);
    }

    [Fact]
    public void Fixed_NoCommasAsNumber_InterpretsCorrectly()
    {
        var func = FixedFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1234.567),
            CellValue.FromNumber(2),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("1234.57", result.StringValue);
    }

    [Fact]
    public void Fixed_SmallNumber_ReturnsExpected()
    {
        var func = FixedFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(12.5),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("12.5", result.StringValue);
    }

    [Fact]
    public void Fixed_LargeNumber_ReturnsExpected()
    {
        var func = FixedFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1234567.89),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("1,234,567.89", result.StringValue);
    }

    [Fact]
    public void Fixed_NonNumericInput_ReturnsError()
    {
        var func = FixedFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("text"),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Fixed_ErrorValue_PropagatesError()
    {
        var func = FixedFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void Fixed_WrongArgumentCount_ReturnsError()
    {
        var func = FixedFunction.Instance;
        var args = Array.Empty<CellValue>();

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region DOLLAR Function Tests

    [Fact]
    public void Dollar_BasicFormatting_ReturnsExpected()
    {
        var func = DollarFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1234.567),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("$1,234.57", result.StringValue);
    }

    [Fact]
    public void Dollar_DefaultDecimals_ReturnsTwoDecimals()
    {
        var func = DollarFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1234.567),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("$1,234.57", result.StringValue);
    }

    [Fact]
    public void Dollar_ZeroDecimals_ReturnsRounded()
    {
        var func = DollarFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1234.567),
            CellValue.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("$1,235", result.StringValue);
    }

    [Fact]
    public void Dollar_ThreeDecimals_ReturnsExpected()
    {
        var func = DollarFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1234.56789),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("$1,234.568", result.StringValue);
    }

    [Fact]
    public void Dollar_NegativeNumber_ReturnsExpected()
    {
        var func = DollarFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-1234.567),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("$-1,234.57", result.StringValue);
    }

    [Fact]
    public void Dollar_NegativeDecimals_UsesZero()
    {
        var func = DollarFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1234.567),
            CellValue.FromNumber(-1),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("$1,235", result.StringValue);
    }

    [Fact]
    public void Dollar_SmallNumber_ReturnsExpected()
    {
        var func = DollarFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(12.5),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("$12.50", result.StringValue);
    }

    [Fact]
    public void Dollar_LargeNumber_ReturnsExpected()
    {
        var func = DollarFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1234567.89),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("$1,234,567.89", result.StringValue);
    }

    [Fact]
    public void Dollar_OneDecimal_ReturnsExpected()
    {
        var func = DollarFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(99.99),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("$100.0", result.StringValue);
    }

    [Fact]
    public void Dollar_NonNumericInput_ReturnsError()
    {
        var func = DollarFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("text"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Dollar_ErrorValue_PropagatesError()
    {
        var func = DollarFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#N/A"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    [Fact]
    public void Dollar_WrongArgumentCount_ReturnsError()
    {
        var func = DollarFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(100),
            CellValue.FromNumber(2),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region NUMBERVALUE Function Tests

    [Fact]
    public void NumberValue_CustomSeparators_ReturnsExpected()
    {
        var func = NumberValueFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("1.234,56"),
            CellValue.FromString(","),
            CellValue.FromString("."),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1234.56, result.NumericValue);
    }

    [Fact]
    public void NumberValue_DefaultSeparators_ReturnsExpected()
    {
        var func = NumberValueFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("1,234.56"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(1234.56, result.NumericValue);
    }

    [Fact]
    public void NumberValue_EuropeanFormat_ReturnsExpected()
    {
        var func = NumberValueFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("2.500,75"),
            CellValue.FromString(","),
            CellValue.FromString("."),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(2500.75, result.NumericValue);
    }

    [Fact]
    public void NumberValue_NoGroupSeparator_ReturnsExpected()
    {
        var func = NumberValueFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("1234.56"),
            CellValue.FromString("."),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(1234.56, result.NumericValue);
    }

    [Fact]
    public void NumberValue_WithPercentage_ReturnsExpected()
    {
        var func = NumberValueFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("50%"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(0.5, result.NumericValue);
    }

    [Fact]
    public void NumberValue_PercentageWithDecimals_ReturnsExpected()
    {
        var func = NumberValueFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("12.5%"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(0.125, result.NumericValue);
    }

    [Fact]
    public void NumberValue_NegativeNumber_ReturnsExpected()
    {
        var func = NumberValueFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("-1,234.56"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(-1234.56, result.NumericValue);
    }

    [Fact]
    public void NumberValue_NegativeWithCustomSeparators_ReturnsExpected()
    {
        var func = NumberValueFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("-1.234,56"),
            CellValue.FromString(","),
            CellValue.FromString("."),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(-1234.56, result.NumericValue);
    }

    [Fact]
    public void NumberValue_SpaceSeparator_ReturnsExpected()
    {
        var func = NumberValueFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("1 234,56"),
            CellValue.FromString(","),
            CellValue.FromString(" "),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(1234.56, result.NumericValue);
    }

    [Fact]
    public void NumberValue_SimpleNumber_ReturnsExpected()
    {
        var func = NumberValueFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("123"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(123.0, result.NumericValue);
    }

    [Fact]
    public void NumberValue_DecimalOnly_ReturnsExpected()
    {
        var func = NumberValueFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("0.5"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(0.5, result.NumericValue);
    }

    [Fact]
    public void NumberValue_WithWhitespace_ReturnsExpected()
    {
        var func = NumberValueFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("  1,234.56  "),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(1234.56, result.NumericValue);
    }

    [Fact]
    public void NumberValue_EmptyString_ReturnsError()
    {
        var func = NumberValueFunction.Instance;
        var args = new[]
        {
            CellValue.FromString(string.Empty),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void NumberValue_InvalidText_ReturnsError()
    {
        var func = NumberValueFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("not a number"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void NumberValue_SameSeparators_ReturnsError()
    {
        var func = NumberValueFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("1,234,56"),
            CellValue.FromString(","),
            CellValue.FromString(","),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void NumberValue_ErrorValue_PropagatesError()
    {
        var func = NumberValueFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#REF!"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    [Fact]
    public void NumberValue_WrongArgumentCount_ReturnsError()
    {
        var func = NumberValueFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("123"),
            CellValue.FromString("."),
            CellValue.FromString(","),
            CellValue.FromString("extra"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion
}
