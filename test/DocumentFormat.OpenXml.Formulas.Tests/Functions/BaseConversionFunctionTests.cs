// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;
using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for base conversion functions (BASE, DECIMAL, ARABIC, ROMAN).
/// </summary>
public class BaseConversionFunctionTests
{
    // BASE Tests
    [Fact]
    public void Base_Binary_ReturnsCorrectValue()
    {
        var func = BaseFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(2),
        });

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("1010", result.StringValue);
    }

    [Fact]
    public void Base_Hexadecimal_ReturnsCorrectValue()
    {
        var func = BaseFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(255),
            CellValue.FromNumber(16),
        });

        Assert.Equal("FF", result.StringValue);
    }

    [Fact]
    public void Base_WithMinLength_PadsWithZeros()
    {
        var func = BaseFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(2),
            CellValue.FromNumber(8),
        });

        Assert.Equal("00001010", result.StringValue);
    }

    [Fact]
    public void Base_Base36_ReturnsCorrectValue()
    {
        var func = BaseFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1234),
            CellValue.FromNumber(36),
        });

        Assert.Equal("YA", result.StringValue);
    }

    [Fact]
    public void Base_Zero_ReturnsZero()
    {
        var func = BaseFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
            CellValue.FromNumber(10),
        });

        Assert.Equal("0", result.StringValue);
    }

    [Fact]
    public void Base_InvalidRadix_ReturnsError()
    {
        var func = BaseFunction.Instance;

        // Radix < 2
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(1),
        });

        Assert.True(result1.IsError);
        Assert.Equal("#NUM!", result1.ErrorValue);

        // Radix > 36
        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(37),
        });

        Assert.True(result2.IsError);
        Assert.Equal("#NUM!", result2.ErrorValue);
    }

    [Fact]
    public void Base_NegativeNumber_ReturnsError()
    {
        var func = BaseFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-10),
            CellValue.FromNumber(2),
        });

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    // DECIMAL Tests
    [Fact]
    public void Decimal_Binary_ReturnsCorrectValue()
    {
        var func = DecimalFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromString("1010"),
            CellValue.FromNumber(2),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void Decimal_Hexadecimal_ReturnsCorrectValue()
    {
        var func = DecimalFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromString("FF"),
            CellValue.FromNumber(16),
        });

        Assert.Equal(255.0, result.NumericValue);
    }

    [Fact]
    public void Decimal_Base36_ReturnsCorrectValue()
    {
        var func = DecimalFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromString("YA"),
            CellValue.FromNumber(36),
        });

        Assert.Equal(1234.0, result.NumericValue);
    }

    [Fact]
    public void Decimal_CaseInsensitive_ReturnsCorrectValue()
    {
        var func = DecimalFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromString("ff"),
            CellValue.FromNumber(16),
        });

        Assert.Equal(255.0, result.NumericValue);
    }

    [Fact]
    public void Decimal_InvalidCharacter_ReturnsError()
    {
        var func = DecimalFunction.Instance;

        // "G" is not valid in base 16
        var result = func.Execute(null!, new[]
        {
            CellValue.FromString("FG"),
            CellValue.FromNumber(16),
        });

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Decimal_InvalidRadix_ReturnsError()
    {
        var func = DecimalFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromString("10"),
            CellValue.FromNumber(1),
        });

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    // ARABIC Tests
    [Fact]
    public void Arabic_SimpleRoman_ReturnsCorrectValue()
    {
        var func = ArabicFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromString("IV"),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(4.0, result.NumericValue);
    }

    [Fact]
    public void Arabic_ComplexRoman_ReturnsCorrectValue()
    {
        var func = ArabicFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromString("MCMXCIV"),
        });

        Assert.Equal(1994.0, result.NumericValue);
    }

    [Fact]
    public void Arabic_CaseInsensitive_ReturnsCorrectValue()
    {
        var func = ArabicFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromString("mcmxciv"),
        });

        Assert.Equal(1994.0, result.NumericValue);
    }

    [Fact]
    public void Arabic_SimpleValues_ReturnCorrectValues()
    {
        var func = ArabicFunction.Instance;

        var tests = new[]
        {
            ("I", 1.0),
            ("V", 5.0),
            ("X", 10.0),
            ("L", 50.0),
            ("C", 100.0),
            ("D", 500.0),
            ("M", 1000.0),
            ("IX", 9.0),
            ("XL", 40.0),
            ("XC", 90.0),
            ("CD", 400.0),
            ("CM", 900.0),
        };

        foreach (var (roman, expected) in tests)
        {
            var result = func.Execute(null!, new[]
            {
                CellValue.FromString(roman),
            });

            Assert.Equal(expected, result.NumericValue);
        }
    }

    [Fact]
    public void Arabic_InvalidCharacter_ReturnsError()
    {
        var func = ArabicFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromString("ABC"),
        });

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Arabic_EmptyString_ReturnsError()
    {
        var func = ArabicFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromString(""),
        });

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    // ROMAN Tests
    [Fact]
    public void Roman_SimpleNumber_ReturnsCorrectValue()
    {
        var func = RomanFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(4),
        });

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("IV", result.StringValue);
    }

    [Fact]
    public void Roman_ComplexNumber_ReturnsCorrectValue()
    {
        var func = RomanFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1994),
        });

        Assert.Equal("MCMXCIV", result.StringValue);
    }

    [Fact]
    public void Roman_Zero_ReturnsEmptyString()
    {
        var func = RomanFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
        });

        Assert.Equal(string.Empty, result.StringValue);
    }

    [Fact]
    public void Roman_WithForm_ReturnsCorrectValue()
    {
        var func = RomanFunction.Instance;

        // Form 0 is classic
        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(499),
            CellValue.FromNumber(0),
        });

        Assert.Equal("CDXCIX", result.StringValue);
    }

    [Fact]
    public void Roman_SimpleValues_ReturnCorrectValues()
    {
        var func = RomanFunction.Instance;

        var tests = new[]
        {
            (1, "I"),
            (5, "V"),
            (10, "X"),
            (50, "L"),
            (100, "C"),
            (500, "D"),
            (1000, "M"),
            (9, "IX"),
            (40, "XL"),
            (90, "XC"),
            (400, "CD"),
            (900, "CM"),
        };

        foreach (var (number, expected) in tests)
        {
            var result = func.Execute(null!, new[]
            {
                CellValue.FromNumber(number),
            });

            Assert.Equal(expected, result.StringValue);
        }
    }

    [Fact]
    public void Roman_OutOfRange_ReturnsError()
    {
        var func = RomanFunction.Instance;

        // Negative number
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-1),
        });

        Assert.True(result1.IsError);
        Assert.Equal("#VALUE!", result1.ErrorValue);

        // Number > 3999
        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(4000),
        });

        Assert.True(result2.IsError);
        Assert.Equal("#VALUE!", result2.ErrorValue);
    }

    [Fact]
    public void Roman_InvalidForm_ReturnsError()
    {
        var func = RomanFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(100),
            CellValue.FromNumber(5),
        });

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }
}
