// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for engineering functions.
/// </summary>
public class EngineeringFunctionTests
{
    #region CONVERT Tests

    [Fact]
    public void Convert_MetersToFeet_ReturnsCorrectValue()
    {
        var func = ConvertFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromString("m"),
            CellValue.FromString("ft"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(3.28084, result.NumericValue, 5);
    }

    [Fact]
    public void Convert_KilogramsToGrams_ReturnsCorrectValue()
    {
        var func = ConvertFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromString("kg"),
            CellValue.FromString("g"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1000.0, result.NumericValue);
    }

    [Fact]
    public void Convert_CelsiusToFahrenheit_ReturnsCorrectValue()
    {
        var func = ConvertFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0),
            CellValue.FromString("C"),
            CellValue.FromString("F"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(32.0, result.NumericValue);
    }

    [Fact]
    public void Convert_FahrenheitToCelsius_ReturnsCorrectValue()
    {
        var func = ConvertFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(32),
            CellValue.FromString("F"),
            CellValue.FromString("C"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.0, result.NumericValue, 5);
    }

    [Fact]
    public void Convert_CelsiusToKelvin_ReturnsCorrectValue()
    {
        var func = ConvertFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0),
            CellValue.FromString("C"),
            CellValue.FromString("K"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(273.15, result.NumericValue);
    }

    [Fact]
    public void Convert_GallonsToLiters_ReturnsCorrectValue()
    {
        var func = ConvertFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromString("gal"),
            CellValue.FromString("l"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(3.785411784, result.NumericValue, 5);
    }

    [Fact]
    public void Convert_IncompatibleUnits_ReturnsNA()
    {
        var func = ConvertFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromString("m"),
            CellValue.FromString("kg"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    [Fact]
    public void Convert_InvalidUnit_ReturnsNA()
    {
        var func = ConvertFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromString("xyz"),
            CellValue.FromString("m"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    [Fact]
    public void Convert_WrongArgumentCount_ReturnsError()
    {
        var func = ConvertFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromString("m"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region HEX2DEC Tests

    [Fact]
    public void Hex2Dec_FF_Returns255()
    {
        var func = Hex2DecFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("FF"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(255.0, result.NumericValue);
    }

    [Fact]
    public void Hex2Dec_A_Returns10()
    {
        var func = Hex2DecFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("A"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void Hex2Dec_LowerCase_ReturnsCorrectValue()
    {
        var func = Hex2DecFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("ff"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(255.0, result.NumericValue);
    }

    [Fact]
    public void Hex2Dec_InvalidCharacter_ReturnsError()
    {
        var func = Hex2DecFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("FG"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Hex2Dec_TooLong_ReturnsError()
    {
        var func = Hex2DecFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("12345678901"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    #endregion

    #region DEC2HEX Tests

    [Fact]
    public void Dec2Hex_255_ReturnsFF()
    {
        var func = Dec2HexFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(255),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("FF", result.StringValue);
    }

    [Fact]
    public void Dec2Hex_10_ReturnsA()
    {
        var func = Dec2HexFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("A", result.StringValue);
    }

    [Fact]
    public void Dec2Hex_WithPlaces_ReturnsPadded()
    {
        var func = Dec2HexFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(4),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("000A", result.StringValue);
    }

    [Fact]
    public void Dec2Hex_Negative_ReturnsComplement()
    {
        var func = Dec2HexFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-1),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("FFFFFFFFFF", result.StringValue);
    }

    [Fact]
    public void Dec2Hex_OutOfRange_ReturnsError()
    {
        var func = Dec2HexFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(549755813888),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    #endregion

    #region BIN2DEC Tests

    [Fact]
    public void Bin2Dec_1010_Returns10()
    {
        var func = Bin2DecFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("1010"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void Bin2Dec_1_Returns1()
    {
        var func = Bin2DecFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("1"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void Bin2Dec_0_Returns0()
    {
        var func = Bin2DecFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("0"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void Bin2Dec_InvalidCharacter_ReturnsError()
    {
        var func = Bin2DecFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("102"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Bin2Dec_TooLong_ReturnsError()
    {
        var func = Bin2DecFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("10101010101"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    #endregion

    #region DEC2BIN Tests

    [Fact]
    public void Dec2Bin_10_Returns1010()
    {
        var func = Dec2BinFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("1010", result.StringValue);
    }

    [Fact]
    public void Dec2Bin_1_Returns1()
    {
        var func = Dec2BinFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("1", result.StringValue);
    }

    [Fact]
    public void Dec2Bin_WithPlaces_ReturnsPadded()
    {
        var func = Dec2BinFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(8),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("00001010", result.StringValue);
    }

    [Fact]
    public void Dec2Bin_Negative_ReturnsComplement()
    {
        var func = Dec2BinFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-1),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("1111111111", result.StringValue);
    }

    [Fact]
    public void Dec2Bin_OutOfRange_ReturnsError()
    {
        var func = Dec2BinFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(512),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    #endregion

    #region OCT2DEC Tests

    [Fact]
    public void Oct2Dec_77_Returns63()
    {
        var func = Oct2DecFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("77"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(63.0, result.NumericValue);
    }

    [Fact]
    public void Oct2Dec_10_Returns8()
    {
        var func = Oct2DecFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("10"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(8.0, result.NumericValue);
    }

    [Fact]
    public void Oct2Dec_InvalidCharacter_ReturnsError()
    {
        var func = Oct2DecFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("78"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Oct2Dec_TooLong_ReturnsError()
    {
        var func = Oct2DecFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("12345678901"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    #endregion

    #region DEC2OCT Tests

    [Fact]
    public void Dec2Oct_63_Returns77()
    {
        var func = Dec2OctFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(63),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("77", result.StringValue);
    }

    [Fact]
    public void Dec2Oct_8_Returns10()
    {
        var func = Dec2OctFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(8),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("10", result.StringValue);
    }

    [Fact]
    public void Dec2Oct_WithPlaces_ReturnsPadded()
    {
        var func = Dec2OctFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(8),
            CellValue.FromNumber(4),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("0010", result.StringValue);
    }

    [Fact]
    public void Dec2Oct_Negative_ReturnsComplement()
    {
        var func = Dec2OctFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-1),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("7777777777", result.StringValue);
    }

    [Fact]
    public void Dec2Oct_OutOfRange_ReturnsError()
    {
        var func = Dec2OctFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(536870912),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    #endregion

    #region Complex Number Tests

    [Fact]
    public void Complex_BasicCreation_ReturnsComplexNumber()
    {
        var func = ComplexFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(3),
            CellValue.FromNumber(4),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("3+4i", result.StringValue);
    }

    [Fact]
    public void Complex_WithJSuffix_ReturnsComplexNumberWithJ()
    {
        var func = ComplexFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(3),
            CellValue.FromNumber(4),
            CellValue.FromString("j"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("3+4j", result.StringValue);
    }

    [Fact]
    public void Complex_NegativeImaginary_ReturnsCorrectFormat()
    {
        var func = ComplexFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(3),
            CellValue.FromNumber(-4),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("3-4i", result.StringValue);
    }

    [Fact]
    public void ImReal_ReturnsRealPart()
    {
        var func = ImRealFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("3+4i"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(3.0, result.NumericValue);
    }

    [Fact]
    public void Imaginary_ReturnsImaginaryPart()
    {
        var func = ImaginaryFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("3+4i"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(4.0, result.NumericValue);
    }

    [Fact]
    public void ImAbs_ReturnsModulus()
    {
        var func = ImAbsFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("3+4i"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(5.0, result.NumericValue, 10);
    }

    [Fact]
    public void ImArgument_ReturnsAngle()
    {
        var func = ImArgumentFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("1+i"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.7853981633974483, result.NumericValue, 10);
    }

    [Fact]
    public void ImConjugate_ReturnsConjugate()
    {
        var func = ImConjugateFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("3+4i"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("3-4i", result.StringValue);
    }

    [Fact]
    public void ImSum_AddsComplexNumbers()
    {
        var func = ImSumFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("3+4i"),
            CellValue.FromString("1+2i"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("4+6i", result.StringValue);
    }

    [Fact]
    public void ImSub_SubtractsComplexNumbers()
    {
        var func = ImSubFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("5+7i"),
            CellValue.FromString("2+3i"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("3+4i", result.StringValue);
    }

    [Fact]
    public void ImProduct_MultipliesComplexNumbers()
    {
        var func = ImProductFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("2+3i"),
            CellValue.FromString("4+5i"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("-7+22i", result.StringValue);
    }

    [Fact]
    public void ImDiv_DividesComplexNumbers()
    {
        var func = ImDivFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("1+i"),
            CellValue.FromString("1-i"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("i", result.StringValue);
    }

    [Fact]
    public void ImDiv_DivisionByZero_ReturnsError()
    {
        var func = ImDivFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("1+i"),
            CellValue.FromString("0+0i"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void ImPower_RaisesToPower()
    {
        var func = ImPowerFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("1+i"),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("2i", result.StringValue);
    }

    [Fact]
    public void ImSqrt_ReturnsSquareRoot()
    {
        var func = ImSqrtFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("-1+0i"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("i", result.StringValue);
    }

    [Fact]
    public void ImExp_ReturnsExponential()
    {
        var func = ImExpFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("0+0i"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("1", result.StringValue);
    }

    [Fact]
    public void ImLn_ReturnsNaturalLog()
    {
        var func = ImLnFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("i"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Contains("1.5707963267948966i", result.StringValue);
    }

    [Fact]
    public void ImSin_ReturnsSine()
    {
        var func = ImSinFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("0+0i"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("0", result.StringValue);
    }

    [Fact]
    public void ImCos_ReturnsCosine()
    {
        var func = ImCosFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("0+0i"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("1", result.StringValue);
    }

    [Fact]
    public void Complex_InvalidSuffix_ReturnsError()
    {
        var func = ComplexFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(3),
            CellValue.FromNumber(4),
            CellValue.FromString("k"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void ImReal_InvalidComplexNumber_ReturnsError()
    {
        var func = ImRealFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("not a complex number"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    #endregion

    #region Bitwise Operation Tests

    [Fact]
    public void BitAnd_BasicOperation_ReturnsResult()
    {
        var func = BitAndFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void BitOr_BasicOperation_ReturnsResult()
    {
        var func = BitOrFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(7.0, result.NumericValue);
    }

    [Fact]
    public void BitXor_BasicOperation_ReturnsResult()
    {
        var func = BitXorFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(6.0, result.NumericValue);
    }

    [Fact]
    public void BitLShift_BasicOperation_ReturnsResult()
    {
        var func = BitLShiftFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(20.0, result.NumericValue);
    }

    [Fact]
    public void BitRShift_BasicOperation_ReturnsResult()
    {
        var func = BitRShiftFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(20),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void BitAnd_NegativeNumber_ReturnsError()
    {
        var func = BitAndFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-5),
            CellValue.FromNumber(3),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void BitOr_TooLargeNumber_ReturnsError()
    {
        var func = BitOrFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(281474976710656),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void BitLShift_NegativeShift_ShiftsRight()
    {
        var func = BitLShiftFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(20),
            CellValue.FromNumber(-2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void BitXor_SameNumber_ReturnsZero()
    {
        var func = BitXorFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(42),
            CellValue.FromNumber(42),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.0, result.NumericValue);
    }

    #endregion
}
