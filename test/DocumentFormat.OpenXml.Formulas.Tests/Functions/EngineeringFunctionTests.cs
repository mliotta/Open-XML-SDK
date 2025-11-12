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
}
