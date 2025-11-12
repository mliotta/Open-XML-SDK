// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for web functions and extended array functions.
/// </summary>
public class WebAndExtendedArrayFunctionTests
{
    #region Web Function Tests

    [Fact]
    public void EncodeUrl_BasicString_ReturnsEncoded()
    {
        var func = EncodeUrlFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("hello world"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("hello%20world", result.StringValue);
    }

    [Fact]
    public void EncodeUrl_SpecialCharacters_ReturnsEncoded()
    {
        var func = EncodeUrlFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("hello@example.com?foo=bar&baz=qux"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.True(result.StringValue.Contains("%40")); // @ encoded
        Assert.True(result.StringValue.Contains("%3F")); // ? encoded
    }

    [Fact]
    public void EncodeUrl_EmptyString_ReturnsEmpty()
    {
        var func = EncodeUrlFunction.Instance;
        var args = new[]
        {
            CellValue.FromString(string.Empty),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal(string.Empty, result.StringValue);
    }

    [Fact]
    public void EncodeUrl_Error_PropagatesError()
    {
        var func = EncodeUrlFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#N/A"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    [Fact]
    public void WebService_AnyUrl_ReturnsError()
    {
        var func = WebServiceFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("http://example.com"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void FilterXml_SimpleXPath_ReturnsValue()
    {
        var func = FilterXmlFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("<root><value>42</value></root>"),
            CellValue.FromString("/root/value"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(42.0, result.NumericValue);
    }

    [Fact]
    public void FilterXml_TextValue_ReturnsText()
    {
        var func = FilterXmlFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("<root><name>John</name></root>"),
            CellValue.FromString("/root/name"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("John", result.StringValue);
    }

    [Fact]
    public void FilterXml_InvalidXml_ReturnsError()
    {
        var func = FilterXmlFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("not xml"),
            CellValue.FromString("/root"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void FilterXml_NoMatch_ReturnsNA()
    {
        var func = FilterXmlFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("<root><value>42</value></root>"),
            CellValue.FromString("/root/missing"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    #endregion

    #region Array Manipulation Tests

    [Fact]
    public void SortBy_BasicSort_ReturnsFirstElement()
    {
        var func = SortByFunction.Instance;
        // Array: [10, 20, 30]
        // By_array: [3, 1, 2]
        // Result should sort to: [20, 30, 10] (by values 1, 2, 3)
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
            CellValue.FromNumber(3),
            CellValue.FromNumber(1),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(20.0, result.NumericValue); // First element after sorting
    }

    [Fact]
    public void Take_PositiveRows_ReturnsFirstElement()
    {
        var func = TakeFunction.Instance;
        // TAKE([10, 20, 30], 2) should return [10, 20]
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
            CellValue.FromNumber(2), // rows
        };

        var result = func.Execute(null!, args);

        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void Take_NegativeRows_ReturnsTailElement()
    {
        var func = TakeFunction.Instance;
        // TAKE([10, 20, 30], -2) should return [20, 30]
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
            CellValue.FromNumber(-2), // rows (from end)
        };

        var result = func.Execute(null!, args);

        Assert.Equal(20.0, result.NumericValue);
    }

    [Fact]
    public void Drop_PositiveRows_ReturnsRemainingFirst()
    {
        var func = DropFunction.Instance;
        // DROP([10, 20, 30], 1) should return [20, 30]
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
            CellValue.FromNumber(1), // rows to drop
        };

        var result = func.Execute(null!, args);

        Assert.Equal(20.0, result.NumericValue);
    }

    [Fact]
    public void Drop_NegativeRows_ReturnsRemainingFirst()
    {
        var func = DropFunction.Instance;
        // DROP([10, 20, 30], -1) should return [10, 20]
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
            CellValue.FromNumber(-1), // rows to drop from end
        };

        var result = func.Execute(null!, args);

        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void VStack_MultipleArrays_ReturnsFirstElement()
    {
        var func = VStackFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void HStack_MultipleArrays_ReturnsFirstElement()
    {
        var func = HStackFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void ToCol_BasicArray_ReturnsFirstElement()
    {
        var func = ToColFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void ToCol_IgnoreBlanks_SkipsBlanks()
    {
        var func = ToColFunction.Instance;
        var args = new[]
        {
            CellValue.Empty,
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
            CellValue.FromNumber(1), // ignore blanks
        };

        var result = func.Execute(null!, args);

        Assert.Equal(20.0, result.NumericValue); // First non-blank
    }

    [Fact]
    public void ToRow_BasicArray_ReturnsFirstElement()
    {
        var func = ToRowFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void ChooseCols_FirstColumn_ReturnsFirstElement()
    {
        var func = ChooseColsFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
            CellValue.FromNumber(1), // column number
        };

        var result = func.Execute(null!, args);

        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void ChooseRows_FirstRow_ReturnsFirstElement()
    {
        var func = ChooseRowsFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
            CellValue.FromNumber(1), // row number
        };

        var result = func.Execute(null!, args);

        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void Expand_BasicArray_ReturnsFirstElement()
    {
        var func = ExpandFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(5), // rows
        };

        var result = func.Execute(null!, args);

        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void WrapCols_BasicArray_ReturnsFirstElement()
    {
        var func = WrapColsFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
            CellValue.FromNumber(2), // wrap_count
        };

        var result = func.Execute(null!, args);

        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void WrapRows_BasicArray_ReturnsFirstElement()
    {
        var func = WrapRowsFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
            CellValue.FromNumber(2), // wrap_count
        };

        var result = func.Execute(null!, args);

        Assert.Equal(10.0, result.NumericValue);
    }

    #endregion
}
