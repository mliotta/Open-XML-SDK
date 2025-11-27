// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for modern lookup functions: XLOOKUP, XMATCH, and related reference functions.
/// </summary>
public class ModernLookupFunctionTests
{
    #region XLOOKUP Function Tests

    [Fact]
    public void XLookup_ExactMatch_ReturnsCorrectValue()
    {
        var func = XLookupFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Banana"), // lookup_value
            FormulaResult.FromString("Apple"),  // lookup_array
            FormulaResult.FromString("Banana"),
            FormulaResult.FromString("Cherry"),
            FormulaResult.FromNumber(100), // return_array
            FormulaResult.FromNumber(200),
            FormulaResult.FromNumber(300),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(200.0, result.NumericValue);
    }

    [Fact]
    public void XLookup_NoMatch_ReturnsIfNotFound()
    {
        var func = XLookupFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Orange"), // lookup_value
            FormulaResult.FromString("Apple"),  // lookup_array
            FormulaResult.FromString("Banana"),
            FormulaResult.FromString("Cherry"),
            FormulaResult.FromNumber(100), // return_array
            FormulaResult.FromNumber(200),
            FormulaResult.FromNumber(300),
            FormulaResult.FromString("Not Found"), // if_not_found
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("Not Found", result.StringValue);
    }

    [Fact]
    public void XLookup_ExactOrNextSmaller_ReturnsCorrectValue()
    {
        var func = XLookupFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(25), // lookup_value
            FormulaResult.FromNumber(10), // lookup_array
            FormulaResult.FromNumber(20),
            FormulaResult.FromNumber(30),
            FormulaResult.FromNumber(40),
            FormulaResult.FromString("A"), // return_array
            FormulaResult.FromString("B"),
            FormulaResult.FromString("C"),
            FormulaResult.FromString("D"),
            FormulaResult.Error("#N/A"), // if_not_found
            FormulaResult.FromNumber(-1), // match_mode (exact or next smaller)
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("B", result.StringValue); // 20 is the next smaller value
    }

    [Fact]
    public void XLookup_ExactOrNextLarger_ReturnsCorrectValue()
    {
        var func = XLookupFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(25), // lookup_value
            FormulaResult.FromNumber(10), // lookup_array
            FormulaResult.FromNumber(20),
            FormulaResult.FromNumber(30),
            FormulaResult.FromNumber(40),
            FormulaResult.FromString("A"), // return_array
            FormulaResult.FromString("B"),
            FormulaResult.FromString("C"),
            FormulaResult.FromString("D"),
            FormulaResult.Error("#N/A"), // if_not_found
            FormulaResult.FromNumber(1), // match_mode (exact or next larger)
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("C", result.StringValue); // 30 is the next larger value
    }

    [Fact]
    public void XLookup_WildcardMatch_ReturnsCorrectValue()
    {
        var func = XLookupFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("B*"), // lookup_value with wildcard
            FormulaResult.FromString("Apple"), // lookup_array
            FormulaResult.FromString("Banana"),
            FormulaResult.FromString("Cherry"),
            FormulaResult.FromNumber(100), // return_array
            FormulaResult.FromNumber(200),
            FormulaResult.FromNumber(300),
            FormulaResult.Error("#N/A"), // if_not_found
            FormulaResult.FromNumber(2), // match_mode (wildcard)
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(200.0, result.NumericValue); // Matches "Banana"
    }

    [Fact]
    public void XLookup_SearchLastToFirst_ReturnsLastMatch()
    {
        var func = XLookupFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Apple"), // lookup_value
            FormulaResult.FromString("Apple"), // lookup_array (duplicate)
            FormulaResult.FromString("Banana"),
            FormulaResult.FromString("Apple"), // Another Apple
            FormulaResult.FromNumber(100), // return_array
            FormulaResult.FromNumber(200),
            FormulaResult.FromNumber(300),
            FormulaResult.Error("#N/A"), // if_not_found
            FormulaResult.FromNumber(0), // match_mode (exact)
            FormulaResult.FromNumber(-1), // search_mode (last to first)
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(300.0, result.NumericValue); // Last match
    }

    [Fact]
    public void XLookup_InsufficientArguments_ReturnsError()
    {
        var func = XLookupFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Apple"),
            FormulaResult.FromString("Banana"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void XLookup_ErrorInLookupValue_PropagatesError()
    {
        var func = XLookupFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#DIV/0!"), // lookup_value
            FormulaResult.FromString("Apple"),
            FormulaResult.FromString("Banana"),
            FormulaResult.FromNumber(100),
            FormulaResult.FromNumber(200),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    #endregion

    #region XMATCH Function Tests

    [Fact]
    public void XMatch_ExactMatch_ReturnsPosition()
    {
        var func = XMatchFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Banana"), // lookup_value
            FormulaResult.FromString("Apple"),
            FormulaResult.FromString("Banana"),
            FormulaResult.FromString("Cherry"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(2.0, result.NumericValue); // Second position
    }

    [Fact]
    public void XMatch_ExactMatch_NoMatch_ReturnsError()
    {
        var func = XMatchFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Orange"), // lookup_value
            FormulaResult.FromString("Apple"),
            FormulaResult.FromString("Banana"),
            FormulaResult.FromString("Cherry"),
            FormulaResult.FromNumber(0), // match_mode (exact)
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    [Fact]
    public void XMatch_ExactOrNextSmaller_ReturnsCorrectPosition()
    {
        var func = XMatchFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(25), // lookup_value
            FormulaResult.FromNumber(10),
            FormulaResult.FromNumber(20),
            FormulaResult.FromNumber(30),
            FormulaResult.FromNumber(40),
            FormulaResult.FromNumber(-1), // match_mode (exact or next smaller)
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(2.0, result.NumericValue); // Position of 20
    }

    [Fact]
    public void XMatch_ExactOrNextLarger_ReturnsCorrectPosition()
    {
        var func = XMatchFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(25), // lookup_value
            FormulaResult.FromNumber(10),
            FormulaResult.FromNumber(20),
            FormulaResult.FromNumber(30),
            FormulaResult.FromNumber(40),
            FormulaResult.FromNumber(1), // match_mode (exact or next larger)
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(3.0, result.NumericValue); // Position of 30
    }

    [Fact]
    public void XMatch_WildcardMatch_ReturnsPosition()
    {
        var func = XMatchFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("B*"), // lookup_value with wildcard
            FormulaResult.FromString("Apple"),
            FormulaResult.FromString("Banana"),
            FormulaResult.FromString("Cherry"),
            FormulaResult.FromNumber(2), // match_mode (wildcard)
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(2.0, result.NumericValue); // Matches "Banana"
    }

    [Fact]
    public void XMatch_SearchLastToFirst_ReturnsLastPosition()
    {
        var func = XMatchFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Apple"), // lookup_value
            FormulaResult.FromString("Apple"),
            FormulaResult.FromString("Banana"),
            FormulaResult.FromString("Apple"), // Another Apple
            FormulaResult.FromNumber(0), // match_mode (exact)
            FormulaResult.FromNumber(-1), // search_mode (last to first)
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(3.0, result.NumericValue); // Last match (third position)
    }

    [Fact]
    public void XMatch_BinarySearchAscending_FindsValue()
    {
        var func = XMatchFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(30), // lookup_value
            FormulaResult.FromNumber(10),
            FormulaResult.FromNumber(20),
            FormulaResult.FromNumber(30),
            FormulaResult.FromNumber(40),
            FormulaResult.FromNumber(50),
            FormulaResult.FromNumber(0), // match_mode (exact)
            FormulaResult.FromNumber(2), // search_mode (binary search ascending)
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(3.0, result.NumericValue);
    }

    [Fact]
    public void XMatch_InvalidMatchMode_ReturnsError()
    {
        var func = XMatchFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(10), // lookup_value
            FormulaResult.FromNumber(10),
            FormulaResult.FromNumber(20),
            FormulaResult.FromNumber(5), // invalid match_mode
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void XMatch_InvalidSearchMode_ReturnsError()
    {
        var func = XMatchFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(10), // lookup_value
            FormulaResult.FromNumber(10),
            FormulaResult.FromNumber(20),
            FormulaResult.FromNumber(0), // match_mode
            FormulaResult.FromNumber(0), // invalid search_mode (0 is not allowed)
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void XMatch_InsufficientArguments_ReturnsError()
    {
        var func = XMatchFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(10), // lookup_value only
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void XMatch_ErrorInLookupValue_PropagatesError()
    {
        var func = XMatchFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#REF!"), // lookup_value
            FormulaResult.FromNumber(10),
            FormulaResult.FromNumber(20),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    #endregion

    #region HYPERLINK Function Tests

    [Fact]
    public void Hyperlink_WithFriendlyName_ReturnsFriendlyName()
    {
        var func = HyperlinkFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("https://example.com"), // link_location
            FormulaResult.FromString("Click here"), // friendly_name
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("Click here", result.StringValue);
    }

    [Fact]
    public void Hyperlink_WithoutFriendlyName_ReturnsLinkLocation()
    {
        var func = HyperlinkFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("https://example.com"), // link_location
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("https://example.com", result.StringValue);
    }

    [Fact]
    public void Hyperlink_NumericFriendlyName_ReturnsNumber()
    {
        var func = HyperlinkFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("https://example.com"),
            FormulaResult.FromNumber(42), // numeric friendly_name
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(42.0, result.NumericValue);
    }

    [Fact]
    public void Hyperlink_ErrorInLinkLocation_PropagatesError()
    {
        var func = HyperlinkFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#NAME?"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NAME?", result.ErrorValue);
    }

    [Fact]
    public void Hyperlink_NoArguments_ReturnsError()
    {
        var func = HyperlinkFunction.Instance;
        var args = System.Array.Empty<FormulaResult>();

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region SHEET Function Tests

    [Fact]
    public void Sheet_NoArguments_ReturnsDefaultSheetNumber()
    {
        var func = SheetFunction.Instance;
        var args = System.Array.Empty<FormulaResult>();

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue); // Default sheet number
    }

    [Fact]
    public void Sheet_WithReference_ReturnsSheetNumber()
    {
        var func = SheetFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Sheet1!A1"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue); // Default implementation
    }

    [Fact]
    public void Sheet_ErrorInReference_PropagatesError()
    {
        var func = SheetFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#REF!"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    #endregion

    #region SHEETS Function Tests

    [Fact]
    public void Sheets_NoArguments_ReturnsDefaultCount()
    {
        var func = SheetsFunction.Instance;
        var args = System.Array.Empty<FormulaResult>();

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue); // Default sheet count
    }

    [Fact]
    public void Sheets_WithReference_ReturnsSheetCount()
    {
        var func = SheetsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Sheet1:Sheet3"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue); // Default implementation
    }

    [Fact]
    public void Sheets_ErrorInReference_PropagatesError()
    {
        var func = SheetsFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#NAME?"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NAME?", result.ErrorValue);
    }

    #endregion

    #region ISFORMULA Function Tests

    [Fact]
    public void IsFormula_ReturnsFalse_WhenNoFormula()
    {
        var func = IsFormulaFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("A1"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Boolean, result.Type);
        Assert.False(result.BoolValue); // Default implementation
    }

    [Fact]
    public void IsFormula_ErrorInReference_PropagatesError()
    {
        var func = IsFormulaFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#REF!"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    [Fact]
    public void IsFormula_NoArguments_ReturnsError()
    {
        var func = IsFormulaFunction.Instance;
        var args = System.Array.Empty<FormulaResult>();

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region FORMULATEXT Function Tests

    [Fact]
    public void FormulaText_ReturnsNotAvailable()
    {
        var func = FormulaTextFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("A1"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue); // Not fully implemented
    }

    [Fact]
    public void FormulaText_ErrorInReference_PropagatesError()
    {
        var func = FormulaTextFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#REF!"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    [Fact]
    public void FormulaText_NoArguments_ReturnsError()
    {
        var func = FormulaTextFunction.Instance;
        var args = System.Array.Empty<FormulaResult>();

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region GETPIVOTDATA Function Tests

    [Fact]
    public void GetPivotData_ReturnsRefError()
    {
        var func = GetPivotDataFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Sales"),
            FormulaResult.FromString("A1"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue); // Not fully implemented
    }

    [Fact]
    public void GetPivotData_InsufficientArguments_ReturnsError()
    {
        var func = GetPivotDataFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Sales"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void GetPivotData_ErrorInDataField_PropagatesError()
    {
        var func = GetPivotDataFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#NAME?"),
            FormulaResult.FromString("A1"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NAME?", result.ErrorValue);
    }

    #endregion
}
