// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for INDEX and MATCH lookup functions.
/// </summary>
public class LookupFunctionTests
{
    #region INDEX Function Tests

    [Fact]
    public void Index_1DArray_ReturnsCorrectValue()
    {
        var func = IndexFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
            CellValue.FromNumber(40),
            CellValue.FromNumber(50),
            CellValue.FromNumber(3), // row_num
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(30.0, result.NumericValue);
    }

    [Fact]
    public void Index_2DArray_WithRowAndColumn_ReturnsCorrectValue()
    {
        var func = IndexFunction.Instance;
        // 3x2 array: [10, 20]
        //            [30, 40]
        //            [50, 60]
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
            CellValue.FromNumber(40),
            CellValue.FromNumber(50),
            CellValue.FromNumber(60),
            CellValue.FromNumber(2), // col_num
            CellValue.FromNumber(2), // row_num
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(40.0, result.NumericValue);
    }

    [Fact]
    public void Index_2DArray_FirstRow_FirstColumn()
    {
        var func = IndexFunction.Instance;
        // 2x2 array: [10, 20]
        //            [30, 40]
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
            CellValue.FromNumber(40),
            CellValue.FromNumber(1), // col_num
            CellValue.FromNumber(1), // row_num
        };

        var result = func.Execute(null!, args);

        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void Index_2DArray_LastRow_LastColumn()
    {
        var func = IndexFunction.Instance;
        // 2x2 array: [10, 20]
        //            [30, 40]
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
            CellValue.FromNumber(40),
            CellValue.FromNumber(2), // col_num
            CellValue.FromNumber(2), // row_num
        };

        var result = func.Execute(null!, args);

        Assert.Equal(40.0, result.NumericValue);
    }

    [Fact]
    public void Index_SingleCell_ValidIndices_ReturnsCell()
    {
        var func = IndexFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(42),
            CellValue.FromNumber(1), // row_num
        };

        var result = func.Execute(null!, args);

        Assert.Equal(42.0, result.NumericValue);
    }

    [Fact]
    public void Index_SingleCell_InvalidIndices_ReturnsError()
    {
        var func = IndexFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(42),
            CellValue.FromNumber(2), // row_num (out of bounds)
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    [Fact]
    public void Index_OutOfBoundsRow_ReturnsError()
    {
        var func = IndexFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
            CellValue.FromNumber(5), // row_num (out of bounds)
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    [Fact]
    public void Index_OutOfBoundsColumn_ReturnsError()
    {
        var func = IndexFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
            CellValue.FromNumber(40),
            CellValue.FromNumber(5), // col_num (out of bounds)
            CellValue.FromNumber(1), // row_num
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    [Fact]
    public void Index_InvalidRowNum_ReturnsError()
    {
        var func = IndexFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromString("invalid"), // row_num
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Index_NegativeRowNum_ReturnsError()
    {
        var func = IndexFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(-1), // row_num
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Index_ErrorInArray_PropagatesError()
    {
        var func = IndexFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(30),
            CellValue.FromNumber(1), // row_num
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void Index_TextValues_ReturnsText()
    {
        var func = IndexFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("Apple"),
            CellValue.FromString("Banana"),
            CellValue.FromString("Cherry"),
            CellValue.FromNumber(2), // row_num
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("Banana", result.StringValue);
    }

    [Fact]
    public void Index_InsufficientArguments_ReturnsError()
    {
        var func = IndexFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region MATCH Function Tests

    [Fact]
    public void Match_ExactMatch_ReturnsPosition()
    {
        var func = MatchFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("Apple"), // lookup_value
            CellValue.FromString("Banana"),
            CellValue.FromString("Apple"),
            CellValue.FromString("Cherry"),
            CellValue.FromNumber(0), // match_type (exact)
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(2.0, result.NumericValue); // Second position
    }

    [Fact]
    public void Match_ExactMatch_NoMatch_ReturnsError()
    {
        var func = MatchFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("Orange"), // lookup_value
            CellValue.FromString("Apple"),
            CellValue.FromString("Banana"),
            CellValue.FromString("Cherry"),
            CellValue.FromNumber(0), // match_type (exact)
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    [Fact]
    public void Match_ApproximateMatch_Ascending_ReturnsLargestLessOrEqual()
    {
        var func = MatchFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(25), // lookup_value
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
            CellValue.FromNumber(40),
            CellValue.FromNumber(50),
            CellValue.FromNumber(1), // match_type (largest <= lookup_value)
        };

        var result = func.Execute(null!, args);

        Assert.Equal(2.0, result.NumericValue); // Position of 20
    }

    [Fact]
    public void Match_ApproximateMatch_Descending_ReturnsSmallestGreaterOrEqual()
    {
        var func = MatchFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(25), // lookup_value
            CellValue.FromNumber(50),
            CellValue.FromNumber(40),
            CellValue.FromNumber(30),
            CellValue.FromNumber(20),
            CellValue.FromNumber(10),
            CellValue.FromNumber(-1), // match_type (smallest >= lookup_value)
        };

        var result = func.Execute(null!, args);

        Assert.Equal(3.0, result.NumericValue); // Position of 30
    }

    [Fact]
    public void Match_DefaultMatchType_UsesApproximateAscending()
    {
        var func = MatchFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(25), // lookup_value
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
            CellValue.FromNumber(40),
            // No match_type specified (defaults to 1)
        };

        var result = func.Execute(null!, args);

        Assert.Equal(2.0, result.NumericValue); // Position of 20
    }

    [Fact]
    public void Match_ExactMatchWithNumbers_ReturnsCorrectPosition()
    {
        var func = MatchFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(30), // lookup_value
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
            CellValue.FromNumber(40),
            CellValue.FromNumber(0), // match_type (exact)
        };

        var result = func.Execute(null!, args);

        Assert.Equal(3.0, result.NumericValue);
    }

    [Fact]
    public void Match_FirstPosition_ReturnsOne()
    {
        var func = MatchFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("Apple"), // lookup_value
            CellValue.FromString("Apple"),
            CellValue.FromString("Banana"),
            CellValue.FromString("Cherry"),
            CellValue.FromNumber(0), // match_type (exact)
        };

        var result = func.Execute(null!, args);

        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void Match_LastPosition_ReturnsCorrectIndex()
    {
        var func = MatchFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("Cherry"), // lookup_value
            CellValue.FromString("Apple"),
            CellValue.FromString("Banana"),
            CellValue.FromString("Cherry"),
            CellValue.FromNumber(0), // match_type (exact)
        };

        var result = func.Execute(null!, args);

        Assert.Equal(3.0, result.NumericValue);
    }

    [Fact]
    public void Match_ErrorInLookupValue_PropagatesError()
    {
        var func = MatchFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#DIV/0!"), // lookup_value
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void Match_ErrorInArray_PropagatesError()
    {
        var func = MatchFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10), // lookup_value
            CellValue.FromNumber(10),
            CellValue.Error("#REF!"),
            CellValue.FromNumber(30),
            CellValue.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    [Fact]
    public void Match_InvalidMatchType_ReturnsError()
    {
        var func = MatchFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10), // lookup_value
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(2), // invalid match_type
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Match_EmptyArray_ReturnsError()
    {
        var func = MatchFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10), // lookup_value
            CellValue.FromNumber(0), // match_type
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    [Fact]
    public void Match_InsufficientArguments_ReturnsError()
    {
        var func = MatchFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10), // lookup_value only
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Match_CaseInsensitiveText_Matches()
    {
        var func = MatchFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("APPLE"), // lookup_value
            CellValue.FromString("apple"),
            CellValue.FromString("banana"),
            CellValue.FromNumber(0), // match_type (exact)
        };

        var result = func.Execute(null!, args);

        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void Match_ApproximateMatch_ExactValueExists_ReturnsExactPosition()
    {
        var func = MatchFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(30), // lookup_value
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
            CellValue.FromNumber(40),
            CellValue.FromNumber(1), // match_type (approximate)
        };

        var result = func.Execute(null!, args);

        Assert.Equal(3.0, result.NumericValue);
    }

    [Fact]
    public void Match_ApproximateMatch_NoValueLessOrEqual_ReturnsError()
    {
        var func = MatchFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5), // lookup_value (smaller than all)
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
            CellValue.FromNumber(1), // match_type
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    #endregion

    #region INDEX-MATCH Combination Tests

    [Fact]
    public void IndexMatch_CombinedPattern_ReturnsCorrectValue()
    {
        // Simulates INDEX(A1:A3, MATCH("Banana", A1:A3, 0))
        var matchFunc = MatchFunction.Instance;
        var indexFunc = IndexFunction.Instance;

        // First, execute MATCH
        var matchArgs = new[]
        {
            CellValue.FromString("Banana"), // lookup_value
            CellValue.FromString("Apple"),
            CellValue.FromString("Banana"),
            CellValue.FromString("Cherry"),
            CellValue.FromNumber(0), // match_type
        };

        var matchResult = matchFunc.Execute(null!, matchArgs);
        Assert.Equal(2.0, matchResult.NumericValue);

        // Then, use MATCH result in INDEX
        var indexArgs = new[]
        {
            CellValue.FromNumber(100),
            CellValue.FromNumber(200),
            CellValue.FromNumber(300),
            matchResult, // row_num from MATCH
        };

        var indexResult = indexFunc.Execute(null!, indexArgs);
        Assert.Equal(200.0, indexResult.NumericValue);
    }

    #endregion
}
