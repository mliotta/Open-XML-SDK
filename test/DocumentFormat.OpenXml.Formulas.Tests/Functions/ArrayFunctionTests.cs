// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for modern Excel array functions (TRANSPOSE, SORT, FILTER, UNIQUE, SEQUENCE).
/// </summary>
public class ArrayFunctionTests
{
    #region TRANSPOSE Function Tests

    [Fact]
    public void Transpose_SingleCell_ReturnsSameValue()
    {
        var func = TransposeFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(42),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(42.0, result.NumericValue);
    }

    [Fact]
    public void Transpose_2x2Array_ReturnsFirstElement()
    {
        var func = TransposeFunction.Instance;
        // Original: [10, 20]
        //           [30, 40]
        // Transposed: [10, 30]
        //             [20, 40]
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
            CellValue.FromNumber(40),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(10.0, result.NumericValue); // First element unchanged
    }

    [Fact]
    public void Transpose_TextValues_Works()
    {
        var func = TransposeFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("A"),
            CellValue.FromString("B"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("A", result.StringValue);
    }

    [Fact]
    public void Transpose_ErrorInArray_PropagatesError()
    {
        var func = TransposeFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.Error("#DIV/0!"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void Transpose_EmptyArray_ReturnsError()
    {
        var func = TransposeFunction.Instance;
        var args = new CellValue[0];

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region SORT Function Tests

    [Fact]
    public void Sort_SingleCell_ReturnsSameValue()
    {
        var func = SortFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(42),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(42.0, result.NumericValue);
    }

    [Fact]
    public void Sort_NumbersAscending_ReturnsFirstElement()
    {
        var func = SortFunction.Instance;
        // Array: [30, 10, 20]
        // Sorted: [10, 20, 30]
        var args = new[]
        {
            CellValue.FromNumber(30),
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(10.0, result.NumericValue); // First element of sorted array
    }

    [Fact]
    public void Sort_NumbersDescending_ReturnsFirstElement()
    {
        var func = SortFunction.Instance;
        // Array: [10, 20, 30]
        // Sort order: -1 (descending)
        // Sorted: [30, 20, 10]
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(30),
            CellValue.FromNumber(-1), // sort_order
        };

        var result = func.Execute(null!, args);

        Assert.Equal(30.0, result.NumericValue);
    }

    [Fact]
    public void Sort_TextValues_SortsAlphabetically()
    {
        var func = SortFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("Zebra"),
            CellValue.FromString("Apple"),
            CellValue.FromString("Banana"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("Apple", result.StringValue);
    }

    [Fact]
    public void Sort_2x2ArrayByFirstColumn_Works()
    {
        var func = SortFunction.Instance;
        // Array: [30, "X"]
        //        [10, "Y"]
        // Sort by column 1 (default)
        // Sorted: [10, "Y"]
        //         [30, "X"]
        var args = new[]
        {
            CellValue.FromNumber(30),
            CellValue.FromString("X"),
            CellValue.FromNumber(10),
            CellValue.FromString("Y"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void Sort_ErrorInArray_PropagatesError()
    {
        var func = SortFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.Error("#REF!"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    [Fact]
    public void Sort_EmptyArray_ReturnsError()
    {
        var func = SortFunction.Instance;
        var args = new CellValue[0];

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region FILTER Function Tests

    [Fact]
    public void Filter_SingleValueMatches_ReturnsValue()
    {
        var func = FilterFunction.Instance;
        // Array: [42]
        // Include: [TRUE]
        var args = new[]
        {
            CellValue.FromNumber(42),
            CellValue.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(42.0, result.NumericValue);
    }

    [Fact]
    public void Filter_SingleValueDoesNotMatch_ReturnsError()
    {
        var func = FilterFunction.Instance;
        // Array: [42]
        // Include: [FALSE]
        var args = new[]
        {
            CellValue.FromNumber(42),
            CellValue.FromBool(false),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#CALC!", result.ErrorValue);
    }

    [Fact]
    public void Filter_WithIfEmpty_ReturnsCustomValue()
    {
        var func = FilterFunction.Instance;
        // Array: [42]
        // Include: [FALSE]
        // If_empty: "No matches"
        var args = new[]
        {
            CellValue.FromNumber(42),
            CellValue.FromBool(false),
            CellValue.FromString("No matches"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("No matches", result.StringValue);
    }

    [Fact]
    public void Filter_MultipleRows_FiltersCorrectly()
    {
        var func = FilterFunction.Instance;
        // Array: [10, "A"]
        //        [20, "B"]
        //        [30, "C"]
        // Include: [TRUE, FALSE, TRUE]
        // Result: [10, "A"]
        //         [30, "C"]
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromString("A"),
            CellValue.FromNumber(20),
            CellValue.FromString("B"),
            CellValue.FromNumber(30),
            CellValue.FromString("C"),
            CellValue.FromBool(true),
            CellValue.FromBool(false),
            CellValue.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(10.0, result.NumericValue); // First element of filtered array
    }

    [Fact]
    public void Filter_NumericIncludeValues_TreatsNonZeroAsTrue()
    {
        var func = FilterFunction.Instance;
        // Array: [42]
        // Include: [1] (treated as TRUE)
        var args = new[]
        {
            CellValue.FromNumber(42),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(42.0, result.NumericValue);
    }

    [Fact]
    public void Filter_ErrorInArray_PropagatesError()
    {
        var func = FilterFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#N/A"),
            CellValue.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    [Fact]
    public void Filter_InsufficientArguments_ReturnsError()
    {
        var func = FilterFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(42),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region UNIQUE Function Tests

    [Fact]
    public void Unique_SingleValue_ReturnsSameValue()
    {
        var func = UniqueFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(42),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(42.0, result.NumericValue);
    }

    [Fact]
    public void Unique_DuplicateNumbers_ReturnsFirstUnique()
    {
        var func = UniqueFunction.Instance;
        // Array: [10, 20, 10, 30]
        // Unique: [10, 20, 30]
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(10),
            CellValue.FromNumber(30),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void Unique_DuplicateStrings_ReturnsFirstUnique()
    {
        var func = UniqueFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("Apple"),
            CellValue.FromString("Banana"),
            CellValue.FromString("Apple"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal("Apple", result.StringValue);
    }

    [Fact]
    public void Unique_OccursOnce_FiltersToOnlyOnce()
    {
        var func = UniqueFunction.Instance;
        // Array: [10, 20, 10, 30]
        // occurs_once: TRUE
        // Result: [20, 30] (values that appear exactly once)
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
            CellValue.FromNumber(10),
            CellValue.FromNumber(30),
            CellValue.FromBool(true), // occurs_once
        };

        var result = func.Execute(null!, args);

        Assert.Equal(20.0, result.NumericValue); // First value that occurs once
    }

    [Fact]
    public void Unique_AllDuplicates_WithOccursOnce_ReturnsError()
    {
        var func = UniqueFunction.Instance;
        // Array: [10, 10]
        // occurs_once: TRUE
        // Result: Error (no values occur exactly once)
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromNumber(10),
            CellValue.FromBool(true), // occurs_once
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#CALC!", result.ErrorValue);
    }

    [Fact]
    public void Unique_2x2Array_UniqueRows()
    {
        var func = UniqueFunction.Instance;
        // Array: [10, "A"]
        //        [20, "B"]
        //        [10, "A"]  (duplicate of first row)
        // Unique rows: [10, "A"]
        //              [20, "B"]
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.FromString("A"),
            CellValue.FromNumber(20),
            CellValue.FromString("B"),
            CellValue.FromNumber(10),
            CellValue.FromString("A"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void Unique_ErrorInArray_PropagatesError()
    {
        var func = UniqueFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10),
            CellValue.Error("#VALUE!"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Unique_EmptyArray_ReturnsError()
    {
        var func = UniqueFunction.Instance;
        var args = new CellValue[0];

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region SEQUENCE Function Tests

    [Fact]
    public void Sequence_SingleRow_ReturnsFirstValue()
    {
        var func = SequenceFunction.Instance;
        // SEQUENCE(5) -> [1, 2, 3, 4, 5]
        var args = new[]
        {
            CellValue.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void Sequence_CustomStart_ReturnsStartValue()
    {
        var func = SequenceFunction.Instance;
        // SEQUENCE(3, 1, 10) -> [10, 11, 12]
        var args = new[]
        {
            CellValue.FromNumber(3),
            CellValue.FromNumber(1), // columns
            CellValue.FromNumber(10), // start
        };

        var result = func.Execute(null!, args);

        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void Sequence_CustomStep_Works()
    {
        var func = SequenceFunction.Instance;
        // SEQUENCE(3, 1, 1, 2) -> [1, 3, 5]
        var args = new[]
        {
            CellValue.FromNumber(3),
            CellValue.FromNumber(1), // columns
            CellValue.FromNumber(1), // start
            CellValue.FromNumber(2), // step
        };

        var result = func.Execute(null!, args);

        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void Sequence_NegativeStep_Works()
    {
        var func = SequenceFunction.Instance;
        // SEQUENCE(3, 1, 10, -1) -> [10, 9, 8]
        var args = new[]
        {
            CellValue.FromNumber(3),
            CellValue.FromNumber(1), // columns
            CellValue.FromNumber(10), // start
            CellValue.FromNumber(-1), // step
        };

        var result = func.Execute(null!, args);

        Assert.Equal(10.0, result.NumericValue);
    }

    [Fact]
    public void Sequence_MultiColumn_ReturnsFirstValue()
    {
        var func = SequenceFunction.Instance;
        // SEQUENCE(2, 2) -> [1, 2]
        //                   [3, 4]
        var args = new[]
        {
            CellValue.FromNumber(2),
            CellValue.FromNumber(2), // columns
        };

        var result = func.Execute(null!, args);

        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void Sequence_ZeroRows_ReturnsError()
    {
        var func = SequenceFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Sequence_NegativeRows_ReturnsError()
    {
        var func = SequenceFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-5),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Sequence_InvalidType_ReturnsError()
    {
        var func = SequenceFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("invalid"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Sequence_ErrorInArgument_PropagatesError()
    {
        var func = SequenceFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#DIV/0!"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void Sequence_NoArguments_ReturnsError()
    {
        var func = SequenceFunction.Instance;
        var args = new CellValue[0];

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion
}
