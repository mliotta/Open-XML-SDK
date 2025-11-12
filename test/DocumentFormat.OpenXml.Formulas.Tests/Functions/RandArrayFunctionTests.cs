// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;
using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for RANDARRAY function.
/// </summary>
public class RandArrayFunctionTests
{
    [Fact]
    public void RandArray_NoArguments_ReturnsSingleValue()
    {
        var func = RandArrayFunction.Instance;

        var result = func.Execute(null!, System.Array.Empty<CellValue>());

        Assert.Equal(CellValueType.Number, result.Type);
        var value = result.NumericValue;
        Assert.True(value >= 0.0 && value < 1.0);
    }

    [Fact]
    public void RandArray_WithDimensions_ReturnsValue()
    {
        var func = RandArrayFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(3),
            CellValue.FromNumber(4),
        });

        // Currently returns single value (full array support requires engine changes)
        Assert.Equal(CellValueType.Number, result.Type);
        var value = result.NumericValue;
        Assert.True(value >= 0.0 && value < 1.0);
    }

    [Fact]
    public void RandArray_WithMinMax_ReturnsValueInRange()
    {
        var func = RandArrayFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(2),
            CellValue.FromNumber(2),
            CellValue.FromNumber(10),
            CellValue.FromNumber(20),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        var value = result.NumericValue;
        Assert.True(value >= 10.0 && value < 20.0);
    }

    [Fact]
    public void RandArray_WholeNumbers_ReturnsInteger()
    {
        var func = RandArrayFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(3),
            CellValue.FromNumber(3),
            CellValue.FromNumber(1),
            CellValue.FromNumber(10),
            CellValue.FromBool(true),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        var value = result.NumericValue;
        Assert.True(value >= 1.0 && value < 10.0);
        Assert.Equal(System.Math.Floor(value), value);
    }

    [Fact]
    public void RandArray_InvalidDimensions_ReturnsError()
    {
        var func = RandArrayFunction.Instance;

        // Negative rows
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-1),
            CellValue.FromNumber(1),
        });

        Assert.True(result1.IsError);
        Assert.Equal("#VALUE!", result1.ErrorValue);

        // Zero columns
        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(0),
        });

        Assert.True(result2.IsError);
        Assert.Equal("#VALUE!", result2.ErrorValue);
    }

    [Fact]
    public void RandArray_MinGreaterThanMax_ReturnsError()
    {
        var func = RandArrayFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(1),
            CellValue.FromNumber(20),
            CellValue.FromNumber(10),
        });

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void RandArray_MinEqualsMax_ReturnsError()
    {
        var func = RandArrayFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(1),
            CellValue.FromNumber(10),
            CellValue.FromNumber(10),
        });

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void RandArray_InvalidArgumentTypes_ReturnsError()
    {
        var func = RandArrayFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromString("text"),
            CellValue.FromNumber(1),
        });

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void RandArray_ErrorPropagation_ReturnsError()
    {
        var func = RandArrayFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(1),
        });

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }
}
