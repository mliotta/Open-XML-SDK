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

        var result = func.Execute(null!, System.Array.Empty<FormulaResult>());

        Assert.Equal(FormulaResultType.Number, result.Type);
        var value = result.NumericValue;
        Assert.True(value >= 0.0 && value < 1.0);
    }

    [Fact]
    public void RandArray_WithDimensions_ReturnsValue()
    {
        var func = RandArrayFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            FormulaResult.FromNumber(3),
            FormulaResult.FromNumber(4),
        });

        // Currently returns single value (full array support requires engine changes)
        Assert.Equal(FormulaResultType.Number, result.Type);
        var value = result.NumericValue;
        Assert.True(value >= 0.0 && value < 1.0);
    }

    [Fact]
    public void RandArray_WithMinMax_ReturnsValueInRange()
    {
        var func = RandArrayFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            FormulaResult.FromNumber(2),
            FormulaResult.FromNumber(2),
            FormulaResult.FromNumber(10),
            FormulaResult.FromNumber(20),
        });

        Assert.Equal(FormulaResultType.Number, result.Type);
        var value = result.NumericValue;
        Assert.True(value >= 10.0 && value < 20.0);
    }

    [Fact]
    public void RandArray_WholeNumbers_ReturnsInteger()
    {
        var func = RandArrayFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            FormulaResult.FromNumber(3),
            FormulaResult.FromNumber(3),
            FormulaResult.FromNumber(1),
            FormulaResult.FromNumber(10),
            FormulaResult.FromBool(true),
        });

        Assert.Equal(FormulaResultType.Number, result.Type);
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
            FormulaResult.FromNumber(-1),
            FormulaResult.FromNumber(1),
        });

        Assert.True(result1.IsError);
        Assert.Equal("#VALUE!", result1.ErrorValue);

        // Zero columns
        var result2 = func.Execute(null!, new[]
        {
            FormulaResult.FromNumber(1),
            FormulaResult.FromNumber(0),
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
            FormulaResult.FromNumber(1),
            FormulaResult.FromNumber(1),
            FormulaResult.FromNumber(20),
            FormulaResult.FromNumber(10),
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
            FormulaResult.FromNumber(1),
            FormulaResult.FromNumber(1),
            FormulaResult.FromNumber(10),
            FormulaResult.FromNumber(10),
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
            FormulaResult.FromString("text"),
            FormulaResult.FromNumber(1),
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
            FormulaResult.Error("#DIV/0!"),
            FormulaResult.FromNumber(1),
        });

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }
}
