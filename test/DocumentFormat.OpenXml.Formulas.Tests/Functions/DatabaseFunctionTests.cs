// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for database aggregate functions.
/// Phase 0: Tests simplified implementation with single values.
/// Future: Add tests for full range support with database headers and criteria ranges.
/// </summary>
public class DatabaseFunctionTests
{
    // DSUM Tests
    [Fact]
    public void DSum_NumericValueMatchesCriteria_ReturnsSum()
    {
        var func = DSumFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(35), // database value
            CellValue.FromString("Age"), // field (ignored in Phase 0)
            CellValue.FromString(">30"), // criteria
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(35.0, result.NumericValue);
    }

    [Fact]
    public void DSum_NumericValueDoesNotMatchCriteria_ReturnsZero()
    {
        var func = DSumFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(25), // database value
            CellValue.FromString("Age"), // field (ignored in Phase 0)
            CellValue.FromString(">30"), // criteria
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void DSum_ErrorValue_PropagatesError()
    {
        var func = DSumFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#DIV/0!"),
            CellValue.FromString("Age"),
            CellValue.FromString(">30"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void DSum_IncorrectArgumentCount_ReturnsError()
    {
        var func = DSumFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(35),
            CellValue.FromString("Age"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    // DCOUNT Tests
    [Fact]
    public void DCount_NumericValueMatchesCriteria_ReturnsOne()
    {
        var func = DCountFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(35), // database value
            CellValue.FromString("Age"), // field (ignored in Phase 0)
            CellValue.FromString(">30"), // criteria
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void DCount_NumericValueDoesNotMatchCriteria_ReturnsZero()
    {
        var func = DCountFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(25), // database value
            CellValue.FromString("Age"), // field (ignored in Phase 0)
            CellValue.FromString(">30"), // criteria
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void DCount_TextValueMatchesCriteria_ReturnsZero()
    {
        var func = DCountFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("John"), // database value
            CellValue.FromString("Name"), // field (ignored in Phase 0)
            CellValue.FromString("John"), // criteria
        };

        var result = func.Execute(null!, args);

        // DCOUNT only counts numeric values
        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void DCount_ErrorValue_PropagatesError()
    {
        var func = DCountFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#VALUE!"),
            CellValue.FromString("Age"),
            CellValue.FromString(">30"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    // DCOUNTA Tests
    [Fact]
    public void DCountA_NumericValueMatchesCriteria_ReturnsOne()
    {
        var func = DCountAFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(35), // database value
            CellValue.FromString("Age"), // field (ignored in Phase 0)
            CellValue.FromString(">30"), // criteria
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void DCountA_TextValueMatchesCriteria_ReturnsOne()
    {
        var func = DCountAFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("John"), // database value
            CellValue.FromString("Name"), // field (ignored in Phase 0)
            CellValue.FromString("John"), // criteria
        };

        var result = func.Execute(null!, args);

        // DCOUNTA counts non-empty values including text
        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void DCountA_EmptyValueMatchesCriteria_ReturnsZero()
    {
        var func = DCountAFunction.Instance;
        var args = new[]
        {
            CellValue.Empty, // database value
            CellValue.FromString("Name"), // field (ignored in Phase 0)
            CellValue.FromString("John"), // criteria - won't match empty
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void DCountA_ErrorValue_PropagatesError()
    {
        var func = DCountAFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#REF!"),
            CellValue.FromString("Age"),
            CellValue.FromString(">30"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    // DAVERAGE Tests
    [Fact]
    public void DAverage_NumericValueMatchesCriteria_ReturnsValue()
    {
        var func = DAverageFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(50000), // database value
            CellValue.FromString("Salary"), // field (ignored in Phase 0)
            CellValue.FromString(">30000"), // criteria
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(50000.0, result.NumericValue);
    }

    [Fact]
    public void DAverage_NumericValueDoesNotMatchCriteria_ReturnsDivZeroError()
    {
        var func = DAverageFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(25000), // database value
            CellValue.FromString("Salary"), // field (ignored in Phase 0)
            CellValue.FromString(">30000"), // criteria
        };

        var result = func.Execute(null!, args);

        // No values match criteria, should return #DIV/0!
        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void DAverage_ErrorValue_PropagatesError()
    {
        var func = DAverageFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#N/A"),
            CellValue.FromString("Salary"),
            CellValue.FromString(">30000"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    // DMAX Tests
    [Fact]
    public void DMax_NumericValueMatchesCriteria_ReturnsValue()
    {
        var func = DMaxFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(45), // database value
            CellValue.FromString("Age"), // field (ignored in Phase 0)
            CellValue.FromString(">30"), // criteria
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(45.0, result.NumericValue);
    }

    [Fact]
    public void DMax_NumericValueDoesNotMatchCriteria_ReturnsZero()
    {
        var func = DMaxFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(20), // database value
            CellValue.FromString("Age"), // field (ignored in Phase 0)
            CellValue.FromString(">30"), // criteria
        };

        var result = func.Execute(null!, args);

        // No values match criteria, should return 0
        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void DMax_ErrorValue_PropagatesError()
    {
        var func = DMaxFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#NUM!"),
            CellValue.FromString("Age"),
            CellValue.FromString(">30"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    // DMIN Tests
    [Fact]
    public void DMin_NumericValueMatchesCriteria_ReturnsValue()
    {
        var func = DMinFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(35), // database value
            CellValue.FromString("Age"), // field (ignored in Phase 0)
            CellValue.FromString(">30"), // criteria
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(35.0, result.NumericValue);
    }

    [Fact]
    public void DMin_NumericValueDoesNotMatchCriteria_ReturnsZero()
    {
        var func = DMinFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(20), // database value
            CellValue.FromString("Age"), // field (ignored in Phase 0)
            CellValue.FromString(">30"), // criteria
        };

        var result = func.Execute(null!, args);

        // No values match criteria, should return 0
        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void DMin_ErrorValue_PropagatesError()
    {
        var func = DMinFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#NAME?"),
            CellValue.FromString("Age"),
            CellValue.FromString(">30"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NAME?", result.ErrorValue);
    }

    // Criteria matching tests for various operators
    [Fact]
    public void DSum_GreaterThanOrEqualCriteria_WorksCorrectly()
    {
        var func = DSumFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(30), // database value
            CellValue.FromString("Age"),
            CellValue.FromString(">=30"), // criteria
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(30.0, result.NumericValue);
    }

    [Fact]
    public void DSum_LessThanOrEqualCriteria_WorksCorrectly()
    {
        var func = DSumFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(25), // database value
            CellValue.FromString("Age"),
            CellValue.FromString("<=30"), // criteria
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(25.0, result.NumericValue);
    }

    [Fact]
    public void DSum_NotEqualCriteria_WorksCorrectly()
    {
        var func = DSumFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(25), // database value
            CellValue.FromString("Age"),
            CellValue.FromString("<>30"), // criteria
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(25.0, result.NumericValue);
    }

    [Fact]
    public void DSum_EqualCriteria_WorksCorrectly()
    {
        var func = DSumFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(30), // database value
            CellValue.FromString("Age"),
            CellValue.FromString("=30"), // criteria
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(30.0, result.NumericValue);
    }

    [Fact]
    public void DSum_NumericCriteriaDirectComparison_WorksCorrectly()
    {
        var func = DSumFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(30), // database value
            CellValue.FromString("Age"),
            CellValue.FromNumber(30), // criteria as number
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(30.0, result.NumericValue);
    }

    [Fact]
    public void DCountA_TextCriteriaDirectComparison_WorksCorrectly()
    {
        var func = DCountAFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("Engineer"), // database value
            CellValue.FromString("Title"),
            CellValue.FromString("Engineer"), // criteria as text
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void DCountA_TextCriteriaDoesNotMatch_ReturnsZero()
    {
        var func = DCountAFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("Engineer"), // database value
            CellValue.FromString("Title"),
            CellValue.FromString("Manager"), // criteria as text
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.0, result.NumericValue);
    }

    // DGET Tests
    [Fact]
    public void DGet_NumericValueMatchesCriteria_ReturnsValue()
    {
        var func = DGetFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(42), // database value
            CellValue.FromString("ID"), // field (ignored in Phase 0)
            CellValue.FromString("=42"), // criteria
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(42.0, result.NumericValue);
    }

    [Fact]
    public void DGet_TextValueMatchesCriteria_ReturnsValue()
    {
        var func = DGetFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("Alice"), // database value
            CellValue.FromString("Name"), // field (ignored in Phase 0)
            CellValue.FromString("Alice"), // criteria
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("Alice", result.StringValue);
    }

    [Fact]
    public void DGet_NoMatchingValue_ReturnsError()
    {
        var func = DGetFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(25), // database value
            CellValue.FromString("Age"), // field (ignored in Phase 0)
            CellValue.FromString(">30"), // criteria
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void DGet_ErrorValue_PropagatesError()
    {
        var func = DGetFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#REF!"),
            CellValue.FromString("Field"),
            CellValue.FromString("criteria"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    [Fact]
    public void DGet_IncorrectArgumentCount_ReturnsError()
    {
        var func = DGetFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(42),
            CellValue.FromString("ID"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    // DPRODUCT Tests
    [Fact]
    public void DProduct_NumericValueMatchesCriteria_ReturnsValue()
    {
        var func = DProductFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5), // database value
            CellValue.FromString("Quantity"), // field (ignored in Phase 0)
            CellValue.FromString(">0"), // criteria
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void DProduct_NumericValueDoesNotMatchCriteria_ReturnsZero()
    {
        var func = DProductFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5), // database value
            CellValue.FromString("Quantity"), // field (ignored in Phase 0)
            CellValue.FromString(">10"), // criteria
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void DProduct_ErrorValue_PropagatesError()
    {
        var func = DProductFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#NUM!"),
            CellValue.FromString("Quantity"),
            CellValue.FromString(">0"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    // DSTDEV Tests
    [Fact]
    public void DStDev_InsufficientValues_ReturnsDivZeroError()
    {
        var func = DStDevFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(50), // database value (only 1 value)
            CellValue.FromString("Score"), // field (ignored in Phase 0)
            CellValue.FromString(">30"), // criteria
        };

        var result = func.Execute(null!, args);

        // DSTDEV requires at least 2 values
        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void DStDev_NoMatchingValues_ReturnsDivZeroError()
    {
        var func = DStDevFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(20), // database value
            CellValue.FromString("Score"), // field (ignored in Phase 0)
            CellValue.FromString(">30"), // criteria
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void DStDev_ErrorValue_PropagatesError()
    {
        var func = DStDevFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#VALUE!"),
            CellValue.FromString("Score"),
            CellValue.FromString(">30"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    // DSTDEVP Tests
    [Fact]
    public void DStDevP_SingleValue_ReturnsZero()
    {
        var func = DStDevPFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(50), // database value (only 1 value)
            CellValue.FromString("Score"), // field (ignored in Phase 0)
            CellValue.FromString(">30"), // criteria
        };

        var result = func.Execute(null!, args);

        // Population standard deviation of single value is 0
        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void DStDevP_NoMatchingValues_ReturnsDivZeroError()
    {
        var func = DStDevPFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(20), // database value
            CellValue.FromString("Score"), // field (ignored in Phase 0)
            CellValue.FromString(">30"), // criteria
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void DStDevP_ErrorValue_PropagatesError()
    {
        var func = DStDevPFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#N/A"),
            CellValue.FromString("Score"),
            CellValue.FromString(">30"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    // DVAR Tests
    [Fact]
    public void DVar_InsufficientValues_ReturnsDivZeroError()
    {
        var func = DVarFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(50), // database value (only 1 value)
            CellValue.FromString("Value"), // field (ignored in Phase 0)
            CellValue.FromString(">30"), // criteria
        };

        var result = func.Execute(null!, args);

        // DVAR requires at least 2 values
        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void DVar_NoMatchingValues_ReturnsDivZeroError()
    {
        var func = DVarFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(20), // database value
            CellValue.FromString("Value"), // field (ignored in Phase 0)
            CellValue.FromString(">30"), // criteria
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void DVar_ErrorValue_PropagatesError()
    {
        var func = DVarFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#NAME?"),
            CellValue.FromString("Value"),
            CellValue.FromString(">30"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NAME?", result.ErrorValue);
    }

    // DVARP Tests
    [Fact]
    public void DVarP_SingleValue_ReturnsZero()
    {
        var func = DVarPFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(50), // database value (only 1 value)
            CellValue.FromString("Value"), // field (ignored in Phase 0)
            CellValue.FromString(">30"), // criteria
        };

        var result = func.Execute(null!, args);

        // Population variance of single value is 0
        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void DVarP_NoMatchingValues_ReturnsDivZeroError()
    {
        var func = DVarPFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(20), // database value
            CellValue.FromString("Value"), // field (ignored in Phase 0)
            CellValue.FromString(">30"), // criteria
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void DVarP_ErrorValue_PropagatesError()
    {
        var func = DVarPFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#REF!"),
            CellValue.FromString("Value"),
            CellValue.FromString(">30"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    // Additional integration tests for criteria matching
    [Fact]
    public void DProduct_MultipleMatchingCriteria_WorksCorrectly()
    {
        var func = DProductFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(7), // database value
            CellValue.FromString("Multiplier"),
            CellValue.FromString(">=5"), // criteria
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(7.0, result.NumericValue);
    }

    [Fact]
    public void DGet_BooleanValueMatchesCriteria_ReturnsValue()
    {
        var func = DGetFunction.Instance;
        var args = new[]
        {
            CellValue.FromBoolean(true), // database value
            CellValue.FromString("IsActive"),
            CellValue.FromBoolean(true), // criteria
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }
}
