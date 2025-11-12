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
}
