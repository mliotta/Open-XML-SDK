// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for information functions (ISEVEN, ISODD, ISLOGICAL, ISNONTEXT, TYPE, N).
/// </summary>
public class InformationFunctionTests
{
    #region ISEVEN Tests

    [Fact]
    public void IsEven_EvenInteger_ReturnsTrue()
    {
        var func = IsEvenFunction.Instance;
        var args = new[] { CellValue.FromNumber(4) };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void IsEven_OddInteger_ReturnsFalse()
    {
        var func = IsEvenFunction.Instance;
        var args = new[] { CellValue.FromNumber(3) };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void IsEven_Zero_ReturnsTrue()
    {
        var func = IsEvenFunction.Instance;
        var args = new[] { CellValue.FromNumber(0) };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void IsEven_NegativeEven_ReturnsTrue()
    {
        var func = IsEvenFunction.Instance;
        var args = new[] { CellValue.FromNumber(-6) };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void IsEven_NegativeOdd_ReturnsFalse()
    {
        var func = IsEvenFunction.Instance;
        var args = new[] { CellValue.FromNumber(-5) };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void IsEven_DecimalEven_ReturnsTrue()
    {
        var func = IsEvenFunction.Instance;
        var args = new[] { CellValue.FromNumber(4.7) };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void IsEven_DecimalOdd_ReturnsFalse()
    {
        var func = IsEvenFunction.Instance;
        var args = new[] { CellValue.FromNumber(3.2) };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void IsEven_Text_ReturnsFalse()
    {
        var func = IsEvenFunction.Instance;
        var args = new[] { CellValue.FromString("text") };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void IsEven_Error_ReturnsFalse()
    {
        var func = IsEvenFunction.Instance;
        var args = new[] { CellValue.Error("#DIV/0!") };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void IsEven_WrongArgumentCount_ReturnsError()
    {
        var func = IsEvenFunction.Instance;
        var args = new[] { CellValue.FromNumber(1), CellValue.FromNumber(2) };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region ISODD Tests

    [Fact]
    public void IsOdd_OddInteger_ReturnsTrue()
    {
        var func = IsOddFunction.Instance;
        var args = new[] { CellValue.FromNumber(3) };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void IsOdd_EvenInteger_ReturnsFalse()
    {
        var func = IsOddFunction.Instance;
        var args = new[] { CellValue.FromNumber(4) };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void IsOdd_Zero_ReturnsFalse()
    {
        var func = IsOddFunction.Instance;
        var args = new[] { CellValue.FromNumber(0) };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void IsOdd_NegativeOdd_ReturnsTrue()
    {
        var func = IsOddFunction.Instance;
        var args = new[] { CellValue.FromNumber(-5) };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void IsOdd_NegativeEven_ReturnsFalse()
    {
        var func = IsOddFunction.Instance;
        var args = new[] { CellValue.FromNumber(-6) };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void IsOdd_DecimalOdd_ReturnsTrue()
    {
        var func = IsOddFunction.Instance;
        var args = new[] { CellValue.FromNumber(3.9) };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void IsOdd_Text_ReturnsFalse()
    {
        var func = IsOddFunction.Instance;
        var args = new[] { CellValue.FromString("text") };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void IsOdd_WrongArgumentCount_ReturnsError()
    {
        var func = IsOddFunction.Instance;
        var args = new CellValue[] { };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region ISLOGICAL Tests

    [Fact]
    public void IsLogical_True_ReturnsTrue()
    {
        var func = IsLogicalFunction.Instance;
        var args = new[] { CellValue.FromBool(true) };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void IsLogical_False_ReturnsTrue()
    {
        var func = IsLogicalFunction.Instance;
        var args = new[] { CellValue.FromBool(false) };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void IsLogical_Number_ReturnsFalse()
    {
        var func = IsLogicalFunction.Instance;
        var args = new[] { CellValue.FromNumber(123) };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void IsLogical_Text_ReturnsFalse()
    {
        var func = IsLogicalFunction.Instance;
        var args = new[] { CellValue.FromString("TRUE") };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void IsLogical_Empty_ReturnsFalse()
    {
        var func = IsLogicalFunction.Instance;
        var args = new[] { CellValue.Empty };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void IsLogical_Error_ReturnsFalse()
    {
        var func = IsLogicalFunction.Instance;
        var args = new[] { CellValue.Error("#VALUE!") };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void IsLogical_WrongArgumentCount_ReturnsError()
    {
        var func = IsLogicalFunction.Instance;
        var args = new[] { CellValue.FromBool(true), CellValue.FromBool(false) };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region ISNONTEXT Tests

    [Fact]
    public void IsNonText_Number_ReturnsTrue()
    {
        var func = IsNonTextFunction.Instance;
        var args = new[] { CellValue.FromNumber(123) };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void IsNonText_Boolean_ReturnsTrue()
    {
        var func = IsNonTextFunction.Instance;
        var args = new[] { CellValue.FromBool(true) };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void IsNonText_Empty_ReturnsTrue()
    {
        var func = IsNonTextFunction.Instance;
        var args = new[] { CellValue.Empty };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void IsNonText_Error_ReturnsTrue()
    {
        var func = IsNonTextFunction.Instance;
        var args = new[] { CellValue.Error("#DIV/0!") };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void IsNonText_Text_ReturnsFalse()
    {
        var func = IsNonTextFunction.Instance;
        var args = new[] { CellValue.FromString("text") };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void IsNonText_EmptyString_ReturnsFalse()
    {
        var func = IsNonTextFunction.Instance;
        var args = new[] { CellValue.FromString("") };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void IsNonText_WrongArgumentCount_ReturnsError()
    {
        var func = IsNonTextFunction.Instance;
        var args = new CellValue[] { };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region TYPE Tests

    [Fact]
    public void Type_Number_Returns1()
    {
        var func = TypeFunction.Instance;
        var args = new[] { CellValue.FromNumber(123) };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void Type_Text_Returns2()
    {
        var func = TypeFunction.Instance;
        var args = new[] { CellValue.FromString("text") };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(2.0, result.NumericValue);
    }

    [Fact]
    public void Type_Boolean_Returns4()
    {
        var func = TypeFunction.Instance;
        var args = new[] { CellValue.FromBool(true) };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(4.0, result.NumericValue);
    }

    [Fact]
    public void Type_Error_Returns16()
    {
        var func = TypeFunction.Instance;
        var args = new[] { CellValue.Error("#VALUE!") };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(16.0, result.NumericValue);
    }

    [Fact]
    public void Type_Empty_Returns1()
    {
        var func = TypeFunction.Instance;
        var args = new[] { CellValue.Empty };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void Type_WrongArgumentCount_ReturnsError()
    {
        var func = TypeFunction.Instance;
        var args = new[] { CellValue.FromNumber(1), CellValue.FromNumber(2) };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region N Tests

    [Fact]
    public void N_Number_ReturnsNumber()
    {
        var func = NFunction.Instance;
        var args = new[] { CellValue.FromNumber(123.45) };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(123.45, result.NumericValue);
    }

    [Fact]
    public void N_True_Returns1()
    {
        var func = NFunction.Instance;
        var args = new[] { CellValue.FromBool(true) };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void N_False_Returns0()
    {
        var func = NFunction.Instance;
        var args = new[] { CellValue.FromBool(false) };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void N_Text_Returns0()
    {
        var func = NFunction.Instance;
        var args = new[] { CellValue.FromString("text") };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void N_EmptyString_Returns0()
    {
        var func = NFunction.Instance;
        var args = new[] { CellValue.FromString("") };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void N_Empty_Returns0()
    {
        var func = NFunction.Instance;
        var args = new[] { CellValue.Empty };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void N_Error_PropagatesError()
    {
        var func = NFunction.Instance;
        var args = new[] { CellValue.Error("#DIV/0!") };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void N_WrongArgumentCount_ReturnsError()
    {
        var func = NFunction.Instance;
        var args = new CellValue[] { };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion
}
