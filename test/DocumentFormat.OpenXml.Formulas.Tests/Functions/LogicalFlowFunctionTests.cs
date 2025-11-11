// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for advanced logical flow control functions (IFS, SWITCH, XOR).
/// </summary>
public class LogicalFlowFunctionTests
{
    #region IFS Tests

    [Fact]
    public void Ifs_FirstConditionTrue_ReturnsFirstValue()
    {
        var func = IfsFunction.Instance;
        var args = new[]
        {
            CellValue.FromBool(true),
            CellValue.FromString("A"),
            CellValue.FromBool(false),
            CellValue.FromString("B"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("A", result.StringValue);
    }

    [Fact]
    public void Ifs_SecondConditionTrue_ReturnsSecondValue()
    {
        var func = IfsFunction.Instance;
        var args = new[]
        {
            CellValue.FromBool(false),
            CellValue.FromString("A"),
            CellValue.FromBool(true),
            CellValue.FromString("B"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("B", result.StringValue);
    }

    [Fact]
    public void Ifs_NumericConditions_ReturnsCorrectValue()
    {
        var func = IfsFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(95),
            CellValue.FromString("A"),
            CellValue.FromNumber(85),
            CellValue.FromString("B"),
            CellValue.FromNumber(75),
            CellValue.FromString("C"),
            CellValue.FromBool(true),
            CellValue.FromString("F"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("A", result.StringValue);
    }

    [Fact]
    public void Ifs_AllConditionsFalse_ReturnsNA()
    {
        var func = IfsFunction.Instance;
        var args = new[]
        {
            CellValue.FromBool(false),
            CellValue.FromString("A"),
            CellValue.FromBool(false),
            CellValue.FromString("B"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    [Fact]
    public void Ifs_ZeroEvaluatesAsFalse_ContinuesToNextCondition()
    {
        var func = IfsFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0),
            CellValue.FromString("Zero"),
            CellValue.FromBool(true),
            CellValue.FromString("True"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("True", result.StringValue);
    }

    [Fact]
    public void Ifs_NonZeroEvaluatesAsTrue_ReturnsValue()
    {
        var func = IfsFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromString("NonZero"),
            CellValue.FromBool(true),
            CellValue.FromString("True"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("NonZero", result.StringValue);
    }

    [Fact]
    public void Ifs_TextCondition_EvaluatesAsTrue()
    {
        var func = IfsFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("Hello"),
            CellValue.FromString("Text Found"),
            CellValue.FromBool(true),
            CellValue.FromString("Default"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("Text Found", result.StringValue);
    }

    [Fact]
    public void Ifs_EmptyTextCondition_EvaluatesAsFalse()
    {
        var func = IfsFunction.Instance;
        var args = new[]
        {
            CellValue.FromString(string.Empty),
            CellValue.FromString("Empty"),
            CellValue.FromBool(true),
            CellValue.FromString("Default"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("Default", result.StringValue);
    }

    [Fact]
    public void Ifs_ErrorInCondition_PropagatesError()
    {
        var func = IfsFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#DIV/0!"),
            CellValue.FromString("A"),
            CellValue.FromBool(true),
            CellValue.FromString("B"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void Ifs_ErrorInValue_PropagatesError()
    {
        var func = IfsFunction.Instance;
        var args = new[]
        {
            CellValue.FromBool(true),
            CellValue.Error("#REF!"),
            CellValue.FromBool(true),
            CellValue.FromString("B"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    [Fact]
    public void Ifs_OddArgumentCount_ReturnsError()
    {
        var func = IfsFunction.Instance;
        var args = new[]
        {
            CellValue.FromBool(true),
            CellValue.FromString("A"),
            CellValue.FromBool(false),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Ifs_InsufficientArguments_ReturnsError()
    {
        var func = IfsFunction.Instance;
        var args = new[]
        {
            CellValue.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Ifs_NoArguments_ReturnsError()
    {
        var func = IfsFunction.Instance;
        var args = System.Array.Empty<CellValue>();

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region SWITCH Tests

    [Fact]
    public void Switch_FirstValueMatches_ReturnsFirstResult()
    {
        var func = SwitchFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(1),
            CellValue.FromString("One"),
            CellValue.FromNumber(2),
            CellValue.FromString("Two"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("One", result.StringValue);
    }

    [Fact]
    public void Switch_SecondValueMatches_ReturnsSecondResult()
    {
        var func = SwitchFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(2),
            CellValue.FromNumber(1),
            CellValue.FromString("One"),
            CellValue.FromNumber(2),
            CellValue.FromString("Two"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("Two", result.StringValue);
    }

    [Fact]
    public void Switch_TextMatch_ReturnsResult()
    {
        var func = SwitchFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("Apple"),
            CellValue.FromString("Apple"),
            CellValue.FromString("Fruit"),
            CellValue.FromString("Carrot"),
            CellValue.FromString("Vegetable"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("Fruit", result.StringValue);
    }

    [Fact]
    public void Switch_TextMatchCaseInsensitive_ReturnsResult()
    {
        var func = SwitchFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("APPLE"),
            CellValue.FromString("apple"),
            CellValue.FromString("Fruit"),
            CellValue.FromString("Carrot"),
            CellValue.FromString("Vegetable"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("Fruit", result.StringValue);
    }

    [Fact]
    public void Switch_NoMatchWithDefault_ReturnsDefault()
    {
        var func = SwitchFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(1),
            CellValue.FromString("One"),
            CellValue.FromNumber(2),
            CellValue.FromString("Two"),
            CellValue.FromString("Other"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("Other", result.StringValue);
    }

    [Fact]
    public void Switch_NoMatchNoDefault_ReturnsNA()
    {
        var func = SwitchFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(1),
            CellValue.FromString("One"),
            CellValue.FromNumber(2),
            CellValue.FromString("Two"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    [Fact]
    public void Switch_BooleanMatch_ReturnsResult()
    {
        var func = SwitchFunction.Instance;
        var args = new[]
        {
            CellValue.FromBool(true),
            CellValue.FromBool(true),
            CellValue.FromString("Yes"),
            CellValue.FromBool(false),
            CellValue.FromString("No"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("Yes", result.StringValue);
    }

    [Fact]
    public void Switch_TypeMismatch_NoMatch()
    {
        var func = SwitchFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromString("1"),
            CellValue.FromString("Text One"),
            CellValue.FromString("Default"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Text, result.Type);
        Assert.Equal("Default", result.StringValue);
    }

    [Fact]
    public void Switch_ErrorInExpression_PropagatesError()
    {
        var func = SwitchFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(1),
            CellValue.FromString("One"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void Switch_ErrorInValue_PropagatesError()
    {
        var func = SwitchFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.Error("#REF!"),
            CellValue.FromString("One"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    [Fact]
    public void Switch_ErrorInResult_PropagatesError()
    {
        var func = SwitchFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(1),
            CellValue.Error("#N/A"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    [Fact]
    public void Switch_ErrorInDefault_PropagatesError()
    {
        var func = SwitchFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromNumber(1),
            CellValue.FromString("One"),
            CellValue.Error("#VALUE!"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Switch_InsufficientArguments_ReturnsError()
    {
        var func = SwitchFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Switch_NoArguments_ReturnsError()
    {
        var func = SwitchFunction.Instance;
        var args = System.Array.Empty<CellValue>();

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region XOR Tests

    [Fact]
    public void Xor_OneTrueValue_ReturnsTrue()
    {
        var func = XorFunction.Instance;
        var args = new[]
        {
            CellValue.FromBool(true),
            CellValue.FromBool(false),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void Xor_TwoTrueValues_ReturnsFalse()
    {
        var func = XorFunction.Instance;
        var args = new[]
        {
            CellValue.FromBool(true),
            CellValue.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void Xor_ThreeTrueValues_ReturnsTrue()
    {
        var func = XorFunction.Instance;
        var args = new[]
        {
            CellValue.FromBool(true),
            CellValue.FromBool(true),
            CellValue.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void Xor_FourTrueValues_ReturnsFalse()
    {
        var func = XorFunction.Instance;
        var args = new[]
        {
            CellValue.FromBool(true),
            CellValue.FromBool(true),
            CellValue.FromBool(true),
            CellValue.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void Xor_AllFalse_ReturnsFalse()
    {
        var func = XorFunction.Instance;
        var args = new[]
        {
            CellValue.FromBool(false),
            CellValue.FromBool(false),
            CellValue.FromBool(false),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void Xor_NumericNonZero_EvaluatesAsTrue()
    {
        var func = XorFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(5),
            CellValue.FromBool(false),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void Xor_NumericZero_EvaluatesAsFalse()
    {
        var func = XorFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0),
            CellValue.FromBool(false),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void Xor_NonEmptyText_EvaluatesAsTrue()
    {
        var func = XorFunction.Instance;
        var args = new[]
        {
            CellValue.FromString("Hello"),
            CellValue.FromBool(false),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void Xor_EmptyText_EvaluatesAsFalse()
    {
        var func = XorFunction.Instance;
        var args = new[]
        {
            CellValue.FromString(string.Empty),
            CellValue.FromBool(false),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void Xor_MixedTypes_CountsTrueValues()
    {
        var func = XorFunction.Instance;
        var args = new[]
        {
            CellValue.FromBool(true),
            CellValue.FromNumber(5),
            CellValue.FromString("Text"),
            CellValue.FromBool(false),
        };

        var result = func.Execute(null!, args);

        // 3 true values (odd), should return TRUE
        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void Xor_ErrorValue_PropagatesError()
    {
        var func = XorFunction.Instance;
        var args = new[]
        {
            CellValue.FromBool(true),
            CellValue.Error("#DIV/0!"),
            CellValue.FromBool(false),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void Xor_NoArguments_ReturnsError()
    {
        var func = XorFunction.Instance;
        var args = System.Array.Empty<CellValue>();

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Xor_SingleArgument_ReturnsTrue()
    {
        var func = XorFunction.Instance;
        var args = new[]
        {
            CellValue.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void Xor_SingleFalseArgument_ReturnsFalse()
    {
        var func = XorFunction.Instance;
        var args = new[]
        {
            CellValue.FromBool(false),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    #endregion
}
