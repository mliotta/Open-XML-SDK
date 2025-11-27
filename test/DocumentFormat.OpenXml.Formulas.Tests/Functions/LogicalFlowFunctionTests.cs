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
            FormulaResult.FromBool(true),
            FormulaResult.FromString("A"),
            FormulaResult.FromBool(false),
            FormulaResult.FromString("B"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("A", result.StringValue);
    }

    [Fact]
    public void Ifs_SecondConditionTrue_ReturnsSecondValue()
    {
        var func = IfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromBool(false),
            FormulaResult.FromString("A"),
            FormulaResult.FromBool(true),
            FormulaResult.FromString("B"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("B", result.StringValue);
    }

    [Fact]
    public void Ifs_NumericConditions_ReturnsCorrectValue()
    {
        var func = IfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(95),
            FormulaResult.FromString("A"),
            FormulaResult.FromNumber(85),
            FormulaResult.FromString("B"),
            FormulaResult.FromNumber(75),
            FormulaResult.FromString("C"),
            FormulaResult.FromBool(true),
            FormulaResult.FromString("F"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("A", result.StringValue);
    }

    [Fact]
    public void Ifs_AllConditionsFalse_ReturnsNA()
    {
        var func = IfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromBool(false),
            FormulaResult.FromString("A"),
            FormulaResult.FromBool(false),
            FormulaResult.FromString("B"),
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
            FormulaResult.FromNumber(0),
            FormulaResult.FromString("Zero"),
            FormulaResult.FromBool(true),
            FormulaResult.FromString("True"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("True", result.StringValue);
    }

    [Fact]
    public void Ifs_NonZeroEvaluatesAsTrue_ReturnsValue()
    {
        var func = IfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(5),
            FormulaResult.FromString("NonZero"),
            FormulaResult.FromBool(true),
            FormulaResult.FromString("True"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("NonZero", result.StringValue);
    }

    [Fact]
    public void Ifs_TextCondition_EvaluatesAsTrue()
    {
        var func = IfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Hello"),
            FormulaResult.FromString("Text Found"),
            FormulaResult.FromBool(true),
            FormulaResult.FromString("Default"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("Text Found", result.StringValue);
    }

    [Fact]
    public void Ifs_EmptyTextCondition_EvaluatesAsFalse()
    {
        var func = IfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString(string.Empty),
            FormulaResult.FromString("Empty"),
            FormulaResult.FromBool(true),
            FormulaResult.FromString("Default"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("Default", result.StringValue);
    }

    [Fact]
    public void Ifs_ErrorInCondition_PropagatesError()
    {
        var func = IfsFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#DIV/0!"),
            FormulaResult.FromString("A"),
            FormulaResult.FromBool(true),
            FormulaResult.FromString("B"),
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
            FormulaResult.FromBool(true),
            FormulaResult.Error("#REF!"),
            FormulaResult.FromBool(true),
            FormulaResult.FromString("B"),
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
            FormulaResult.FromBool(true),
            FormulaResult.FromString("A"),
            FormulaResult.FromBool(false),
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
            FormulaResult.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Ifs_NoArguments_ReturnsError()
    {
        var func = IfsFunction.Instance;
        var args = System.Array.Empty<FormulaResult>();

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
            FormulaResult.FromNumber(1),
            FormulaResult.FromNumber(1),
            FormulaResult.FromString("One"),
            FormulaResult.FromNumber(2),
            FormulaResult.FromString("Two"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("One", result.StringValue);
    }

    [Fact]
    public void Switch_SecondValueMatches_ReturnsSecondResult()
    {
        var func = SwitchFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(2),
            FormulaResult.FromNumber(1),
            FormulaResult.FromString("One"),
            FormulaResult.FromNumber(2),
            FormulaResult.FromString("Two"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("Two", result.StringValue);
    }

    [Fact]
    public void Switch_TextMatch_ReturnsResult()
    {
        var func = SwitchFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Apple"),
            FormulaResult.FromString("Apple"),
            FormulaResult.FromString("Fruit"),
            FormulaResult.FromString("Carrot"),
            FormulaResult.FromString("Vegetable"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("Fruit", result.StringValue);
    }

    [Fact]
    public void Switch_TextMatchCaseInsensitive_ReturnsResult()
    {
        var func = SwitchFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("APPLE"),
            FormulaResult.FromString("apple"),
            FormulaResult.FromString("Fruit"),
            FormulaResult.FromString("Carrot"),
            FormulaResult.FromString("Vegetable"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("Fruit", result.StringValue);
    }

    [Fact]
    public void Switch_NoMatchWithDefault_ReturnsDefault()
    {
        var func = SwitchFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(5),
            FormulaResult.FromNumber(1),
            FormulaResult.FromString("One"),
            FormulaResult.FromNumber(2),
            FormulaResult.FromString("Two"),
            FormulaResult.FromString("Other"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("Other", result.StringValue);
    }

    [Fact]
    public void Switch_NoMatchNoDefault_ReturnsNA()
    {
        var func = SwitchFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(5),
            FormulaResult.FromNumber(1),
            FormulaResult.FromString("One"),
            FormulaResult.FromNumber(2),
            FormulaResult.FromString("Two"),
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
            FormulaResult.FromBool(true),
            FormulaResult.FromBool(true),
            FormulaResult.FromString("Yes"),
            FormulaResult.FromBool(false),
            FormulaResult.FromString("No"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("Yes", result.StringValue);
    }

    [Fact]
    public void Switch_TypeMismatch_NoMatch()
    {
        var func = SwitchFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(1),
            FormulaResult.FromString("1"),
            FormulaResult.FromString("Text One"),
            FormulaResult.FromString("Default"),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Text, result.Type);
        Assert.Equal("Default", result.StringValue);
    }

    [Fact]
    public void Switch_ErrorInExpression_PropagatesError()
    {
        var func = SwitchFunction.Instance;
        var args = new[]
        {
            FormulaResult.Error("#DIV/0!"),
            FormulaResult.FromNumber(1),
            FormulaResult.FromString("One"),
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
            FormulaResult.FromNumber(1),
            FormulaResult.Error("#REF!"),
            FormulaResult.FromString("One"),
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
            FormulaResult.FromNumber(1),
            FormulaResult.FromNumber(1),
            FormulaResult.Error("#N/A"),
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
            FormulaResult.FromNumber(5),
            FormulaResult.FromNumber(1),
            FormulaResult.FromString("One"),
            FormulaResult.Error("#VALUE!"),
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
            FormulaResult.FromNumber(1),
            FormulaResult.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Switch_NoArguments_ReturnsError()
    {
        var func = SwitchFunction.Instance;
        var args = System.Array.Empty<FormulaResult>();

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region TRUE Tests

    [Fact]
    public void True_NoArguments_ReturnsTrue()
    {
        var func = TrueFunction.Instance;
        var args = System.Array.Empty<FormulaResult>();

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void True_WithArguments_ReturnsError()
    {
        var func = TrueFunction.Instance;
        var args = new[] { FormulaResult.FromNumber(1) };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    #endregion

    #region FALSE Tests

    [Fact]
    public void False_NoArguments_ReturnsFalse()
    {
        var func = FalseFunction.Instance;
        var args = System.Array.Empty<FormulaResult>();

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void False_WithArguments_ReturnsError()
    {
        var func = FalseFunction.Instance;
        var args = new[] { FormulaResult.FromNumber(1) };

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
            FormulaResult.FromBool(true),
            FormulaResult.FromBool(false),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void Xor_TwoTrueValues_ReturnsFalse()
    {
        var func = XorFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromBool(true),
            FormulaResult.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void Xor_ThreeTrueValues_ReturnsTrue()
    {
        var func = XorFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromBool(true),
            FormulaResult.FromBool(true),
            FormulaResult.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void Xor_FourTrueValues_ReturnsFalse()
    {
        var func = XorFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromBool(true),
            FormulaResult.FromBool(true),
            FormulaResult.FromBool(true),
            FormulaResult.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void Xor_AllFalse_ReturnsFalse()
    {
        var func = XorFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromBool(false),
            FormulaResult.FromBool(false),
            FormulaResult.FromBool(false),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void Xor_NumericNonZero_EvaluatesAsTrue()
    {
        var func = XorFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(5),
            FormulaResult.FromBool(false),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void Xor_NumericZero_EvaluatesAsFalse()
    {
        var func = XorFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(0),
            FormulaResult.FromBool(false),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void Xor_NonEmptyText_EvaluatesAsTrue()
    {
        var func = XorFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString("Hello"),
            FormulaResult.FromBool(false),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void Xor_EmptyText_EvaluatesAsFalse()
    {
        var func = XorFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromString(string.Empty),
            FormulaResult.FromBool(false),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    [Fact]
    public void Xor_MixedTypes_CountsTrueValues()
    {
        var func = XorFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromBool(true),
            FormulaResult.FromNumber(5),
            FormulaResult.FromString("Text"),
            FormulaResult.FromBool(false),
        };

        var result = func.Execute(null!, args);

        // 3 true values (odd), should return TRUE
        Assert.Equal(FormulaResultType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void Xor_ErrorValue_PropagatesError()
    {
        var func = XorFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromBool(true),
            FormulaResult.Error("#DIV/0!"),
            FormulaResult.FromBool(false),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void Xor_NoArguments_ReturnsError()
    {
        var func = XorFunction.Instance;
        var args = System.Array.Empty<FormulaResult>();

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
            FormulaResult.FromBool(true),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Boolean, result.Type);
        Assert.True(result.BoolValue);
    }

    [Fact]
    public void Xor_SingleFalseArgument_ReturnsFalse()
    {
        var func = XorFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromBool(false),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Boolean, result.Type);
        Assert.False(result.BoolValue);
    }

    #endregion
}
