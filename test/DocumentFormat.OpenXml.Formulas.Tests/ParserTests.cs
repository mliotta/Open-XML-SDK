// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Parsing;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests;

/// <summary>
/// Tests for formula parsing.
/// </summary>
public class ParserTests
{
    [Fact]
    public void Parse_SimpleAddition_Success()
    {
        // Arrange
        var parser = new FormulaParser();

        // Act
        var result = parser.Parse("=A1+B1");

        // Assert
        Assert.IsType<BinaryOpNode>(result);
        var binary = (BinaryOpNode)result;
        Assert.Equal(BinaryOperator.Add, binary.Operator);
        Assert.IsType<CellReferenceNode>(binary.Left);
        Assert.IsType<CellReferenceNode>(binary.Right);
        Assert.Equal("A1", ((CellReferenceNode)binary.Left).Reference);
        Assert.Equal("B1", ((CellReferenceNode)binary.Right).Reference);
    }

    [Fact]
    public void Parse_MultipleOperations_RespectsOperatorPrecedence()
    {
        // Arrange
        var parser = new FormulaParser();

        // Act
        var result = parser.Parse("=A1*B1-C1");

        // Assert
        Assert.IsType<BinaryOpNode>(result);
        var binary = (BinaryOpNode)result;
        Assert.Equal(BinaryOperator.Subtract, binary.Operator);

        // Left side should be A1*B1
        Assert.IsType<BinaryOpNode>(binary.Left);
        var leftOp = (BinaryOpNode)binary.Left;
        Assert.Equal(BinaryOperator.Multiply, leftOp.Operator);

        // Right side should be C1
        Assert.IsType<CellReferenceNode>(binary.Right);
        Assert.Equal("C1", ((CellReferenceNode)binary.Right).Reference);
    }

    [Fact]
    public void Parse_SumFunction_Success()
    {
        // Arrange
        var parser = new FormulaParser();

        // Act
        var result = parser.Parse("=SUM(A1:A10)");

        // Assert
        Assert.IsType<FunctionCallNode>(result);
        var func = (FunctionCallNode)result;
        Assert.Equal("SUM", func.FunctionName);
        Assert.Single(func.Arguments);
        Assert.IsType<RangeNode>(func.Arguments[0]);
        var range = (RangeNode)func.Arguments[0];
        Assert.Equal("A1", range.Start);
        Assert.Equal("A10", range.End);
    }

    [Fact]
    public void Parse_AverageFunction_Success()
    {
        // Arrange
        var parser = new FormulaParser();

        // Act
        var result = parser.Parse("=AVERAGE(A1:A10)");

        // Assert
        Assert.IsType<FunctionCallNode>(result);
        var func = (FunctionCallNode)result;
        Assert.Equal("AVERAGE", func.FunctionName);
        Assert.Single(func.Arguments);
    }

    [Fact]
    public void Parse_IfFunction_Success()
    {
        // Arrange
        var parser = new FormulaParser();

        // Act
        var result = parser.Parse("=IF(A1>10, B1, C1)");

        // Assert
        Assert.IsType<FunctionCallNode>(result);
        var func = (FunctionCallNode)result;
        Assert.Equal("IF", func.FunctionName);
        Assert.Equal(3, func.Arguments.Count);

        // First argument is comparison
        Assert.IsType<BinaryOpNode>(func.Arguments[0]);
        var comparison = (BinaryOpNode)func.Arguments[0];
        Assert.Equal(BinaryOperator.GreaterThan, comparison.Operator);
    }

    [Fact]
    public void Parse_Parentheses_Success()
    {
        // Arrange
        var parser = new FormulaParser();

        // Act
        var result = parser.Parse("=(A1+B1)*C1");

        // Assert
        Assert.IsType<BinaryOpNode>(result);
        var binary = (BinaryOpNode)result;
        Assert.Equal(BinaryOperator.Multiply, binary.Operator);

        // Left side should be (A1+B1)
        Assert.IsType<BinaryOpNode>(binary.Left);
        var leftOp = (BinaryOpNode)binary.Left;
        Assert.Equal(BinaryOperator.Add, leftOp.Operator);
    }

    [Fact]
    public void Parse_MultipleAdditions_Success()
    {
        // Arrange
        var parser = new FormulaParser();

        // Act
        var result = parser.Parse("=A1+B1+C1+D1");

        // Assert
        Assert.IsType<BinaryOpNode>(result);
        // Should be left-associative: ((A1+B1)+C1)+D1
    }

    [Fact]
    public void Parse_MultipleSums_Success()
    {
        // Arrange
        var parser = new FormulaParser();

        // Act
        var result = parser.Parse("=SUM(A1:A5)+SUM(B1:B5)");

        // Assert
        Assert.IsType<BinaryOpNode>(result);
        var binary = (BinaryOpNode)result;
        Assert.Equal(BinaryOperator.Add, binary.Operator);
        Assert.IsType<FunctionCallNode>(binary.Left);
        Assert.IsType<FunctionCallNode>(binary.Right);
    }

    [Fact]
    public void Parse_NestedIfWithSum_Success()
    {
        // Arrange
        var parser = new FormulaParser();

        // Act
        var result = parser.Parse("=IF(A1>0, SUM(B1:B10), 0)");

        // Assert
        Assert.IsType<FunctionCallNode>(result);
        var func = (FunctionCallNode)result;
        Assert.Equal("IF", func.FunctionName);
        Assert.Equal(3, func.Arguments.Count);

        // Second argument should be SUM
        Assert.IsType<FunctionCallNode>(func.Arguments[1]);
        var sumFunc = (FunctionCallNode)func.Arguments[1];
        Assert.Equal("SUM", sumFunc.FunctionName);
    }

    [Fact]
    public void Parse_Division_Success()
    {
        // Arrange
        var parser = new FormulaParser();

        // Act
        var result = parser.Parse("=A1/B1");

        // Assert
        Assert.IsType<BinaryOpNode>(result);
        var binary = (BinaryOpNode)result;
        Assert.Equal(BinaryOperator.Divide, binary.Operator);
    }

    [Fact]
    public void Lexer_TokenizesNumbers_Success()
    {
        // Arrange
        var lexer = new Lexer("=123.45");

        // Act
        var tokens = lexer.Tokenize();

        // Assert
        Assert.Equal(2, tokens.Count); // Number + EndOfFormula
        Assert.Equal(TokenType.Number, tokens[0].Type);
        Assert.Equal("123.45", tokens[0].Text);
    }

    [Fact]
    public void Lexer_TokenizesCellReferences_Success()
    {
        // Arrange
        var lexer = new Lexer("=A1+$B$2");

        // Act
        var tokens = lexer.Tokenize();

        // Assert
        Assert.Equal(4, tokens.Count); // A1 + $B$2 + EndOfFormula
        Assert.Equal(TokenType.CellReference, tokens[0].Type);
        Assert.Equal("A1", tokens[0].Text);
        Assert.Equal(TokenType.Plus, tokens[1].Type);
        Assert.Equal(TokenType.CellReference, tokens[2].Type);
        Assert.Equal("$B$2", tokens[2].Text);
    }

    [Fact]
    public void Lexer_TokenizesFunctions_Success()
    {
        // Arrange
        var lexer = new Lexer("=SUM(A1:A10)");

        // Act
        var tokens = lexer.Tokenize();

        // Assert
        Assert.Equal(6, tokens.Count); // SUM ( A1 : A10 ) EndOfFormula
        Assert.Equal(TokenType.Function, tokens[0].Type);
        Assert.Equal("SUM", tokens[0].Text);
        Assert.Equal(TokenType.LeftParen, tokens[1].Type);
        Assert.Equal(TokenType.CellReference, tokens[2].Type);
        Assert.Equal(TokenType.Colon, tokens[3].Type);
        Assert.Equal(TokenType.CellReference, tokens[4].Type);
        Assert.Equal(TokenType.RightParen, tokens[5].Type);
    }
}
