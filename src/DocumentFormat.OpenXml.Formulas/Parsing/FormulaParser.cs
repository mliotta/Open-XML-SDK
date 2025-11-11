// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Globalization;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Parsing;

/// <summary>
/// Parses Excel formulas into abstract syntax trees.
/// </summary>
public class FormulaParser
{
    private List<Token> _tokens = new();
    private int _position;

    /// <summary>
    /// Parses a formula string into an AST.
    /// </summary>
    /// <param name="formula">The formula string to parse.</param>
    /// <returns>The root node of the AST.</returns>
    public FormulaNode Parse(string formula)
    {
        var lexer = new Lexer(formula);
        _tokens = lexer.Tokenize();
        _position = 0;

        return ParseExpression();
    }

    private Token CurrentToken => _tokens[_position];

    private void Advance()
    {
        if (_position < _tokens.Count - 1)
        {
            _position++;
        }
    }

    private void Expect(TokenType type)
    {
        if (CurrentToken.Type != type)
        {
            throw new ParserException($"Expected {type} but got {CurrentToken.Type} at position {CurrentToken.Position}");
        }

        Advance();
    }

    private FormulaNode ParseExpression()
    {
        return ParseComparison();
    }

    private FormulaNode ParseComparison()
    {
        var left = ParseConcat();

        while (CurrentToken.Type == TokenType.GreaterThan ||
               CurrentToken.Type == TokenType.LessThan ||
               CurrentToken.Type == TokenType.Equals ||
               CurrentToken.Type == TokenType.NotEqual ||
               CurrentToken.Type == TokenType.LessThanOrEqual ||
               CurrentToken.Type == TokenType.GreaterThanOrEqual)
        {
            var op = CurrentToken.Type switch
            {
                TokenType.GreaterThan => BinaryOperator.GreaterThan,
                TokenType.LessThan => BinaryOperator.LessThan,
                TokenType.Equals => BinaryOperator.Equals,
                TokenType.NotEqual => BinaryOperator.NotEqual,
                TokenType.LessThanOrEqual => BinaryOperator.LessThanOrEqual,
                TokenType.GreaterThanOrEqual => BinaryOperator.GreaterThanOrEqual,
                _ => throw new ParserException($"Unexpected operator {CurrentToken.Type}"),
            };

            Advance();
            var right = ParseConcat();

            left = new BinaryOpNode
            {
                Left = left,
                Right = right,
                Operator = op,
            };
        }

        return left;
    }

    private FormulaNode ParseConcat()
    {
        var left = ParseAdditive();

        while (CurrentToken.Type == TokenType.Concat)
        {
            Advance();
            var right = ParseAdditive();

            left = new BinaryOpNode
            {
                Left = left,
                Right = right,
                Operator = BinaryOperator.Concat,
            };
        }

        return left;
    }

    private FormulaNode ParseAdditive()
    {
        var left = ParseMultiplicative();

        while (CurrentToken.Type == TokenType.Plus || CurrentToken.Type == TokenType.Minus)
        {
            var op = CurrentToken.Type == TokenType.Plus ? BinaryOperator.Add : BinaryOperator.Subtract;
            Advance();
            var right = ParseMultiplicative();

            left = new BinaryOpNode
            {
                Left = left,
                Right = right,
                Operator = op,
            };
        }

        return left;
    }

    private FormulaNode ParseMultiplicative()
    {
        var left = ParsePower();

        while (CurrentToken.Type == TokenType.Multiply || CurrentToken.Type == TokenType.Divide)
        {
            var op = CurrentToken.Type == TokenType.Multiply ? BinaryOperator.Multiply : BinaryOperator.Divide;
            Advance();
            var right = ParsePower();

            left = new BinaryOpNode
            {
                Left = left,
                Right = right,
                Operator = op,
            };
        }

        return left;
    }

    private FormulaNode ParsePower()
    {
        var left = ParseUnary();

        // Power is right-associative
        if (CurrentToken.Type == TokenType.Power)
        {
            Advance();
            var right = ParsePower(); // Recursive for right-associativity

            return new BinaryOpNode
            {
                Left = left,
                Right = right,
                Operator = BinaryOperator.Power,
            };
        }

        return left;
    }

    private FormulaNode ParseUnary()
    {
        if (CurrentToken.Type == TokenType.Minus)
        {
            Advance();
            var operand = ParseUnary();

            return new UnaryOpNode
            {
                Operand = operand,
                Operator = UnaryOperator.Negate,
            };
        }

        if (CurrentToken.Type == TokenType.Plus)
        {
            Advance();
            var operand = ParseUnary();

            return new UnaryOpNode
            {
                Operand = operand,
                Operator = UnaryOperator.Plus,
            };
        }

        return ParsePostfix();
    }

    private FormulaNode ParsePostfix()
    {
        var node = ParsePrimary();

        // Handle percentage operator
        if (CurrentToken.Type == TokenType.Percent)
        {
            Advance();
            return new UnaryOpNode
            {
                Operand = node,
                Operator = UnaryOperator.Percent,
            };
        }

        return node;
    }

    private FormulaNode ParsePrimary()
    {
        var token = CurrentToken;

        switch (token.Type)
        {
            case TokenType.Number:
                Advance();
                if (!double.TryParse(token.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
                {
                    throw new ParserException($"Invalid number '{token.Text}' at position {token.Position}");
                }

                return new LiteralNode { Value = number };

            case TokenType.String:
                Advance();
                return new LiteralNode { Value = token.Text };

            case TokenType.Boolean:
                Advance();
                var boolValue = string.Equals(token.Text, "TRUE", StringComparison.OrdinalIgnoreCase);
                return new LiteralNode { Value = boolValue };

            case TokenType.Error:
                Advance();
                return new LiteralNode { Value = CellValue.Error(token.Text) };

            case TokenType.Function:
                return ParseFunction();

            case TokenType.CellReference:
                return ParseCellReferenceOrRange();

            case TokenType.LeftParen:
                Advance();
                var expr = ParseExpression();
                Expect(TokenType.RightParen);
                return expr;

            default:
                throw new ParserException($"Unexpected token {token.Type} at position {token.Position}");
        }
    }

    private FormulaNode ParseFunction()
    {
        var functionName = CurrentToken.Text;
        Advance();

        Expect(TokenType.LeftParen);

        var arguments = new List<FormulaNode>();

        if (CurrentToken.Type != TokenType.RightParen)
        {
            arguments.Add(ParseExpression());

            while (CurrentToken.Type == TokenType.Comma)
            {
                Advance();
                arguments.Add(ParseExpression());
            }
        }

        Expect(TokenType.RightParen);

        return new FunctionCallNode
        {
            FunctionName = functionName,
            Arguments = arguments,
        };
    }

    private FormulaNode ParseCellReferenceOrRange()
    {
        var startRef = CurrentToken.Text;
        Advance();

        // Check if this is a range (A1:B10)
        if (CurrentToken.Type == TokenType.Colon)
        {
            Advance();

            if (CurrentToken.Type != TokenType.CellReference)
            {
                throw new ParserException($"Expected cell reference after ':' at position {CurrentToken.Position}");
            }

            var endRef = CurrentToken.Text;
            Advance();

            return new RangeNode
            {
                Start = startRef,
                End = endRef,
            };
        }

        return new CellReferenceNode { Reference = startRef };
    }
}
