// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Parsing;

/// <summary>
/// Tokenizes Excel formula strings.
/// </summary>
public class Lexer
{
    private readonly string _formula;
    private int _position;

    /// <summary>
    /// Initializes a new instance of the <see cref="Lexer"/> class.
    /// </summary>
    /// <param name="formula">The formula string to tokenize.</param>
    public Lexer(string formula)
    {
        _formula = formula ?? throw new ArgumentNullException(nameof(formula));
        _position = 0;

        // Skip leading '=' if present
        if (_formula.StartsWith("=", StringComparison.Ordinal))
        {
            _position = 1;
        }
    }

    /// <summary>
    /// Tokenizes the formula into a list of tokens.
    /// </summary>
    /// <returns>List of tokens.</returns>
    public List<Token> Tokenize()
    {
        var tokens = new List<Token>();

        while (_position < _formula.Length)
        {
            SkipWhitespace();

            if (_position >= _formula.Length)
            {
                break;
            }

            var token = ReadNextToken();
            if (token != null)
            {
                tokens.Add(token);
            }
        }

        tokens.Add(new Token(TokenType.EndOfFormula, string.Empty, _position));
        return tokens;
    }

    private void SkipWhitespace()
    {
        while (_position < _formula.Length && char.IsWhiteSpace(_formula[_position]))
        {
            _position++;
        }
    }

    private Token? ReadNextToken()
    {
        var currentChar = _formula[_position];
        var startPosition = _position;

        // Operators
        switch (currentChar)
        {
            case '+':
                _position++;
                return new Token(TokenType.Plus, "+", startPosition);
            case '-':
                _position++;
                return new Token(TokenType.Minus, "-", startPosition);
            case '*':
                _position++;
                return new Token(TokenType.Multiply, "*", startPosition);
            case '/':
                _position++;
                return new Token(TokenType.Divide, "/", startPosition);
            case '(':
                _position++;
                return new Token(TokenType.LeftParen, "(", startPosition);
            case ')':
                _position++;
                return new Token(TokenType.RightParen, ")", startPosition);
            case ',':
                _position++;
                return new Token(TokenType.Comma, ",", startPosition);
            case ':':
                _position++;
                return new Token(TokenType.Colon, ":", startPosition);
            case '>':
                _position++;
                if (_position < _formula.Length && _formula[_position] == '=')
                {
                    _position++;
                    return new Token(TokenType.GreaterThanOrEqual, ">=", startPosition);
                }

                return new Token(TokenType.GreaterThan, ">", startPosition);
            case '<':
                _position++;
                if (_position < _formula.Length && _formula[_position] == '=')
                {
                    _position++;
                    return new Token(TokenType.LessThanOrEqual, "<=", startPosition);
                }

                if (_position < _formula.Length && _formula[_position] == '>')
                {
                    _position++;
                    return new Token(TokenType.NotEqual, "<>", startPosition);
                }

                return new Token(TokenType.LessThan, "<", startPosition);
            case '=':
                _position++;
                return new Token(TokenType.Equals, "=", startPosition);
            case '"':
                return ReadString();
            case '&':
                _position++;
                return new Token(TokenType.Concat, "&", startPosition);
            case '%':
                _position++;
                return new Token(TokenType.Percent, "%", startPosition);
            case '^':
                _position++;
                return new Token(TokenType.Power, "^", startPosition);
            case '!':
                _position++;
                return new Token(TokenType.SheetSeparator, "!", startPosition);
            case '\'':
                return ReadQuotedSheetName();
            case '#':
                return ReadError();
        }

        // Numbers
        if (char.IsDigit(currentChar) || currentChar == '.')
        {
            return ReadNumber();
        }

        // Cell references or function names
        if (char.IsLetter(currentChar) || currentChar == '$')
        {
            return ReadIdentifierOrCellReference();
        }

        throw new ParserException($"Unexpected character '{currentChar}' at position {_position}");
    }

    private Token ReadNumber()
    {
        var startPosition = _position;
        var sb = new StringBuilder();

        while (_position < _formula.Length &&
               (char.IsDigit(_formula[_position]) || _formula[_position] == '.'))
        {
            sb.Append(_formula[_position]);
            _position++;
        }

        return new Token(TokenType.Number, sb.ToString(), startPosition);
    }

    private Token ReadString()
    {
        var startPosition = _position;
        var sb = new StringBuilder();

        _position++; // Skip opening quote

        while (_position < _formula.Length && _formula[_position] != '"')
        {
            sb.Append(_formula[_position]);
            _position++;
        }

        if (_position >= _formula.Length)
        {
            throw new ParserException($"Unterminated string at position {startPosition}");
        }

        _position++; // Skip closing quote

        return new Token(TokenType.String, sb.ToString(), startPosition);
    }

    private Token ReadIdentifierOrCellReference()
    {
        var startPosition = _position;
        var sb = new StringBuilder();

        // Handle absolute cell references ($A$1)
        if (_formula[_position] == '$')
        {
            sb.Append(_formula[_position]);
            _position++;
        }

        // Read letters (column part)
        while (_position < _formula.Length && char.IsLetter(_formula[_position]))
        {
            sb.Append(_formula[_position]);
            _position++;
        }

        // Check if this is a cell reference (followed by optional $ and numbers)
        var hasRowPart = false;
        if (_position < _formula.Length && (_formula[_position] == '$' || char.IsDigit(_formula[_position])))
        {
            if (_formula[_position] == '$')
            {
                sb.Append(_formula[_position]);
                _position++;
            }

            while (_position < _formula.Length && char.IsDigit(_formula[_position]))
            {
                sb.Append(_formula[_position]);
                _position++;
                hasRowPart = true;
            }
        }

        var text = sb.ToString();

        // If it has a row part, it's a cell reference
        if (hasRowPart)
        {
            return new Token(TokenType.CellReference, text, startPosition);
        }

        // Check for boolean literals
        if (string.Equals(text, "TRUE", StringComparison.OrdinalIgnoreCase))
        {
            return new Token(TokenType.Boolean, "TRUE", startPosition);
        }

        if (string.Equals(text, "FALSE", StringComparison.OrdinalIgnoreCase))
        {
            return new Token(TokenType.Boolean, "FALSE", startPosition);
        }

        // Otherwise, check if next char is '(' to determine if it's a function
        SkipWhitespace();
        if (_position < _formula.Length && _formula[_position] == '(')
        {
            return new Token(TokenType.Function, text, startPosition);
        }

        // Could be a named range or cell reference without row (error case)
        return new Token(TokenType.CellReference, text, startPosition);
    }

    private Token ReadError()
    {
        var startPosition = _position;
        var sb = new StringBuilder();

        sb.Append(_formula[_position]); // '#'
        _position++;

        // Read until we hit a non-alphanumeric character or end
        while (_position < _formula.Length &&
               (char.IsLetterOrDigit(_formula[_position]) || _formula[_position] == '/' || _formula[_position] == '?'))
        {
            sb.Append(_formula[_position]);
            _position++;
        }

        // Add trailing '!' if present
        if (_position < _formula.Length && _formula[_position] == '!')
        {
            sb.Append(_formula[_position]);
            _position++;
        }

        return new Token(TokenType.Error, sb.ToString(), startPosition);
    }

    private Token ReadQuotedSheetName()
    {
        var startPosition = _position;
        var sb = new StringBuilder();

        _position++; // Skip opening quote

        while (_position < _formula.Length && _formula[_position] != '\'')
        {
            if (_formula[_position] == '\'' && _position + 1 < _formula.Length && _formula[_position + 1] == '\'')
            {
                // Escaped quote
                sb.Append('\'');
                _position += 2;
            }
            else
            {
                sb.Append(_formula[_position]);
                _position++;
            }
        }

        if (_position >= _formula.Length)
        {
            throw new ParserException($"Unterminated quoted sheet name at position {startPosition}");
        }

        _position++; // Skip closing quote

        return new Token(TokenType.String, sb.ToString(), startPosition);
    }
}
