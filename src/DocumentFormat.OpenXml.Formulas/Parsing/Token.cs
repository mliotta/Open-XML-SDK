// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Parsing;

/// <summary>
/// Represents a token type in a formula.
/// </summary>
public enum TokenType
{
    /// <summary>
    /// Numeric literal.
    /// </summary>
    Number,

    /// <summary>
    /// Cell reference (e.g., A1, $B$2).
    /// </summary>
    CellReference,

    /// <summary>
    /// Cell range (e.g., A1:B10).
    /// </summary>
    Range,

    /// <summary>
    /// Function name.
    /// </summary>
    Function,

    /// <summary>
    /// Left parenthesis.
    /// </summary>
    LeftParen,

    /// <summary>
    /// Right parenthesis.
    /// </summary>
    RightParen,

    /// <summary>
    /// Comma separator.
    /// </summary>
    Comma,

    /// <summary>
    /// Addition operator.
    /// </summary>
    Plus,

    /// <summary>
    /// Subtraction operator.
    /// </summary>
    Minus,

    /// <summary>
    /// Multiplication operator.
    /// </summary>
    Multiply,

    /// <summary>
    /// Division operator.
    /// </summary>
    Divide,

    /// <summary>
    /// Colon for range separator.
    /// </summary>
    Colon,

    /// <summary>
    /// Greater than operator.
    /// </summary>
    GreaterThan,

    /// <summary>
    /// Less than operator.
    /// </summary>
    LessThan,

    /// <summary>
    /// Equals operator.
    /// </summary>
    Equals,

    /// <summary>
    /// String literal.
    /// </summary>
    String,

    /// <summary>
    /// End of formula.
    /// </summary>
    EndOfFormula,

    /// <summary>
    /// Not equal operator.
    /// </summary>
    NotEqual,

    /// <summary>
    /// Less than or equal operator.
    /// </summary>
    LessThanOrEqual,

    /// <summary>
    /// Greater than or equal operator.
    /// </summary>
    GreaterThanOrEqual,

    /// <summary>
    /// String concatenation operator.
    /// </summary>
    Concat,

    /// <summary>
    /// Percentage operator.
    /// </summary>
    Percent,

    /// <summary>
    /// Power operator.
    /// </summary>
    Power,

    /// <summary>
    /// Boolean literal.
    /// </summary>
    Boolean,

    /// <summary>
    /// Error literal.
    /// </summary>
    Error,

    /// <summary>
    /// Sheet reference separator.
    /// </summary>
    SheetSeparator,
}

/// <summary>
/// Represents a lexical token in a formula.
/// </summary>
public class Token
{
    /// <summary>
    /// Initializes a new instance of the <see cref="Token"/> class.
    /// </summary>
    /// <param name="type">The token type.</param>
    /// <param name="text">The token text.</param>
    /// <param name="position">The position in the formula string.</param>
    public Token(TokenType type, string text, int position)
    {
        Type = type;
        Text = text;
        Position = position;
    }

    /// <summary>
    /// Gets the token type.
    /// </summary>
    public TokenType Type { get; }

    /// <summary>
    /// Gets the token text.
    /// </summary>
    public string Text { get; }

    /// <summary>
    /// Gets the position in the formula string.
    /// </summary>
    public int Position { get; }
}
