// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Parsing;

/// <summary>
/// Base class for formula AST nodes.
/// </summary>
public abstract class FormulaNode
{
}

/// <summary>
/// Represents a binary operation node.
/// </summary>
public class BinaryOpNode : FormulaNode
{
    /// <summary>
    /// Gets or sets the left operand.
    /// </summary>
    public FormulaNode Left { get; set; } = null!;

    /// <summary>
    /// Gets or sets the right operand.
    /// </summary>
    public FormulaNode Right { get; set; } = null!;

    /// <summary>
    /// Gets or sets the operator.
    /// </summary>
    public BinaryOperator Operator { get; set; }
}

/// <summary>
/// Binary operators.
/// </summary>
public enum BinaryOperator
{
    /// <summary>
    /// Addition.
    /// </summary>
    Add,

    /// <summary>
    /// Subtraction.
    /// </summary>
    Subtract,

    /// <summary>
    /// Multiplication.
    /// </summary>
    Multiply,

    /// <summary>
    /// Division.
    /// </summary>
    Divide,

    /// <summary>
    /// Greater than.
    /// </summary>
    GreaterThan,

    /// <summary>
    /// Less than.
    /// </summary>
    LessThan,

    /// <summary>
    /// Equals.
    /// </summary>
    Equals,

    /// <summary>
    /// Not equal.
    /// </summary>
    NotEqual,

    /// <summary>
    /// Less than or equal.
    /// </summary>
    LessThanOrEqual,

    /// <summary>
    /// Greater than or equal.
    /// </summary>
    GreaterThanOrEqual,

    /// <summary>
    /// String concatenation.
    /// </summary>
    Concat,

    /// <summary>
    /// Power.
    /// </summary>
    Power,
}

/// <summary>
/// Represents a function call node.
/// </summary>
public class FunctionCallNode : FormulaNode
{
    /// <summary>
    /// Gets or sets the function name.
    /// </summary>
    public string FunctionName { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the function arguments.
    /// </summary>
    public List<FormulaNode> Arguments { get; set; } = new();
}

/// <summary>
/// Represents a cell reference node.
/// </summary>
public class CellReferenceNode : FormulaNode
{
    /// <summary>
    /// Gets or sets the cell reference (e.g., "A1", "$B$2").
    /// </summary>
    public string Reference { get; set; } = string.Empty;
}

/// <summary>
/// Represents a range node.
/// </summary>
public class RangeNode : FormulaNode
{
    /// <summary>
    /// Gets or sets the start cell reference.
    /// </summary>
    public string Start { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the end cell reference.
    /// </summary>
    public string End { get; set; } = string.Empty;
}

/// <summary>
/// Represents a literal value node.
/// </summary>
public class LiteralNode : FormulaNode
{
    /// <summary>
    /// Gets or sets the literal value.
    /// </summary>
    public object Value { get; set; } = null!;
}

/// <summary>
/// Represents a unary operation node.
/// </summary>
public class UnaryOpNode : FormulaNode
{
    /// <summary>
    /// Gets or sets the operand.
    /// </summary>
    public FormulaNode Operand { get; set; } = null!;

    /// <summary>
    /// Gets or sets the operator.
    /// </summary>
    public UnaryOperator Operator { get; set; }
}

/// <summary>
/// Unary operators.
/// </summary>
public enum UnaryOperator
{
    /// <summary>
    /// Negation.
    /// </summary>
    Negate,

    /// <summary>
    /// Plus.
    /// </summary>
    Plus,

    /// <summary>
    /// Percentage.
    /// </summary>
    Percent,
}

/// <summary>
/// Represents a sheet reference node (e.g., Sheet1!A1).
/// </summary>
public class SheetReferenceNode : FormulaNode
{
    /// <summary>
    /// Gets or sets the sheet name.
    /// </summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the cell reference.
    /// </summary>
    public string CellReference { get; set; } = string.Empty;
}
