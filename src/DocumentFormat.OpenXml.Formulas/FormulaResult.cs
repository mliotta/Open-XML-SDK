// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation;

/// <summary>
/// Represents an evaluated formula result.
/// </summary>
/// <remarks>
/// This is intentionally a separate type from <see cref="DocumentFormat.OpenXml.Spreadsheet.CellValue"/>.
/// The SDK's CellValue is a class representing the XML element for serialization,
/// while FormulaResult is a value type optimized for formula evaluation results.
/// </remarks>
public readonly struct FormulaResult : IEquatable<FormulaResult>
{
    /// <summary>
    /// Gets the type of the cell value.
    /// </summary>
    public FormulaResultType Type { get; }

    /// <summary>
    /// Gets the raw value object.
    /// </summary>
    public object? Value { get; }

    /// <summary>
    /// Gets the numeric value. Returns 0 if not a number.
    /// </summary>
    public double NumericValue => Type == FormulaResultType.Number ? (double)Value! : 0;

    /// <summary>
    /// Gets the string value.
    /// </summary>
    /// <remarks>
    /// For Boolean values, returns "TRUE" or "FALSE" (uppercase) to match Excel behavior.
    /// </remarks>
    public string StringValue => Type == FormulaResultType.Boolean
        ? ((bool)Value! ? "TRUE" : "FALSE")
        : (Value?.ToString() ?? string.Empty);

    /// <summary>
    /// Gets the boolean value. Returns false if not a boolean.
    /// </summary>
    public bool BoolValue => Type == FormulaResultType.Boolean && (bool)Value!;

    /// <summary>
    /// Gets a value indicating whether this is an error value.
    /// </summary>
    public bool IsError => Type == FormulaResultType.Error;

    /// <summary>
    /// Gets the error value string. Returns null if not an error.
    /// </summary>
    public string? ErrorValue => IsError ? (string?)Value : null;

    private FormulaResult(FormulaResultType type, object? value)
    {
        Type = type;
        Value = value;
    }

    /// <summary>
    /// Creates a numeric cell value.
    /// </summary>
    /// <param name="value">The numeric value.</param>
    /// <returns>A FormulaResult representing a number.</returns>
    public static FormulaResult FromNumber(double value) => new(FormulaResultType.Number, value);

    /// <summary>
    /// Creates a string cell value.
    /// </summary>
    /// <param name="value">The string value.</param>
    /// <returns>A FormulaResult representing a string.</returns>
    public static FormulaResult FromString(string value) => new(FormulaResultType.Text, value);

    /// <summary>
    /// Creates a boolean cell value.
    /// </summary>
    /// <param name="value">The boolean value.</param>
    /// <returns>A FormulaResult representing a boolean.</returns>
    public static FormulaResult FromBool(bool value) => new(FormulaResultType.Boolean, value);

    /// <summary>
    /// Creates an error cell value.
    /// </summary>
    /// <param name="error">The error string.</param>
    /// <returns>A FormulaResult representing an error.</returns>
    public static FormulaResult Error(string error) => new(FormulaResultType.Error, error);

    /// <summary>
    /// Gets an empty cell value.
    /// </summary>
    public static FormulaResult Empty => new(FormulaResultType.Empty, null);

    /// <inheritdoc/>
    public bool Equals(FormulaResult other) => Type == other.Type && Equals(Value, other.Value);

    /// <inheritdoc/>
    public override bool Equals(object? obj) => obj is FormulaResult other && Equals(other);

    /// <inheritdoc/>
    public override int GetHashCode()
    {
        unchecked
        {
            int hash = 17;
            hash = hash * 31 + Type.GetHashCode();
            hash = hash * 31 + (Value?.GetHashCode() ?? 0);
            return hash;
        }
    }

    /// <summary>
    /// Equality operator.
    /// </summary>
    public static bool operator ==(FormulaResult left, FormulaResult right) => left.Equals(right);

    /// <summary>
    /// Inequality operator.
    /// </summary>
    public static bool operator !=(FormulaResult left, FormulaResult right) => !left.Equals(right);
}

/// <summary>
/// Specifies the type of a formula evaluation result.
/// </summary>
public enum FormulaResultType
{
    /// <summary>
    /// Empty cell.
    /// </summary>
    Empty = 0,

    /// <summary>
    /// Numeric value.
    /// </summary>
    Number = 1,

    /// <summary>
    /// Text value.
    /// </summary>
    Text = 2,

    /// <summary>
    /// Boolean value.
    /// </summary>
    Boolean = 3,

    /// <summary>
    /// Error value.
    /// </summary>
    Error = 4,
}
