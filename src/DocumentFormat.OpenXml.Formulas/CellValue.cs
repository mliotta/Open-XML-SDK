// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation;

/// <summary>
/// Represents a cell value with its type.
/// </summary>
public readonly struct CellValue : IEquatable<CellValue>
{
    /// <summary>
    /// Gets the type of the cell value.
    /// </summary>
    public CellValueType Type { get; }

    /// <summary>
    /// Gets the raw value object.
    /// </summary>
    public object? Value { get; }

    /// <summary>
    /// Gets the numeric value. Returns 0 if not a number.
    /// </summary>
    public double NumericValue => Type == CellValueType.Number ? (double)Value! : 0;

    /// <summary>
    /// Gets the string value.
    /// </summary>
    public string StringValue => Value?.ToString() ?? string.Empty;

    /// <summary>
    /// Gets the boolean value. Returns false if not a boolean.
    /// </summary>
    public bool BoolValue => Type == CellValueType.Boolean && (bool)Value!;

    /// <summary>
    /// Gets a value indicating whether this is an error value.
    /// </summary>
    public bool IsError => Type == CellValueType.Error;

    /// <summary>
    /// Gets the error value string. Returns null if not an error.
    /// </summary>
    public string? ErrorValue => IsError ? (string?)Value : null;

    private CellValue(CellValueType type, object? value)
    {
        Type = type;
        Value = value;
    }

    /// <summary>
    /// Creates a numeric cell value.
    /// </summary>
    /// <param name="value">The numeric value.</param>
    /// <returns>A CellValue representing a number.</returns>
    public static CellValue FromNumber(double value) => new(CellValueType.Number, value);

    /// <summary>
    /// Creates a string cell value.
    /// </summary>
    /// <param name="value">The string value.</param>
    /// <returns>A CellValue representing a string.</returns>
    public static CellValue FromString(string value) => new(CellValueType.Text, value);

    /// <summary>
    /// Creates a boolean cell value.
    /// </summary>
    /// <param name="value">The boolean value.</param>
    /// <returns>A CellValue representing a boolean.</returns>
    public static CellValue FromBool(bool value) => new(CellValueType.Boolean, value);

    /// <summary>
    /// Creates an error cell value.
    /// </summary>
    /// <param name="error">The error string.</param>
    /// <returns>A CellValue representing an error.</returns>
    public static CellValue Error(string error) => new(CellValueType.Error, error);

    /// <summary>
    /// Gets an empty cell value.
    /// </summary>
    public static CellValue Empty => new(CellValueType.Empty, null);

    /// <inheritdoc/>
    public bool Equals(CellValue other) => Type == other.Type && Equals(Value, other.Value);

    /// <inheritdoc/>
    public override bool Equals(object? obj) => obj is CellValue other && Equals(other);

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
    public static bool operator ==(CellValue left, CellValue right) => left.Equals(right);

    /// <summary>
    /// Inequality operator.
    /// </summary>
    public static bool operator !=(CellValue left, CellValue right) => !left.Equals(right);
}

/// <summary>
/// Specifies the type of a cell value.
/// </summary>
public enum CellValueType
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
