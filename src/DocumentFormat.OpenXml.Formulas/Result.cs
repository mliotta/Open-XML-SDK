// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation;

/// <summary>
/// Represents the result of a formula evaluation operation.
/// </summary>
/// <typeparam name="T">The type of the success value.</typeparam>
public readonly struct Result<T>
{
    /// <summary>
    /// Gets a value indicating whether the operation was successful.
    /// </summary>
    public bool IsSuccess { get; }

    /// <summary>
    /// Gets the value if the operation was successful.
    /// </summary>
    public T Value { get; }

    /// <summary>
    /// Gets the error if the operation failed.
    /// </summary>
    public EvaluationError? Error { get; }

    private Result(bool success, T value, EvaluationError? error)
    {
        IsSuccess = success;
        Value = value;
        Error = error;
    }

    /// <summary>
    /// Creates a successful result.
    /// </summary>
    /// <param name="value">The success value.</param>
    /// <returns>A successful Result.</returns>
    public static Result<T> Success(T value) => new(true, value, null);

    /// <summary>
    /// Creates a failed result.
    /// </summary>
    /// <param name="error">The error.</param>
    /// <returns>A failed Result.</returns>
    public static Result<T> Failure(EvaluationError error) => new(false, default!, error);

    /// <summary>
    /// Matches the result to one of two functions.
    /// </summary>
    /// <typeparam name="TResult">The return type.</typeparam>
    /// <param name="onSuccess">Function to call on success.</param>
    /// <param name="onFailure">Function to call on failure.</param>
    /// <returns>The result of the matching function.</returns>
    public TResult Match<TResult>(
        Func<T, TResult> onSuccess,
        Func<EvaluationError, TResult> onFailure)
    {
        return IsSuccess ? onSuccess(Value) : onFailure(Error!);
    }
}

/// <summary>
/// Base class for evaluation errors.
/// </summary>
public abstract class EvaluationError : Exception
{
    /// <summary>
    /// Gets the cell reference where the error occurred, if applicable.
    /// </summary>
    public string? CellReference { get; }

    /// <summary>
    /// Initializes a new instance of the <see cref="EvaluationError"/> class.
    /// </summary>
    /// <param name="message">The error message.</param>
    /// <param name="cellReference">Optional cell reference.</param>
    protected EvaluationError(string message, string? cellReference = null)
        : base(message)
    {
        CellReference = cellReference;
    }
}

/// <summary>
/// Represents a parser exception error.
/// </summary>
public class ParserException : EvaluationError
{
    /// <summary>
    /// Initializes a new instance of the <see cref="ParserException"/> class.
    /// </summary>
    /// <param name="message">The error message.</param>
    public ParserException(string message)
        : base(message)
    {
    }
}

/// <summary>
/// Represents a compilation exception error.
/// </summary>
public class CompilationException : EvaluationError
{
    /// <summary>
    /// Initializes a new instance of the <see cref="CompilationException"/> class.
    /// </summary>
    /// <param name="message">The error message.</param>
    public CompilationException(string message)
        : base(message)
    {
    }
}

/// <summary>
/// Represents an unsupported function error.
/// </summary>
public class UnsupportedFunctionException : EvaluationError
{
    /// <summary>
    /// Gets the function name that is not supported.
    /// </summary>
    public string FunctionName { get; }

    /// <summary>
    /// Initializes a new instance of the <see cref="UnsupportedFunctionException"/> class.
    /// </summary>
    /// <param name="functionName">The unsupported function name.</param>
    public UnsupportedFunctionException(string functionName)
        : base($"Function '{functionName}' is not supported")
    {
        FunctionName = functionName;
    }
}

/// <summary>
/// Exception thrown when a circular reference is detected.
/// </summary>
public class CircularReferenceException : EvaluationError
{
    /// <summary>
    /// The chain of cell references forming the cycle.
    /// </summary>
    public List<string> CellChain { get; }

    /// <summary>
    /// Initializes a new instance of the CircularReferenceException class.
    /// </summary>
    public CircularReferenceException(List<string> cellChain)
        : base($"Circular reference detected: {string.Join(" → ", cellChain.ToArray())}")
    {
        CellChain = cellChain;
    }
}

/// <summary>
/// Exception thrown when an invalid cell reference is encountered.
/// </summary>
public class InvalidReferenceException : EvaluationError
{
    /// <summary>
    /// The invalid reference.
    /// </summary>
    public string Reference { get; }

    /// <summary>
    /// Initializes a new instance of the InvalidReferenceException class.
    /// </summary>
    public InvalidReferenceException(string reference)
        : base($"Invalid cell reference: {reference}")
    {
        Reference = reference;
    }
}
