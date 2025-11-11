// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the IFERROR function.
/// IFERROR(value, value_if_error) - Returns value_if_error if value is an error.
/// </summary>
public sealed class IFErrorFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly IFErrorFunction Instance = new();

    private IFErrorFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "IFERROR";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 2)
        {
            return CellValue.Error("#VALUE!");
        }

        // If the first argument is an error, return the second argument
        if (args[0].IsError)
        {
            return args[1];
        }

        // Otherwise, return the first argument
        return args[0];
    }
}
