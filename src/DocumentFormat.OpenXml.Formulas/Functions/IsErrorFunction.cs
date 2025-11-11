// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ISERROR function.
/// ISERROR(value) - TRUE if value is any error.
/// </summary>
public sealed class IsErrorFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly IsErrorFunction Instance = new();

    private IsErrorFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ISERROR";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 1)
        {
            return CellValue.Error("#VALUE!");
        }

        // Note: Errors are NOT propagated for IS* functions
        var isError = args[0].IsError;
        return CellValue.FromBool(isError);
    }
}
