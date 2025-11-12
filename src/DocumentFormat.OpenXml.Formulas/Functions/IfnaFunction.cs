// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the IFNA function.
/// IFNA(value, value_if_na) - Returns value_if_na if value is #N/A error, otherwise returns value.
/// </summary>
public sealed class IfnaFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly IfnaFunction Instance = new();

    private IfnaFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "IFNA";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 2)
        {
            return CellValue.Error("#VALUE!");
        }

        // If the first argument is specifically #N/A error, return the second argument
        if (args[0].IsError && args[0].ErrorValue == "#N/A")
        {
            return args[1];
        }

        // Otherwise, return the first argument (including other error types)
        return args[0];
    }
}
