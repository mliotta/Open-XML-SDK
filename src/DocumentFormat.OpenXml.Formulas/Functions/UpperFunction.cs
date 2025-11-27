// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the UPPER function.
/// UPPER(text) - converts to uppercase.
/// </summary>
public sealed class UpperFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly UpperFunction Instance = new();

    private UpperFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "UPPER";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 1)
        {
            return FormulaResult.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        return FormulaResult.FromString(args[0].StringValue.ToUpperInvariant());
    }
}
