// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Text;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the CONCAT function.
/// CONCAT(text1, [text2], ...) - concatenates text strings (modern version of CONCATENATE).
/// </summary>
public sealed class ConcatFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ConcatFunction Instance = new();

    private ConcatFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "CONCAT";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        var result = new StringBuilder();

        foreach (var arg in args)
        {
            if (arg.IsError)
            {
                return arg; // Propagate errors
            }

            result.Append(arg.StringValue);
        }

        return CellValue.FromString(result.ToString());
    }
}
