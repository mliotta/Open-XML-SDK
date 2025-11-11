// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Interface for function implementations.
/// </summary>
public interface IFunctionImplementation
{
    /// <summary>
    /// Gets the function name.
    /// </summary>
    string Name { get; }

    /// <summary>
    /// Executes the function.
    /// </summary>
    /// <param name="context">The cell context.</param>
    /// <param name="args">The function arguments.</param>
    /// <returns>The function result.</returns>
    CellValue Execute(CellContext context, CellValue[] args);
}
