// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features;
using DocumentFormat.OpenXml.Packaging;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation;

/// <summary>
/// Extension methods for adding formula evaluation feature to documents.
/// </summary>
public static class FormulaEvaluationExtensions
{
    /// <summary>
    /// Adds formula evaluation feature to a spreadsheet document.
    /// </summary>
    /// <param name="document">The spreadsheet document.</param>
    public static void AddFormulaEvaluationFeature(this SpreadsheetDocument document)
    {
        var evaluator = new FormulaEvaluator(document);
        document.Features.Set<IFormulaEvaluator>(evaluator);
    }

    /// <summary>
    /// Gets the formula evaluator feature from a spreadsheet document.
    /// </summary>
    /// <param name="document">The spreadsheet document.</param>
    /// <returns>The formula evaluator, or null if not added.</returns>
    public static IFormulaEvaluator? GetFormulaEvaluator(this SpreadsheetDocument document)
    {
        return document.Features.Get<IFormulaEvaluator>();
    }
}
