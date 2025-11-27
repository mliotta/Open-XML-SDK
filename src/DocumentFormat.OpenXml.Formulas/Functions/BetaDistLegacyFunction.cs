// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the BETADIST function (legacy Excel 2007 compatibility).
/// BETADIST(x, alpha, beta, [A], [B]) - returns the beta cumulative distribution function.
/// This is the legacy version; modern Excel uses BETA.DIST.
/// </summary>
public sealed class BetaDistLegacyFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly BetaDistLegacyFunction Instance = new();

    private BetaDistLegacyFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "BETADIST";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length < 3 || args.Length > 5)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // BETADIST always uses cumulative=TRUE (the legacy function only returned CDF)
        // Create new args array with cumulative flag added
        var newArgs = new FormulaResult[args.Length + 1];
        for (int i = 0; i < args.Length; i++)
        {
            newArgs[i] = args[i];
        }
        newArgs[args.Length] = FormulaResult.FromBool(true);

        // Delegate to BETA.DIST
        return BetaDistFunction.Instance.Execute(context, newArgs);
    }
}
