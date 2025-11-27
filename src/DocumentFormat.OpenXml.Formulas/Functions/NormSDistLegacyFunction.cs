// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the NORMSDIST function (legacy Excel 2007 compatibility).
/// NORMSDIST(z) - returns the standard normal cumulative distribution function.
/// This is the legacy version; modern Excel uses NORM.S.DIST.
/// </summary>
public sealed class NormSDistLegacyFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly NormSDistLegacyFunction Instance = new();

    private NormSDistLegacyFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "NORMSDIST";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 1)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // NORMSDIST always uses cumulative=TRUE (the legacy function only returned CDF)
        // Create new args array with cumulative flag added
        var newArgs = new FormulaResult[2];
        newArgs[0] = args[0];
        newArgs[1] = FormulaResult.FromBool(true);

        // Delegate to NORM.S.DIST
        return NormSDistFunction.Instance.Execute(context, newArgs);
    }
}
