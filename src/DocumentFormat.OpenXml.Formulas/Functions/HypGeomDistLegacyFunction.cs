// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the HYPGEOMDIST function (legacy Excel 2007 compatibility).
/// HYPGEOMDIST(sample_s, number_sample, population_s, number_pop) - returns the hypergeometric distribution.
/// This is the legacy version; modern Excel uses HYPGEOM.DIST.
/// </summary>
public sealed class HypGeomDistLegacyFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly HypGeomDistLegacyFunction Instance = new();

    private HypGeomDistLegacyFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "HYPGEOMDIST";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 4)
        {
            return CellValue.Error("#VALUE!");
        }

        // HYPGEOMDIST always uses cumulative=FALSE (the legacy function only returned PMF)
        // Create new args array with cumulative flag added
        var newArgs = new CellValue[5];
        newArgs[0] = args[0];
        newArgs[1] = args[1];
        newArgs[2] = args[2];
        newArgs[3] = args[3];
        newArgs[4] = CellValue.FromBool(false);

        // Delegate to HYPGEOM.DIST
        return HypGeomDistFunction.Instance.Execute(context, newArgs);
    }
}
