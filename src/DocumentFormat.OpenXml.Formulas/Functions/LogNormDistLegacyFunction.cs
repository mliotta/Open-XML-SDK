// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the LOGNORMDIST function (legacy Excel 2007 compatibility).
/// LOGNORMDIST(x, mean, standard_dev) - returns the lognormal cumulative distribution function.
/// This is the legacy version; modern Excel uses LOGNORM.DIST.
/// </summary>
public sealed class LogNormDistLegacyFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly LogNormDistLegacyFunction Instance = new();

    private LogNormDistLegacyFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "LOGNORMDIST";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 3)
        {
            return CellValue.Error("#VALUE!");
        }

        // LOGNORMDIST always uses cumulative=TRUE (the legacy function only returned CDF)
        // Create new args array with cumulative flag added
        var newArgs = new CellValue[4];
        newArgs[0] = args[0];
        newArgs[1] = args[1];
        newArgs[2] = args[2];
        newArgs[3] = CellValue.FromBool(true);

        // Delegate to LOGNORM.DIST
        return LogNormDistFunction.Instance.Execute(context, newArgs);
    }
}
