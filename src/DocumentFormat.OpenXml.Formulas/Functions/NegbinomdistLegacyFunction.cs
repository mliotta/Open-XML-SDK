// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the NEGBINOMDIST function (legacy compatibility function).
/// NEGBINOMDIST(number_f, number_s, probability_s) - returns the negative binomial distribution.
/// This is a legacy function that delegates to NEGBINOM.DIST with cumulative=FALSE.
/// </summary>
public sealed class NegbinomdistLegacyFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly NegbinomdistLegacyFunction Instance = new();

    private NegbinomdistLegacyFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "NEGBINOMDIST";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 3)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Legacy NEGBINOMDIST always returns PMF (cumulative=FALSE)
        var argsWithCumulative = new FormulaResult[4];
        argsWithCumulative[0] = args[0];
        argsWithCumulative[1] = args[1];
        argsWithCumulative[2] = args[2];
        argsWithCumulative[3] = FormulaResult.FromBool(false);

        return NegBinomDistFunction.Instance.Execute(context, argsWithCumulative);
    }
}
