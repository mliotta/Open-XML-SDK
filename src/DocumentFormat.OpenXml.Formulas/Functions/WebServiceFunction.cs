// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the WEBSERVICE function.
/// WEBSERVICE(url) - Gets data from a web service.
/// NOTE: This function returns #VALUE! for security reasons. Web service calls are disabled
/// to prevent potential security vulnerabilities when evaluating formulas from untrusted sources.
/// </summary>
public sealed class WebServiceFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly WebServiceFunction Instance = new();

    private WebServiceFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "WEBSERVICE";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 1)
        {
            return CellValue.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        // For security reasons, WEBSERVICE is not supported.
        // Making HTTP requests from formula evaluation could pose security risks:
        // 1. SSRF (Server-Side Request Forgery) attacks
        // 2. Data exfiltration
        // 3. Unwanted external network access
        // 4. Timing-based side channel attacks
        return CellValue.Error("#VALUE!");
    }
}
