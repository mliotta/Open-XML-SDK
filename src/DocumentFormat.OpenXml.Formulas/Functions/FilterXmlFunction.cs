// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Xml;
using System.Xml.XPath;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the FILTERXML function.
/// FILTERXML(xml, xpath) - Extracts data from XML using XPath.
/// NOTE: This is a simplified implementation. Complex XPath expressions may not be fully supported.
/// </summary>
public sealed class FilterXmlFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly FilterXmlFunction Instance = new();

    private FilterXmlFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "FILTERXML";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 2)
        {
            return CellValue.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[1].IsError)
        {
            return args[1];
        }

        var xml = args[0].StringValue;
        var xpath = args[1].StringValue;

        if (string.IsNullOrEmpty(xml) || string.IsNullOrEmpty(xpath))
        {
            return CellValue.Error("#VALUE!");
        }

        try
        {
            var doc = new XmlDocument();
            doc.LoadXml(xml);

            var navigator = doc.CreateNavigator();
            var result = navigator.SelectSingleNode(xpath);

            if (result == null)
            {
                // No matching node found
                return CellValue.Error("#N/A");
            }

            var value = result.Value;

            // Try to parse as number
            if (double.TryParse(value, out var numValue))
            {
                return CellValue.FromNumber(numValue);
            }

            // Return as text
            return CellValue.FromString(value);
        }
        catch (XmlException)
        {
            return CellValue.Error("#VALUE!");
        }
        catch (XPathException)
        {
            return CellValue.Error("#VALUE!");
        }
        catch (Exception)
        {
            return CellValue.Error("#VALUE!");
        }
    }
}
