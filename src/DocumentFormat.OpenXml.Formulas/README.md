# DocumentFormat.OpenXml.Formulas

Formula evaluation and pivot-table generation layered on top of the core Open XML SDK.

This package is intended for scenarios where you need spreadsheet values computed without opening the file in Excel — for example, server-side generation, validation, and automated reporting.

## Scenarios

- Evaluate Excel formulas directly against a worksheet, including a large library of built-in functions (math, statistical, logical, text, lookup, date/time, financial, and engineering).
- Recalculate a sheet or just the dependents of changed cells, using a dependency graph with circular-reference detection.
- Build native OOXML pivot tables whose results are precomputed and written into the worksheet, so the values are visible without recalculation. Supports multiple row, column, and value fields, report filters, grand totals, layout options, and "show value as" percentages.

## Documentation and feedback

- Official SDK docs: https://learn.microsoft.com/en-us/office/open-xml/open-xml-sdk
- Source and feature examples: https://github.com/dotnet/Open-XML-SDK
- Issues and ideas: https://github.com/dotnet/Open-XML-SDK/issues
