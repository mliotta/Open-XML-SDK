using System;
using System.IO;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests;

var filePath = Path.Combine(Path.GetTempPath(), "FormulaOracle.xlsx");
Console.WriteLine($"Generating oracle test file at: {filePath}");

OracleTestFileGenerator.GenerateOracleTestFile(filePath);

Console.WriteLine($"✓ Oracle test file generated successfully!");
Console.WriteLine();
Console.WriteLine("NEXT STEPS:");
Console.WriteLine("1. Open this file in Excel");
Console.WriteLine("2. Excel will calculate all formulas and store cached values");
Console.WriteLine("3. Save and close the file");
Console.WriteLine("4. Copy to: test/DocumentFormat.OpenXml.Formulas.Tests/TestFiles/FormulaOracle.xlsx");
