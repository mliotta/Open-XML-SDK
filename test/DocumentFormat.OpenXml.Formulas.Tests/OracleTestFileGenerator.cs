// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Globalization;
using System.IO;
using System.Linq;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests;

/// <summary>
/// Generates Excel files with comprehensive formula test cases for oracle validation.
/// </summary>
public static class OracleTestFileGenerator
{
    /// <summary>
    /// Generates an Excel file with comprehensive formula test cases.
    /// User should open this file in Excel, which will calculate and cache all formula results.
    /// </summary>
    public static void GenerateOracleTestFile(string filePath)
    {
        using var doc = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
        var workbookPart = doc.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var sheets = workbookPart.Workbook.AppendChild(new Sheets());

        // Create multiple sheets for different function categories
        CreateMathFunctionsSheet(workbookPart, sheets, 1);
        // CreateLogicalFunctionsSheet(workbookPart, sheets, 2); // TODO: Re-enable after fixing
        // CreateTextFunctionsSheet(workbookPart, sheets, 3); // BUG: Test data in header row (F1, G1, H1)
        CreateLookupFunctionsSheet(workbookPart, sheets, 4);
        CreateDateTimeFunctionsSheet(workbookPart, sheets, 5);
        // CreateStatisticalFunctionsSheet(workbookPart, sheets, 6); // BUG: Similar to Text sheet
        // CreateFinancialFunctionsSheet(workbookPart, sheets, 7); // TODO: Re-enable after fixing
        CreateEngineeringFunctionsSheet(workbookPart, sheets, 8);
        CreateDatabaseFunctionsSheet(workbookPart, sheets, 9);
        CreateInformationFunctionsSheet(workbookPart, sheets, 10);
        // CreateErrorHandlingFunctionsSheet(workbookPart, sheets, 11); // TODO: Re-enable after fixing
        CreateForecastingFunctionsSheet(workbookPart, sheets, 12);
        CreateCubeFunctionsSheet(workbookPart, sheets, 13);


        workbookPart.Workbook.Save();
    }

    private static void CreateMathFunctionsSheet(WorkbookPart workbookPart, Sheets sheets, uint sheetId)
    {
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        worksheetPart.Worksheet = new Worksheet(sheetData);

        sheets.AppendChild(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = "Math",
        });

        AddHeader(sheetData, 1);
        var row = 2;

        // SUM tests
        AddTestCase(sheetData, row++, "SUM", "=SUM(1,2,3)", "Basic sum");
        AddTestCase(sheetData, row++, "SUM", "=SUM(F3:F5)", "Range sum", ("F3", "10"), ("G3", "20"), ("H3", "30"));
        AddTestCase(sheetData, row++, "SUM", "=SUM(1,2,F4)", "Mixed sum", ("F4", "5"));

        // AVERAGE tests
        AddTestCase(sheetData, row++, "AVERAGE", "=AVERAGE(10,20,30)", "Basic average");
        AddTestCase(sheetData, row++, "AVERAGE", "=AVERAGE(F6:H6)", "Range average", ("F6", "5"), ("G6", "15"), ("H6", "25"));

        // COUNT tests
        AddTestCase(sheetData, row++, "COUNT", "=COUNT(1,2,3,4,5)", "Basic count");
        AddTestCase(sheetData, row++, "COUNT", "=COUNT(F8:J8)", "Range count", ("F8", "1"), ("G8", "2"), ("H8", "3"), ("I8", "4"), ("J8", "5"));

        // COUNTA tests
        AddTestCase(sheetData, row++, "COUNTA", "=COUNTA(1,\"text\",TRUE)", "Count with mixed types");

        // MAX/MIN tests
        AddTestCase(sheetData, row++, "MAX", "=MAX(5,10,3,8)", "Basic max");
        AddTestCase(sheetData, row++, "MIN", "=MIN(5,10,3,8)", "Basic min");
        AddTestCase(sheetData, row++, "MAX", "=MAX(F12:I12)", "Range max", ("F12", "15"), ("G12", "25"), ("H12", "10"), ("I12", "20"));

        // ROUND tests
        AddTestCase(sheetData, row++, "ROUND", "=ROUND(3.14159, 2)", "Round to 2 decimals");
        AddTestCase(sheetData, row++, "ROUND", "=ROUND(2.5, 0)", "Round half up");
        AddTestCase(sheetData, row++, "ROUNDUP", "=ROUNDUP(3.14, 1)", "Round up");
        AddTestCase(sheetData, row++, "ROUNDDOWN", "=ROUNDDOWN(3.99, 1)", "Round down");

        // ABS tests
        AddTestCase(sheetData, row++, "ABS", "=ABS(-10)", "Absolute value negative");
        AddTestCase(sheetData, row++, "ABS", "=ABS(10)", "Absolute value positive");

        // PRODUCT tests
        AddTestCase(sheetData, row++, "PRODUCT", "=PRODUCT(2,3,4)", "Basic product");
        AddTestCase(sheetData, row++, "PRODUCT", "=PRODUCT(F20:H20)", "Range product", ("F20", "2"), ("G20", "5"), ("H20", "10"));

        // POWER tests
        AddTestCase(sheetData, row++, "POWER", "=POWER(2,3)", "2 to the power of 3");
        AddTestCase(sheetData, row++, "POWER", "=POWER(5,2)", "5 squared");

        // SQRT tests
        AddTestCase(sheetData, row++, "SQRT", "=SQRT(16)", "Square root of 16");
        AddTestCase(sheetData, row++, "SQRT", "=SQRT(2)", "Square root of 2");

        // MOD tests
        AddTestCase(sheetData, row++, "MOD", "=MOD(10,3)", "10 modulo 3");
        AddTestCase(sheetData, row++, "MOD", "=MOD(7,2)", "7 modulo 2");

        // INT tests
        AddTestCase(sheetData, row++, "INT", "=INT(8.9)", "Integer part of 8.9");
        AddTestCase(sheetData, row++, "INT", "=INT(-2.5)", "Integer part of -2.5");

        // CEILING tests
        AddTestCase(sheetData, row++, "CEILING", "=CEILING(2.5, 1)", "Ceiling 2.5 to 1");
        AddTestCase(sheetData, row++, "CEILING", "=CEILING(10.3, 0.5)", "Ceiling 10.3 to 0.5");

        // FLOOR tests
        AddTestCase(sheetData, row++, "FLOOR", "=FLOOR(2.5, 1)", "Floor 2.5 to 1");
        AddTestCase(sheetData, row++, "FLOOR", "=FLOOR(10.7, 0.5)", "Floor 10.7 to 0.5");

        // TRUNC tests
        AddTestCase(sheetData, row++, "TRUNC", "=TRUNC(8.9)", "Truncate 8.9");
        AddTestCase(sheetData, row++, "TRUNC", "=TRUNC(8.9, 1)", "Truncate 8.9 to 1 decimal");

        // SIGN tests
        AddTestCase(sheetData, row++, "SIGN", "=SIGN(10)", "Sign of positive");
        AddTestCase(sheetData, row++, "SIGN", "=SIGN(-5)", "Sign of negative");
        AddTestCase(sheetData, row++, "SIGN", "=SIGN(0)", "Sign of zero");

        // EXP tests
        AddTestCase(sheetData, row++, "EXP", "=EXP(1)", "e to the power of 1");
        AddTestCase(sheetData, row++, "EXP", "=EXP(2)", "e to the power of 2");

        // LN tests
        AddTestCase(sheetData, row++, "LN", "=LN(EXP(1))", "Natural log of e");
        AddTestCase(sheetData, row++, "LN", "=LN(10)", "Natural log of 10");

        // LOG tests
        AddTestCase(sheetData, row++, "LOG", "=LOG(100, 10)", "Log base 10 of 100");
        AddTestCase(sheetData, row++, "LOG", "=LOG(8, 2)", "Log base 2 of 8");

        // LOG10 tests
        AddTestCase(sheetData, row++, "LOG10", "=LOG10(100)", "Log10 of 100");
        AddTestCase(sheetData, row++, "LOG10", "=LOG10(1000)", "Log10 of 1000");

        // PI test
        AddTestCase(sheetData, row++, "PI", "=PI()", "Pi constant");

        // RADIANS tests
        AddTestCase(sheetData, row++, "RADIANS", "=RADIANS(180)", "180 degrees to radians");
        AddTestCase(sheetData, row++, "RADIANS", "=RADIANS(90)", "90 degrees to radians");

        // DEGREES tests
        AddTestCase(sheetData, row++, "DEGREES", "=DEGREES(PI())", "Pi radians to degrees");
        AddTestCase(sheetData, row++, "DEGREES", "=DEGREES(1)", "1 radian to degrees");

        // SIN tests
        AddTestCase(sheetData, row++, "SIN", "=SIN(0)", "Sin of 0");
        AddTestCase(sheetData, row++, "SIN", "=SIN(PI()/2)", "Sin of pi/2");

        // COS tests
        AddTestCase(sheetData, row++, "COS", "=COS(0)", "Cos of 0");
        AddTestCase(sheetData, row++, "COS", "=COS(PI())", "Cos of pi");

        // TAN tests
        AddTestCase(sheetData, row++, "TAN", "=TAN(0)", "Tan of 0");
        AddTestCase(sheetData, row++, "TAN", "=TAN(PI()/4)", "Tan of pi/4");

        // ASIN tests
        AddTestCase(sheetData, row++, "ASIN", "=ASIN(0)", "Asin of 0");
        AddTestCase(sheetData, row++, "ASIN", "=ASIN(0.5)", "Asin of 0.5");

        // ACOS tests
        AddTestCase(sheetData, row++, "ACOS", "=ACOS(1)", "Acos of 1");
        AddTestCase(sheetData, row++, "ACOS", "=ACOS(0.5)", "Acos of 0.5");

        // ATAN tests
        AddTestCase(sheetData, row++, "ATAN", "=ATAN(0)", "Atan of 0");
        AddTestCase(sheetData, row++, "ATAN", "=ATAN(1)", "Atan of 1");

        // ATAN2 tests
        AddTestCase(sheetData, row++, "ATAN2", "=ATAN2(1, 1)", "Atan2 of (1,1)");
        AddTestCase(sheetData, row++, "ATAN2", "=ATAN2(1, 0)", "Atan2 of (1,0)");

        // SINH tests
        AddTestCase(sheetData, row++, "SINH", "=SINH(0)", "Sinh of 0");
        AddTestCase(sheetData, row++, "SINH", "=SINH(1)", "Sinh of 1");

        // COSH tests
        AddTestCase(sheetData, row++, "COSH", "=COSH(0)", "Cosh of 0");
        AddTestCase(sheetData, row++, "COSH", "=COSH(1)", "Cosh of 1");

        // TANH tests
        AddTestCase(sheetData, row++, "TANH", "=TANH(0)", "Tanh of 0");
        AddTestCase(sheetData, row++, "TANH", "=TANH(1)", "Tanh of 1");

        // ASINH tests
        AddTestCase(sheetData, row++, "ASINH", "=ASINH(0)", "Asinh of 0");
        AddTestCase(sheetData, row++, "ASINH", "=ASINH(1)", "Asinh of 1");

        // ACOSH tests
        AddTestCase(sheetData, row++, "ACOSH", "=ACOSH(1)", "Acosh of 1");
        AddTestCase(sheetData, row++, "ACOSH", "=ACOSH(2)", "Acosh of 2");

        // ATANH tests
        AddTestCase(sheetData, row++, "ATANH", "=ATANH(0)", "Atanh of 0");
        AddTestCase(sheetData, row++, "ATANH", "=ATANH(0.5)", "Atanh of 0.5");

        // COMBIN tests
        AddTestCase(sheetData, row++, "COMBIN", "=COMBIN(5, 2)", "5 choose 2");
        AddTestCase(sheetData, row++, "COMBIN", "=COMBIN(10, 3)", "10 choose 3");

        // PERMUT tests
        AddTestCase(sheetData, row++, "PERMUT", "=PERMUT(5, 2)", "5 permute 2");
        AddTestCase(sheetData, row++, "PERMUT", "=PERMUT(10, 3)", "10 permute 3");

        // MROUND tests
        AddTestCase(sheetData, row++, "MROUND", "=MROUND(10, 3)", "Round 10 to nearest 3");
        AddTestCase(sheetData, row++, "MROUND", "=MROUND(7.5, 2)", "Round 7.5 to nearest 2");

        // QUOTIENT tests
        AddTestCase(sheetData, row++, "QUOTIENT", "=QUOTIENT(10, 3)", "Integer quotient 10/3");
        AddTestCase(sheetData, row++, "QUOTIENT", "=QUOTIENT(15, 4)", "Integer quotient 15/4");

        // SUMSQ tests
        AddTestCase(sheetData, row++, "SUMSQ", "=SUMSQ(3, 4)", "Sum of squares 3,4");
        AddTestCase(sheetData, row++, "SUMSQ", "=SUMSQ(1, 2, 3)", "Sum of squares 1,2,3");

        // SUMX2MY2 tests
        AddTestCase(sheetData, row++, "SUMX2MY2", "=SUMX2MY2(F60:F62, G60:G62)", "Sum x^2-y^2", ("F60", "3"), ("G60", "2"), ("F61", "4"), ("G61", "1"), ("F62", "5"), ("G62", "3"));

        // SUMX2PY2 tests
        AddTestCase(sheetData, row++, "SUMX2PY2", "=SUMX2PY2(F64:F66, G64:G66)", "Sum x^2+y^2", ("F64", "3"), ("G64", "2"), ("F65", "4"), ("G65", "1"), ("F66", "5"), ("G66", "3"));

        // SUMXMY2 tests
        AddTestCase(sheetData, row++, "SUMXMY2", "=SUMXMY2(F68:F70, G68:G70)", "Sum (x-y)^2", ("F68", "3"), ("G68", "2"), ("F69", "4"), ("G69", "1"), ("F70", "5"), ("G70", "3"));

        // MULTINOMIAL tests
        AddTestCase(sheetData, row++, "MULTINOMIAL", "=MULTINOMIAL(2, 3, 4)", "Multinomial 2,3,4");

        // SERIESSUM tests
        AddTestCase(sheetData, row++, "SERIESSUM", "=SERIESSUM(2, 0, 1, F73:F75)", "Series sum", ("F73", "1"), ("F74", "2"), ("F75", "3"));

        // SUMIFS tests
        AddTestCase(sheetData, row++, "SUMIFS", "=SUMIFS(F77:F79, G77:G79, \">5\")", "Sum with criteria", ("F77", "10"), ("G77", "6"), ("F78", "20"), ("G78", "4"), ("F79", "30"), ("G79", "8"));

        // COUNTIFS tests
        AddTestCase(sheetData, row++, "COUNTIFS", "=COUNTIFS(F81:F83, \">10\", G81:G83, \"<50\")", "Count multiple criteria", ("F81", "15"), ("G81", "40"), ("F82", "5"), ("G82", "30"), ("F83", "25"), ("G83", "60"));

        // SUMPRODUCT tests
        AddTestCase(sheetData, row++, "SUMPRODUCT", "=SUMPRODUCT(F85:F87, G85:G87)", "Sum product", ("F85", "2"), ("G85", "3"), ("F86", "4"), ("G86", "5"), ("F87", "6"), ("G87", "7"));

        // FACT tests
        AddTestCase(sheetData, row++, "FACT", "=FACT(5)", "Factorial of 5");
        AddTestCase(sheetData, row++, "FACT", "=FACT(0)", "Factorial of 0");

        // GCD tests
        AddTestCase(sheetData, row++, "GCD", "=GCD(12, 18)", "GCD of 12 and 18");
        AddTestCase(sheetData, row++, "GCD", "=GCD(24, 36, 48)", "GCD of 24, 36, 48");

        // LCM tests
        AddTestCase(sheetData, row++, "LCM", "=LCM(12, 18)", "LCM of 12 and 18");
        AddTestCase(sheetData, row++, "LCM", "=LCM(4, 6, 8)", "LCM of 4, 6, 8");

        // EVEN tests
        AddTestCase(sheetData, row++, "EVEN", "=EVEN(3)", "Round 3 to even");
        AddTestCase(sheetData, row++, "EVEN", "=EVEN(2.5)", "Round 2.5 to even");

        // ODD tests
        AddTestCase(sheetData, row++, "ODD", "=ODD(2)", "Round 2 to odd");
        AddTestCase(sheetData, row++, "ODD", "=ODD(3.5)", "Round 3.5 to odd");

        // RAND tests
        AddTestCase(sheetData, row++, "RAND", "=IF(AND(RAND()>=0, RAND()<=1), 1, 0)", "Rand in range check");

        // RANDBETWEEN tests
        AddTestCase(sheetData, row++, "RANDBETWEEN", "=IF(AND(RANDBETWEEN(1,10)>=1, RANDBETWEEN(1,10)<=10), 1, 0)", "Randbetween range check");
        // SQRTPI tests
        AddTestCase(sheetData, row++, "SQRTPI", "=SQRTPI(1)", "Square root of pi");
        AddTestCase(sheetData, row++, "SQRTPI", "=SQRTPI(2)", "Square root of 2*pi");

        // CEILING.MATH tests
        AddTestCase(sheetData, row++, "CEILING.MATH", "=CEILING.MATH(4.3)", "Ceiling.Math positive");
        AddTestCase(sheetData, row++, "CEILING.MATH", "=CEILING.MATH(-4.3)", "Ceiling.Math negative");
        AddTestCase(sheetData, row++, "CEILING.MATH", "=CEILING.MATH(4.3, 2)", "Ceiling.Math significance");

        // CEILING.PRECISE tests
        AddTestCase(sheetData, row++, "CEILING.PRECISE", "=CEILING.PRECISE(4.3, 1)", "Ceiling.Precise positive");
        AddTestCase(sheetData, row++, "CEILING.PRECISE", "=CEILING.PRECISE(-4.3, 1)", "Ceiling.Precise negative");

        // FLOOR.MATH tests
        AddTestCase(sheetData, row++, "FLOOR.MATH", "=FLOOR.MATH(4.8)", "Floor.Math positive");
        AddTestCase(sheetData, row++, "FLOOR.MATH", "=FLOOR.MATH(-4.8)", "Floor.Math negative");
        AddTestCase(sheetData, row++, "FLOOR.MATH", "=FLOOR.MATH(4.8, 2)", "Floor.Math significance");

        // FLOOR.PRECISE tests
        AddTestCase(sheetData, row++, "FLOOR.PRECISE", "=FLOOR.PRECISE(4.8, 1)", "Floor.Precise positive");
        AddTestCase(sheetData, row++, "FLOOR.PRECISE", "=FLOOR.PRECISE(-4.8, 1)", "Floor.Precise negative");

        // ISO.CEILING tests
        AddTestCase(sheetData, row++, "ISO.CEILING", "=ISO.CEILING(4.3)", "ISO.Ceiling default");
        AddTestCase(sheetData, row++, "ISO.CEILING", "=ISO.CEILING(-4.3)", "ISO.Ceiling negative");
        AddTestCase(sheetData, row++, "ISO.CEILING", "=ISO.CEILING(4.3, 2)", "ISO.Ceiling significance");

        // FACTDOUBLE tests
        AddTestCase(sheetData, row++, "FACTDOUBLE", "=FACTDOUBLE(6)", "Double factorial of 6");
        AddTestCase(sheetData, row++, "FACTDOUBLE", "=FACTDOUBLE(7)", "Double factorial of 7");

        // COMBINA tests
        AddTestCase(sheetData, row++, "COMBINA", "=COMBINA(4, 3)", "Combinations with repetitions 4,3");
        AddTestCase(sheetData, row++, "COMBINA", "=COMBINA(10, 3)", "Combinations with repetitions 10,3");

        // PERMUTATIONA tests
        AddTestCase(sheetData, row++, "PERMUTATIONA", "=PERMUTATIONA(3, 2)", "Permutations with repetitions 3,2");
        AddTestCase(sheetData, row++, "PERMUTATIONA", "=PERMUTATIONA(5, 3)", "Permutations with repetitions 5,3");

        // MDETERM tests (Matrix determinant)
        AddTestCase(sheetData, row++, "MDETERM", "=MDETERM(F120:G121)", "Matrix determinant 2x2", ("F120", "1"), ("G120", "2"), ("F121", "3"), ("G121", "4"));
        AddTestCase(sheetData, row++, "MDETERM", "=MDETERM(F123:H125)", "Matrix determinant 3x3", ("F123", "1"), ("G123", "2"), ("H123", "3"), ("F124", "0"), ("G124", "1"), ("H124", "4"), ("F125", "5"), ("G125", "6"), ("H125", "0"));

        // MINVERSE tests (Matrix inverse)
        AddTestCase(sheetData, row++, "MINVERSE", "=INDEX(MINVERSE(F127:G128), 1, 1)", "Matrix inverse element 1,1", ("F127", "4"), ("G127", "7"), ("F128", "2"), ("G128", "6"));

        // MMULT tests (Matrix multiplication)
        AddTestCase(sheetData, row++, "MMULT", "=INDEX(MMULT(F130:G131, I130:J131), 1, 1)", "Matrix multiply element 1,1", ("F130", "1"), ("G130", "2"), ("F131", "3"), ("G131", "4"), ("I130", "2"), ("J130", "0"), ("I131", "1"), ("J131", "2"));

        // MUNIT tests (Unit matrix)
        AddTestCase(sheetData, row++, "MUNIT", "=INDEX(MUNIT(3), 1, 1)", "Unit matrix 3x3 element 1,1");
        AddTestCase(sheetData, row++, "MUNIT", "=INDEX(MUNIT(3), 2, 2)", "Unit matrix 3x3 element 2,2");
        AddTestCase(sheetData, row++, "MUNIT", "=INDEX(MUNIT(3), 1, 2)", "Unit matrix 3x3 element 1,2");

        // LOOKUP tests
        AddTestCase(sheetData, row++, "LOOKUP", "=LOOKUP(3, F135:F138, G135:G138)", "Lookup with result vector", ("F135", "1"), ("G135", "A"), ("F136", "2"), ("G136", "B"), ("F137", "3"), ("G137", "C"), ("F138", "4"), ("G138", "D"));
        AddTestCase(sheetData, row++, "LOOKUP", "=LOOKUP(5, F135:F138)", "Lookup single vector");

        // ACOT tests
        AddTestCase(sheetData, row++, "ACOT", "=ACOT(1)", "Arc cotangent of 1");
        AddTestCase(sheetData, row++, "ACOT", "=ACOT(2)", "Arc cotangent of 2");

        // ACOTH tests
        AddTestCase(sheetData, row++, "ACOTH", "=ACOTH(2)", "Arc hyperbolic cotangent of 2");
        AddTestCase(sheetData, row++, "ACOTH", "=ACOTH(3)", "Arc hyperbolic cotangent of 3");

        // CSC tests
        AddTestCase(sheetData, row++, "CSC", "=CSC(PI()/2)", "Cosecant of pi/2");
        AddTestCase(sheetData, row++, "CSC", "=CSC(PI()/4)", "Cosecant of pi/4");

        // CSCH tests
        AddTestCase(sheetData, row++, "CSCH", "=CSCH(1)", "Hyperbolic cosecant of 1");
        AddTestCase(sheetData, row++, "CSCH", "=CSCH(2)", "Hyperbolic cosecant of 2");

        // SEC tests
        AddTestCase(sheetData, row++, "SEC", "=SEC(0)", "Secant of 0");
        AddTestCase(sheetData, row++, "SEC", "=SEC(PI()/3)", "Secant of pi/3");

        // SECH tests
        AddTestCase(sheetData, row++, "SECH", "=SECH(0)", "Hyperbolic secant of 0");
        AddTestCase(sheetData, row++, "SECH", "=SECH(1)", "Hyperbolic secant of 1");

        // COT tests
        AddTestCase(sheetData, row++, "COT", "=COT(PI()/4)", "Cotangent of pi/4");
        AddTestCase(sheetData, row++, "COT", "=COT(PI()/3)", "Cotangent of pi/3");

        // COTH tests
        AddTestCase(sheetData, row++, "COTH", "=COTH(1)", "Hyperbolic cotangent of 1");
        AddTestCase(sheetData, row++, "COTH", "=COTH(2)", "Hyperbolic cotangent of 2");


        SortCellsInRows(sheetData);
        worksheetPart.Worksheet.Save();
    }

    private static void CreateLogicalFunctionsSheet(WorkbookPart workbookPart, Sheets sheets, uint sheetId)
    {
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        worksheetPart.Worksheet = new Worksheet(sheetData);

        sheets.AppendChild(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = "Logical",
        });

        AddHeader(sheetData, 1);
        var row = 2;

        // IF tests
        AddTestCase(sheetData, row++, "IF", "=IF(TRUE, \"yes\", \"no\")", "IF true");
        AddTestCase(sheetData, row++, "IF", "=IF(FALSE, \"yes\", \"no\")", "IF false");
        AddTestCase(sheetData, row++, "IF", "=IF(10>5, \"greater\", \"less\")", "IF comparison");
        AddTestCase(sheetData, row++, "IF", "=IF(F5>10, \"high\", \"low\")", "IF cell reference", ("F5", "15"));

        // Nested IF
        AddTestCase(sheetData, row++, "IF", "=IF(F6>20, \"high\", IF(F6>10, \"medium\", \"low\"))", "Nested IF", ("F6", "15"));

        // AND tests
        AddTestCase(sheetData, row++, "AND", "=AND(TRUE, TRUE)", "AND both true");
        AddTestCase(sheetData, row++, "AND", "=AND(TRUE, FALSE)", "AND one false");
        AddTestCase(sheetData, row++, "AND", "=AND(10>5, 20>15)", "AND comparisons");

        // OR tests
        AddTestCase(sheetData, row++, "OR", "=OR(TRUE, FALSE)", "OR one true");
        AddTestCase(sheetData, row++, "OR", "=OR(FALSE, FALSE)", "OR both false");
        AddTestCase(sheetData, row++, "OR", "=OR(10>5, 3>15)", "OR comparisons");

        // NOT tests
        AddTestCase(sheetData, row++, "NOT", "=NOT(TRUE)", "NOT true");
        AddTestCase(sheetData, row++, "NOT", "=NOT(FALSE)", "NOT false");

        // Complex logical expressions
        AddTestCase(sheetData, row++, "AND+OR", "=AND(OR(F15>10, F15<5), F15<>0)", "Complex logic", ("F15", "12"));

        // CHOOSE tests
        AddTestCase(sheetData, row++, "CHOOSE", "=CHOOSE(2, \"A\", \"B\", \"C\")", "Choose 2nd value");
        AddTestCase(sheetData, row++, "CHOOSE", "=CHOOSE(1, 10, 20, 30)", "Choose 1st value");
        AddTestCase(sheetData, row++, "CHOOSE", "=CHOOSE(F18, \"Red\", \"Green\", \"Blue\")", "Choose from cell", ("F18", "3"));

        // IFS tests
        AddTestCase(sheetData, row++, "IFS", "=IFS(F20>20, \"High\", F20>10, \"Medium\", TRUE, \"Low\")", "IFS multiple conditions", ("F20", "15"));
        AddTestCase(sheetData, row++, "IFS", "=IFS(5>10, \"No\", 3<2, \"No\", TRUE, \"Yes\")", "IFS with default");

        // SWITCH tests
        AddTestCase(sheetData, row++, "SWITCH", "=SWITCH(2, 1, \"One\", 2, \"Two\", 3, \"Three\")", "Switch number");
        AddTestCase(sheetData, row++, "SWITCH", "=SWITCH(\"B\", \"A\", 10, \"B\", 20, \"C\", 30)", "Switch text");
        AddTestCase(sheetData, row++, "SWITCH", "=SWITCH(F24, 1, \"First\", 2, \"Second\", \"Other\")", "Switch with default", ("F24", "5"));

        // XOR tests
        AddTestCase(sheetData, row++, "XOR", "=XOR(TRUE, FALSE)", "XOR one true");
        AddTestCase(sheetData, row++, "XOR", "=XOR(TRUE, TRUE)", "XOR both true");
        AddTestCase(sheetData, row++, "XOR", "=XOR(FALSE, FALSE)", "XOR both false");

        // IFNA tests
        AddTestCase(sheetData, row++, "IFNA", "=IFNA(10, \"N/A\")", "IFNA with value");
        AddTestCase(sheetData, row++, "IFNA", "=IFNA(F30, \"Missing\")", "IFNA with cell", ("F30", "42"));

        SortCellsInRows(sheetData);
        worksheetPart.Worksheet.Save();
    }

    private static void CreateTextFunctionsSheet(WorkbookPart workbookPart, Sheets sheets, uint sheetId)
    {
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        worksheetPart.Worksheet = new Worksheet(sheetData);

        sheets.AppendChild(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = "Text",
        });

        AddHeader(sheetData, 1);
        var row = 2;

        // Setup test data in columns F, G, H
        AddTextCell(sheetData, "F1", "Hello");
        AddTextCell(sheetData, "G1", "World");
        AddTextCell(sheetData, "H1", "  Excel  ");

        // CONCATENATE tests
        AddTestCase(sheetData, row++, "CONCATENATE", "=CONCATENATE(\"Hello\", \" \", \"World\")", "Basic concatenate");
        AddTestCase(sheetData, row++, "CONCATENATE", "=CONCATENATE(F1, \" \", G1)", "Concatenate cells");

        // LEFT/RIGHT/MID tests
        AddTestCase(sheetData, row++, "LEFT", "=LEFT(\"Hello\", 3)", "Left 3 chars");
        AddTestCase(sheetData, row++, "RIGHT", "=RIGHT(\"World\", 3)", "Right 3 chars");
        AddTestCase(sheetData, row++, "MID", "=MID(\"Hello\", 2, 3)", "Mid substring");

        // LEN tests
        AddTestCase(sheetData, row++, "LEN", "=LEN(\"Hello\")", "Length of string");
        AddTestCase(sheetData, row++, "LEN", "=LEN(F1)", "Length of cell");

        // TRIM tests
        AddTestCase(sheetData, row++, "TRIM", "=TRIM(\"  Hello  \")", "Trim spaces");
        AddTestCase(sheetData, row++, "TRIM", "=TRIM(H1)", "Trim cell");

        // UPPER/LOWER/PROPER tests
        AddTestCase(sheetData, row++, "UPPER", "=UPPER(\"hello\")", "Convert to uppercase");
        AddTestCase(sheetData, row++, "LOWER", "=LOWER(\"WORLD\")", "Convert to lowercase");
        AddTestCase(sheetData, row++, "PROPER", "=PROPER(\"hello world\")", "Convert to proper case");

        // TEXT tests
        AddTestCase(sheetData, row++, "TEXT", "=TEXT(1234.5, \"0.00\")", "Format number");
        AddTestCase(sheetData, row++, "TEXT", "=TEXT(0.75, \"0%\")", "Format as percent");

        // VALUE tests
        AddTestCase(sheetData, row++, "VALUE", "=VALUE(\"123\")", "Convert string to number");
        AddTestCase(sheetData, row++, "VALUE", "=VALUE(\"45.67\")", "Convert decimal string");

        // FIND/SEARCH tests
        AddTestCase(sheetData, row++, "FIND", "=FIND(\"l\", \"Hello\")", "Find character");
        AddTestCase(sheetData, row++, "SEARCH", "=SEARCH(\"o\", \"Hello\")", "Search character");

        // SUBSTITUTE tests
        AddTestCase(sheetData, row++, "SUBSTITUTE", "=SUBSTITUTE(\"Hello\", \"l\", \"L\")", "Substitute all");
        AddTestCase(sheetData, row++, "SUBSTITUTE", "=SUBSTITUTE(\"banana\", \"a\", \"o\", 2)", "Substitute 2nd occurrence");

        // REPLACE tests
        AddTestCase(sheetData, row++, "REPLACE", "=REPLACE(\"Hello\", 1, 2, \"Wi\")", "Replace first 2 chars");
        AddTestCase(sheetData, row++, "REPLACE", "=REPLACE(\"Excel\", 2, 3, \"tra\")", "Replace middle chars");

        // REPT tests
        AddTestCase(sheetData, row++, "REPT", "=REPT(\"*\", 5)", "Repeat star 5 times");
        AddTestCase(sheetData, row++, "REPT", "=REPT(\"AB\", 3)", "Repeat AB 3 times");

        // EXACT tests
        AddTestCase(sheetData, row++, "EXACT", "=EXACT(\"Hello\", \"Hello\")", "Exact match");
        AddTestCase(sheetData, row++, "EXACT", "=EXACT(\"Hello\", \"hello\")", "Case sensitive no match");

        // CHAR tests
        AddTestCase(sheetData, row++, "CHAR", "=CHAR(65)", "Char code 65 (A)");
        AddTestCase(sheetData, row++, "CHAR", "=CHAR(90)", "Char code 90 (Z)");

        // CODE tests
        AddTestCase(sheetData, row++, "CODE", "=CODE(\"A\")", "Code of A");
        AddTestCase(sheetData, row++, "CODE", "=CODE(\"Z\")", "Code of Z");

        // CLEAN tests
        AddTestCase(sheetData, row++, "CLEAN", "=LEN(CLEAN(CHAR(10) & \"Text\" & CHAR(13)))", "Clean non-printable chars");

        // T tests
        AddTestCase(sheetData, row++, "T", "=T(\"Hello\")", "T with text");
        AddTestCase(sheetData, row++, "T", "=T(123)", "T with number");

        // CONCAT tests
        AddTestCase(sheetData, row++, "CONCAT", "=CONCAT(\"Hello\", \" \", \"World\")", "Concat strings");
        AddTestCase(sheetData, row++, "CONCAT", "=CONCAT(F30, \" \", G30)", "Concat cells", ("F30", "Good"), ("G30", "Day"));

        // TEXTJOIN tests
        AddTestCase(sheetData, row++, "TEXTJOIN", "=TEXTJOIN(\", \", TRUE, \"A\", \"B\", \"C\")", "TextJoin with delimiter");
        AddTestCase(sheetData, row++, "TEXTJOIN", "=TEXTJOIN(\"-\", FALSE, F33:H33)", "TextJoin range", ("F33", "Red"), ("G33", "Green"), ("H33", "Blue"));

        // REVERSE tests
        AddTestCase(sheetData, row++, "REVERSE", "=REVERSE(\"Hello\")", "Reverse string");
        AddTestCase(sheetData, row++, "REVERSE", "=REVERSE(\"12345\")", "Reverse numbers");

        // FIXED tests
        AddTestCase(sheetData, row++, "FIXED", "=FIXED(1234.567, 2)", "Fixed 2 decimals");
        AddTestCase(sheetData, row++, "FIXED", "=FIXED(1000.5, 1, TRUE)", "Fixed no commas");

        // DOLLAR tests
        AddTestCase(sheetData, row++, "DOLLAR", "=DOLLAR(1234.567, 2)", "Dollar format");
        AddTestCase(sheetData, row++, "DOLLAR", "=DOLLAR(99.99)", "Dollar default");

        // NUMBERVALUE tests
        AddTestCase(sheetData, row++, "NUMBERVALUE", "=NUMBERVALUE(\"123.45\")", "Number value");
        AddTestCase(sheetData, row++, "NUMBERVALUE", "=NUMBERVALUE(\"1,234.56\")", "Number value with comma");

        // TRIMALL tests (if implemented as custom function)
        AddTestCase(sheetData, row++, "TRIM", "=LEN(TRIM(\"  Extra   Spaces  \"))", "Trim all spaces length");

        // UNICHAR tests
        AddTestCase(sheetData, row++, "UNICHAR", "=UNICHAR(65)", "Unichar 65");
        AddTestCase(sheetData, row++, "UNICHAR", "=UNICHAR(8364)", "Unichar Euro sign");

        // UNICODE tests
        AddTestCase(sheetData, row++, "UNICODE", "=UNICODE(\"A\")", "Unicode of A");
        AddTestCase(sheetData, row++, "UNICODE", "=UNICODE(\"€\")", "Unicode of Euro");

        SortCellsInRows(sheetData);
        worksheetPart.Worksheet.Save();
    }

    private static void CreateLookupFunctionsSheet(WorkbookPart workbookPart, Sheets sheets, uint sheetId)
    {
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        worksheetPart.Worksheet = new Worksheet(sheetData);

        sheets.AppendChild(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = "Lookup",
        });

        // Create lookup table (A1:C5) - note: no standard header for this sheet
        AddTextCell(sheetData, "A1", "ID");
        AddTextCell(sheetData, "B1", "Name");
        AddTextCell(sheetData, "C1", "Score");

        AddCell(sheetData, "A2", "1");
        AddTextCell(sheetData, "B2", "Alice");
        AddCell(sheetData, "C2", "95");

        AddCell(sheetData, "A3", "2");
        AddTextCell(sheetData, "B3", "Bob");
        AddCell(sheetData, "C3", "87");

        AddCell(sheetData, "A4", "3");
        AddTextCell(sheetData, "B4", "Carol");
        AddCell(sheetData, "C4", "92");

        AddCell(sheetData, "A5", "4");
        AddTextCell(sheetData, "B5", "Dave");
        AddCell(sheetData, "C5", "88");

        // Test cases start at row 7 (after lookup table data)
        var row = 7;

        // VLOOKUP tests
        AddTestCase(sheetData, row++, "VLOOKUP", "=VLOOKUP(2, A2:C5, 2, FALSE)", "VLOOKUP name by ID");
        AddTestCase(sheetData, row++, "VLOOKUP", "=VLOOKUP(3, A2:C5, 3, FALSE)", "VLOOKUP score by ID");
        AddTestCase(sheetData, row++, "VLOOKUP", "=VLOOKUP(1, A2:C5, 2, FALSE)", "VLOOKUP first row");

        // HLOOKUP tests (create horizontal table)
        AddTextCell(sheetData, "F1", "Alice");
        AddTextCell(sheetData, "G1", "Bob");
        AddTextCell(sheetData, "H1", "Carol");
        AddCell(sheetData, "F2", "95");
        AddCell(sheetData, "G2", "87");
        AddCell(sheetData, "H2", "92");

        AddTestCase(sheetData, row++, "HLOOKUP", "=HLOOKUP(\"Bob\", F1:H2, 2, FALSE)", "HLOOKUP score by name");

        // INDEX tests
        AddTestCase(sheetData, row++, "INDEX", "=INDEX(A2:C5, 2, 2)", "INDEX row 2 col 2");
        AddTestCase(sheetData, row++, "INDEX", "=INDEX(B2:B5, 3)", "INDEX single column");

        // MATCH tests
        AddTestCase(sheetData, row++, "MATCH", "=MATCH(\"Bob\", B2:B5, 0)", "MATCH exact");
        AddTestCase(sheetData, row++, "MATCH", "=MATCH(3, A2:A5, 0)", "MATCH number");

        // COLUMN tests
        AddTextCell(sheetData, "K12", "Sample");
        AddTestCase(sheetData, row++, "COLUMN", "=COLUMN(K12)", "Column of K12");
        AddTestCase(sheetData, row++, "COLUMN", "=COLUMN(A1)", "Column of A1");

        // ROW tests
        AddTestCase(sheetData, row++, "ROW", "=ROW(A5)", "Row of A5");
        AddTestCase(sheetData, row++, "ROW", "=ROW(B10)", "Row of B10");

        // COLUMNS tests
        AddTestCase(sheetData, row++, "COLUMNS", "=COLUMNS(A1:E1)", "Columns in range");
        AddTestCase(sheetData, row++, "COLUMNS", "=COLUMNS(F1:H2)", "Columns in 2D range");

        // ROWS tests
        AddTestCase(sheetData, row++, "ROWS", "=ROWS(A1:A5)", "Rows in range");
        AddTestCase(sheetData, row++, "ROWS", "=ROWS(F1:H3)", "Rows in 2D range");

        // ADDRESS tests
        AddTestCase(sheetData, row++, "ADDRESS", "=ADDRESS(1, 1)", "Address A1");
        AddTestCase(sheetData, row++, "ADDRESS", "=ADDRESS(5, 3)", "Address C5");

        // OFFSET tests
        AddCell(sheetData, "M1", "100");
        AddTestCase(sheetData, row++, "OFFSET", "=OFFSET(M1, 0, 0)", "Offset no movement");
        AddCell(sheetData, "M2", "200");
        AddTestCase(sheetData, row++, "OFFSET", "=OFFSET(M1, 1, 0)", "Offset down 1");

        // INDIRECT tests
        AddTextCell(sheetData, "N1", "A1");
        AddTestCase(sheetData, row++, "INDIRECT", "=INDIRECT(\"A1\")", "Indirect to A1");
        AddTestCase(sheetData, row++, "INDIRECT", "=ISNUMBER(INDIRECT(N1))", "Indirect from cell");

        SortCellsInRows(sheetData);
        worksheetPart.Worksheet.Save();
    }

    private static void CreateDateTimeFunctionsSheet(WorkbookPart workbookPart, Sheets sheets, uint sheetId)
    {
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        worksheetPart.Worksheet = new Worksheet(sheetData);

        sheets.AppendChild(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = "DateTime",
        });

        AddHeader(sheetData, 1);
        var row = 2;

        // Store test data in column F (avoid conflict with test case columns A-D)
        AddCell(sheetData, "F1", "45096"); // Excel serial date for 2023-06-15
        AddCell(sheetData, "G1", "0.5"); // Noon (12:00:00)

        // DATE tests
        AddTestCase(sheetData, row++, "DATE", "=DATE(2023, 6, 15)", "Create date");
        AddTestCase(sheetData, row++, "DATE", "=DATE(2024, 1, 1)", "New year date");

        // YEAR/MONTH/DAY tests
        AddTestCase(sheetData, row++, "YEAR", "=YEAR(F1)", "Extract year");
        AddTestCase(sheetData, row++, "MONTH", "=MONTH(F1)", "Extract month");
        AddTestCase(sheetData, row++, "DAY", "=DAY(F1)", "Extract day");

        // WEEKDAY test
        AddTestCase(sheetData, row++, "WEEKDAY", "=WEEKDAY(F1)", "Get weekday");

        // TODAY/NOW tests (will vary by when Excel calculates)
        AddTestCase(sheetData, row++, "TODAY", "=YEAR(TODAY())", "Today's year");
        AddTestCase(sheetData, row++, "NOW", "=YEAR(NOW())", "Now's year");

        // HOUR/MINUTE/SECOND tests (using time value)
        AddTestCase(sheetData, row++, "HOUR", "=HOUR(G1)", "Extract hour from time");
        AddTestCase(sheetData, row++, "MINUTE", "=MINUTE(G1)", "Extract minute from time");
        AddTestCase(sheetData, row++, "SECOND", "=SECOND(G1)", "Extract second from time");

        // DAYS tests
        AddCell(sheetData, "H1", "45096"); // 2023-06-15
        AddCell(sheetData, "I1", "45106"); // 2023-06-25
        AddTestCase(sheetData, row++, "DAYS", "=DAYS(I1, H1)", "Days between dates");
        AddTestCase(sheetData, row++, "DAYS", "=DAYS(DATE(2023,12,31), DATE(2023,1,1))", "Days in year");

        // TIME tests
        AddTestCase(sheetData, row++, "TIME", "=HOUR(TIME(14, 30, 0))", "Extract hour from TIME");
        AddTestCase(sheetData, row++, "TIME", "=MINUTE(TIME(14, 30, 45))", "Extract minute from TIME");

        // TIMEVALUE tests
        AddTestCase(sheetData, row++, "TIMEVALUE", "=HOUR(TIMEVALUE(\"14:30:00\"))", "Timevalue hour");
        AddTestCase(sheetData, row++, "TIMEVALUE", "=MINUTE(TIMEVALUE(\"14:30:00\"))", "Timevalue minute");

        // DATEVALUE tests
        AddTestCase(sheetData, row++, "DATEVALUE", "=YEAR(DATEVALUE(\"2023-06-15\"))", "Datevalue year");
        AddTestCase(sheetData, row++, "DATEVALUE", "=MONTH(DATEVALUE(\"2023-06-15\"))", "Datevalue month");

        // DAYS360 tests
        AddTestCase(sheetData, row++, "DAYS360", "=DAYS360(DATE(2023,1,1), DATE(2023,12,31))", "Days360 full year");
        AddTestCase(sheetData, row++, "DAYS360", "=DAYS360(DATE(2023,1,15), DATE(2023,2,15))", "Days360 one month");

        // EOMONTH tests
        AddTestCase(sheetData, row++, "EOMONTH", "=DAY(EOMONTH(DATE(2023,1,15), 0))", "EOMONTH same month");
        AddTestCase(sheetData, row++, "EOMONTH", "=MONTH(EOMONTH(DATE(2023,1,15), 1))", "EOMONTH next month");

        // EDATE tests
        AddTestCase(sheetData, row++, "EDATE", "=MONTH(EDATE(DATE(2023,1,15), 2))", "EDATE 2 months later");
        AddTestCase(sheetData, row++, "EDATE", "=YEAR(EDATE(DATE(2023,11,15), 2))", "EDATE across year");

        // NETWORKDAYS tests
        AddTestCase(sheetData, row++, "NETWORKDAYS", "=NETWORKDAYS(DATE(2023,1,2), DATE(2023,1,6))", "Networkdays week");

        // WORKDAY tests
        AddTestCase(sheetData, row++, "WORKDAY", "=WEEKDAY(WORKDAY(DATE(2023,1,2), 5))", "Workday weekday check");

        // WEEKNUM tests
        AddTestCase(sheetData, row++, "WEEKNUM", "=WEEKNUM(DATE(2023,1,1))", "Week number of year start");
        AddTestCase(sheetData, row++, "WEEKNUM", "=WEEKNUM(DATE(2023,6,15))", "Week number mid year");

        // YEARFRAC tests
        AddTestCase(sheetData, row++, "YEARFRAC", "=ROUND(YEARFRAC(DATE(2023,1,1), DATE(2024,1,1)), 2)", "Year fraction full year");
        AddTestCase(sheetData, row++, "YEARFRAC", "=ROUND(YEARFRAC(DATE(2023,1,1), DATE(2023,7,1)), 2)", "Year fraction half year");

        // DATEDIF tests
        AddTestCase(sheetData, row++, "DATEDIF", "=DATEDIF(DATE(2023,1,1), DATE(2024,1,1), \"Y\")", "Datedif years");
        AddTestCase(sheetData, row++, "DATEDIF", "=DATEDIF(DATE(2023,1,1), DATE(2023,7,1), \"M\")", "Datedif months");
        AddTestCase(sheetData, row++, "DATEDIF", "=DATEDIF(DATE(2023,1,1), DATE(2023,1,15), \"D\")", "Datedif days");

        SortCellsInRows(sheetData);
        worksheetPart.Worksheet.Save();
    }

    private static void CreateStatisticalFunctionsSheet(WorkbookPart workbookPart, Sheets sheets, uint sheetId)
    {
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        worksheetPart.Worksheet = new Worksheet(sheetData);

        sheets.AppendChild(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = "Statistical",
        });

        AddHeader(sheetData, 1);
        var row = 2;

        // Create test data in column F (avoid conflict with test case columns A-D)
        for (int i = 0; i < 10; i++)
        {
            AddCell(sheetData, $"F{i + 1}", ((i + 1) * 5).ToString());
        }

        // MEDIAN tests
        AddTestCase(sheetData, row++, "MEDIAN", "=MEDIAN(1,2,3,4,5)", "Median odd count");
        AddTestCase(sheetData, row++, "MEDIAN", "=MEDIAN(1,2,3,4)", "Median even count");
        AddTestCase(sheetData, row++, "MEDIAN", "=MEDIAN(F1:F10)", "Median range");

        // MODE tests
        AddTestCase(sheetData, row++, "MODE", "=MODE(1,2,2,3,4)", "Mode basic");

        // STDEV tests
        AddTestCase(sheetData, row++, "STDEV", "=STDEV(1,2,3,4,5)", "Standard deviation");
        AddTestCase(sheetData, row++, "STDEV", "=STDEV(F1:F5)", "Stdev range");

        // VAR tests
        AddTestCase(sheetData, row++, "VAR", "=VAR(1,2,3,4,5)", "Variance");
        AddTestCase(sheetData, row++, "VAR", "=VAR(F1:F5)", "Var range");

        // RANK tests
        AddTestCase(sheetData, row++, "RANK", "=RANK(25, F1:F10)", "Rank in range");

        // SUMIF tests
        AddTestCase(sheetData, row++, "SUMIF", "=SUMIF(F1:F10, \">20\")", "Sumif greater than 20");
        AddCell(sheetData, "G11", "10");
        AddCell(sheetData, "H11", "5");
        AddCell(sheetData, "G12", "20");
        AddCell(sheetData, "H12", "8");
        AddCell(sheetData, "G13", "30");
        AddCell(sheetData, "H13", "12");
        AddTestCase(sheetData, row++, "SUMIF", "=SUMIF(G11:G13, \">15\", H11:H13)", "Sumif with sum range");

        // COUNTIF tests
        AddTestCase(sheetData, row++, "COUNTIF", "=COUNTIF(F1:F10, \">25\")", "Countif greater than 25");
        AddTestCase(sheetData, row++, "COUNTIF", "=COUNTIF(F1:F5, \"<=15\")", "Countif less equal 15");

        // AVERAGEIF tests
        AddTestCase(sheetData, row++, "AVERAGEIF", "=AVERAGEIF(F1:F10, \">20\")", "Averageif greater than 20");
        AddTestCase(sheetData, row++, "AVERAGEIF", "=AVERAGEIF(G11:G13, \">15\", H11:H13)", "Averageif with avg range");

        // STDEVP tests
        AddTestCase(sheetData, row++, "STDEVP", "=STDEVP(1,2,3,4,5)", "Population stdev");
        AddTestCase(sheetData, row++, "STDEVP", "=STDEVP(F1:F5)", "Stdevp range");

        // VARP tests
        AddTestCase(sheetData, row++, "VARP", "=VARP(1,2,3,4,5)", "Population variance");
        AddTestCase(sheetData, row++, "VARP", "=VARP(F1:F5)", "Varp range");

        // LARGE tests
        AddTestCase(sheetData, row++, "LARGE", "=LARGE(F1:F10, 1)", "Largest value");
        AddTestCase(sheetData, row++, "LARGE", "=LARGE(F1:F10, 3)", "3rd largest");

        // SMALL tests
        AddTestCase(sheetData, row++, "SMALL", "=SMALL(F1:F10, 1)", "Smallest value");
        AddTestCase(sheetData, row++, "SMALL", "=SMALL(F1:F10, 3)", "3rd smallest");

        // PERCENTILE tests
        AddTestCase(sheetData, row++, "PERCENTILE", "=PERCENTILE(F1:F10, 0.5)", "50th percentile");
        AddTestCase(sheetData, row++, "PERCENTILE", "=PERCENTILE(F1:F10, 0.75)", "75th percentile");

        // QUARTILE tests
        AddTestCase(sheetData, row++, "QUARTILE", "=QUARTILE(F1:F10, 1)", "1st quartile");
        AddTestCase(sheetData, row++, "QUARTILE", "=QUARTILE(F1:F10, 3)", "3rd quartile");

        // AVERAGEIFS tests
        AddCell(sheetData, "J1", "10");
        AddCell(sheetData, "K1", "5");
        AddCell(sheetData, "L1", "A");
        AddCell(sheetData, "J2", "20");
        AddCell(sheetData, "K2", "8");
        AddCell(sheetData, "L2", "B");
        AddCell(sheetData, "J3", "30");
        AddCell(sheetData, "K3", "12");
        AddCell(sheetData, "L3", "A");
        AddTestCase(sheetData, row++, "AVERAGEIFS", "=AVERAGEIFS(K1:K3, J1:J3, \">15\")", "Averageifs single criteria");

        // MAXIFS tests
        AddTestCase(sheetData, row++, "MAXIFS", "=MAXIFS(K1:K3, J1:J3, \">10\")", "Maxifs with criteria");

        // MINIFS tests
        AddTestCase(sheetData, row++, "MINIFS", "=MINIFS(K1:K3, J1:J3, \">10\")", "Minifs with criteria");

        // CORREL tests
        AddCell(sheetData, "M1", "1");
        AddCell(sheetData, "N1", "2");
        AddCell(sheetData, "M2", "2");
        AddCell(sheetData, "N2", "4");
        AddCell(sheetData, "M3", "3");
        AddCell(sheetData, "N3", "6");
        AddTestCase(sheetData, row++, "CORREL", "=ROUND(CORREL(M1:M3, N1:N3), 2)", "Correlation coefficient");

        // COVARIANCE.P tests
        AddTestCase(sheetData, row++, "COVARIANCE.P", "=ROUND(COVARIANCE.P(M1:M3, N1:N3), 2)", "Population covariance");

        // COVARIANCE.S tests
        AddTestCase(sheetData, row++, "COVARIANCE.S", "=ROUND(COVARIANCE.S(M1:M3, N1:N3), 2)", "Sample covariance");

        // SLOPE tests
        AddTestCase(sheetData, row++, "SLOPE", "=SLOPE(N1:N3, M1:M3)", "Slope of regression");

        // INTERCEPT tests
        AddTestCase(sheetData, row++, "INTERCEPT", "=INTERCEPT(N1:N3, M1:M3)", "Intercept of regression");

        // SKEW tests
        AddTestCase(sheetData, row++, "SKEW", "=ROUND(SKEW(1,2,3,4,10), 2)", "Skewness");

        // KURT tests
        AddTestCase(sheetData, row++, "KURT", "=ROUND(KURT(1,2,3,4,5,6,7,8,9,10), 2)", "Kurtosis");

        // FREQUENCY tests (returns array, test first element)
        AddCell(sheetData, "O1", "5");
        AddCell(sheetData, "O2", "15");
        AddCell(sheetData, "O3", "25");
        AddCell(sheetData, "P1", "10");
        AddCell(sheetData, "P2", "20");
        AddTestCase(sheetData, row++, "FREQUENCY", "=SUM(FREQUENCY(O1:O3, P1:P2))", "Frequency sum");

        // COUNTBLANK tests
        AddCell(sheetData, "Q1", "10");
        AddTextCell(sheetData, "Q2", "");
        AddCell(sheetData, "Q3", "30");
        AddTestCase(sheetData, row++, "COUNTBLANK", "=COUNTBLANK(Q1:Q3)", "Count blank cells");

        SortCellsInRows(sheetData);
        worksheetPart.Worksheet.Save();
    }

    private static void CreateFinancialFunctionsSheet(WorkbookPart workbookPart, Sheets sheets, uint sheetId)
    {
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        worksheetPart.Worksheet = new Worksheet(sheetData);

        sheets.AppendChild(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = "Financial",
        });

        AddHeader(sheetData, 1);
        var row = 2;

        // PMT tests
        AddTestCase(sheetData, row++, "PMT", "=ROUND(PMT(0.05/12, 60, 10000), 2)", "Monthly payment");
        AddTestCase(sheetData, row++, "PMT", "=ROUND(PMT(0.08/12, 120, 50000), 2)", "Loan payment");

        // FV tests
        AddTestCase(sheetData, row++, "FV", "=ROUND(FV(0.05/12, 60, -100), 2)", "Future value");
        AddTestCase(sheetData, row++, "FV", "=ROUND(FV(0.06/12, 120, -200, 0, 0), 2)", "FV with PV");

        // PV tests
        AddTestCase(sheetData, row++, "PV", "=ROUND(PV(0.05/12, 60, 100), 2)", "Present value");
        AddTestCase(sheetData, row++, "PV", "=ROUND(PV(0.08/12, 120, 500), 2)", "PV of annuity");

        // NPER tests
        AddTestCase(sheetData, row++, "NPER", "=ROUND(NPER(0.05/12, -100, 5000), 2)", "Number of periods");
        AddTestCase(sheetData, row++, "NPER", "=ROUND(NPER(0.06/12, -200, 10000, 0), 2)", "NPER to payoff");

        // RATE tests
        AddTestCase(sheetData, row++, "RATE", "=ROUND(RATE(60, -100, 5000)*12, 4)", "Annual rate");

        // NPV tests
        AddCell(sheetData, "F12", "-10000");
        AddCell(sheetData, "G12", "3000");
        AddCell(sheetData, "H12", "4000");
        AddCell(sheetData, "I12", "5000");
        AddTestCase(sheetData, row++, "NPV", "=ROUND(NPV(0.1, G12:I12) + F12, 2)", "Net present value");

        // IRR tests
        AddCell(sheetData, "F14", "-10000");
        AddCell(sheetData, "G14", "2000");
        AddCell(sheetData, "H14", "4000");
        AddCell(sheetData, "I14", "6000");
        AddTestCase(sheetData, row++, "IRR", "=ROUND(IRR(F14:I14), 4)", "Internal rate of return");

        // IPMT tests
        AddTestCase(sheetData, row++, "IPMT", "=ROUND(IPMT(0.05/12, 1, 60, 10000), 2)", "Interest payment period 1");
        AddTestCase(sheetData, row++, "IPMT", "=ROUND(IPMT(0.08/12, 6, 120, 50000), 2)", "Interest payment period 6");

        // PPMT tests
        AddTestCase(sheetData, row++, "PPMT", "=ROUND(PPMT(0.05/12, 1, 60, 10000), 2)", "Principal payment period 1");
        AddTestCase(sheetData, row++, "PPMT", "=ROUND(PPMT(0.08/12, 6, 120, 50000), 2)", "Principal payment period 6");

        // SLN tests
        AddTestCase(sheetData, row++, "SLN", "=SLN(10000, 1000, 5)", "Straight line depreciation");
        AddTestCase(sheetData, row++, "SLN", "=SLN(50000, 5000, 10)", "SLN 10 years");

        // DB tests
        AddTestCase(sheetData, row++, "DB", "=ROUND(DB(10000, 1000, 5, 1), 2)", "Declining balance year 1");
        AddTestCase(sheetData, row++, "DB", "=ROUND(DB(10000, 1000, 5, 2), 2)", "Declining balance year 2");

        // DDB tests
        AddTestCase(sheetData, row++, "DDB", "=ROUND(DDB(10000, 1000, 5, 1), 2)", "Double declining year 1");
        AddTestCase(sheetData, row++, "DDB", "=ROUND(DDB(10000, 1000, 5, 2), 2)", "Double declining year 2");

        // SYD tests
        AddTestCase(sheetData, row++, "SYD", "=ROUND(SYD(10000, 1000, 5, 1), 2)", "Sum of years digits year 1");
        AddTestCase(sheetData, row++, "SYD", "=ROUND(SYD(10000, 1000, 5, 2), 2)", "Sum of years digits year 2");

        SortCellsInRows(sheetData);
        worksheetPart.Worksheet.Save();
    }

    private static void CreateEngineeringFunctionsSheet(WorkbookPart workbookPart, Sheets sheets, uint sheetId)
    {
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        worksheetPart.Worksheet = new Worksheet(sheetData);

        sheets.AppendChild(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = "Engineering",
        });

        AddHeader(sheetData, 1);
        var row = 2;

        // CONVERT tests
        AddTestCase(sheetData, row++, "CONVERT", "=CONVERT(1, \"m\", \"ft\")", "Meters to feet");
        AddTestCase(sheetData, row++, "CONVERT", "=CONVERT(100, \"F\", \"C\")", "Fahrenheit to Celsius");

        // HEX2DEC tests
        AddTestCase(sheetData, row++, "HEX2DEC", "=HEX2DEC(\"FF\")", "Hex FF to decimal");
        AddTestCase(sheetData, row++, "HEX2DEC", "=HEX2DEC(\"A5\")", "Hex A5 to decimal");

        // DEC2HEX tests
        AddTestCase(sheetData, row++, "DEC2HEX", "=DEC2HEX(255)", "Decimal 255 to hex");
        AddTestCase(sheetData, row++, "DEC2HEX", "=DEC2HEX(165)", "Decimal 165 to hex");

        // BIN2DEC tests
        AddTestCase(sheetData, row++, "BIN2DEC", "=BIN2DEC(\"1010\")", "Binary 1010 to decimal");
        AddTestCase(sheetData, row++, "BIN2DEC", "=BIN2DEC(\"11111111\")", "Binary 11111111 to decimal");

        // DEC2BIN tests
        AddTestCase(sheetData, row++, "DEC2BIN", "=DEC2BIN(10)", "Decimal 10 to binary");
        AddTestCase(sheetData, row++, "DEC2BIN", "=DEC2BIN(255)", "Decimal 255 to binary");

        // OCT2DEC tests
        AddTestCase(sheetData, row++, "OCT2DEC", "=OCT2DEC(\"77\")", "Octal 77 to decimal");
        AddTestCase(sheetData, row++, "OCT2DEC", "=OCT2DEC(\"100\")", "Octal 100 to decimal");

        // DEC2OCT tests
        AddTestCase(sheetData, row++, "DEC2OCT", "=DEC2OCT(63)", "Decimal 63 to octal");
        AddTestCase(sheetData, row++, "DEC2OCT", "=DEC2OCT(64)", "Decimal 64 to octal");
        // HEX2BIN tests
        AddTestCase(sheetData, row++, "HEX2BIN", "=HEX2BIN(\"F\")", "Hex F to binary");
        AddTestCase(sheetData, row++, "HEX2BIN", "=HEX2BIN(\"A\")", "Hex A to binary");

        // HEX2OCT tests
        AddTestCase(sheetData, row++, "HEX2OCT", "=HEX2OCT(\"FF\")", "Hex FF to octal");
        AddTestCase(sheetData, row++, "HEX2OCT", "=HEX2OCT(\"3F\")", "Hex 3F to octal");

        // BIN2HEX tests
        AddTestCase(sheetData, row++, "BIN2HEX", "=BIN2HEX(\"1111\")", "Binary 1111 to hex");
        AddTestCase(sheetData, row++, "BIN2HEX", "=BIN2HEX(\"10101010\")", "Binary 10101010 to hex");

        // BIN2OCT tests
        AddTestCase(sheetData, row++, "BIN2OCT", "=BIN2OCT(\"1010\")", "Binary 1010 to octal");
        AddTestCase(sheetData, row++, "BIN2OCT", "=BIN2OCT(\"111111\")", "Binary 111111 to octal");

        // OCT2HEX tests
        AddTestCase(sheetData, row++, "OCT2HEX", "=OCT2HEX(\"77\")", "Octal 77 to hex");
        AddTestCase(sheetData, row++, "OCT2HEX", "=OCT2HEX(\"100\")", "Octal 100 to hex");

        // OCT2BIN tests
        AddTestCase(sheetData, row++, "OCT2BIN", "=OCT2BIN(\"7\")", "Octal 7 to binary");
        AddTestCase(sheetData, row++, "OCT2BIN", "=OCT2BIN(\"12\")", "Octal 12 to binary");

        // DELTA tests
        AddTestCase(sheetData, row++, "DELTA", "=DELTA(5, 5)", "Delta equal values");
        AddTestCase(sheetData, row++, "DELTA", "=DELTA(5, 3)", "Delta unequal values");
        AddTestCase(sheetData, row++, "DELTA", "=DELTA(0)", "Delta zero");

        // GESTEP tests
        AddTestCase(sheetData, row++, "GESTEP", "=GESTEP(5, 3)", "Gestep 5 >= 3");
        AddTestCase(sheetData, row++, "GESTEP", "=GESTEP(2, 3)", "Gestep 2 >= 3");
        AddTestCase(sheetData, row++, "GESTEP", "=GESTEP(5)", "Gestep 5 >= 0");

        // ERF tests
        AddTestCase(sheetData, row++, "ERF", "=ROUND(ERF(0), 4)", "Error function of 0");
        AddTestCase(sheetData, row++, "ERF", "=ROUND(ERF(1), 4)", "Error function of 1");

        // ERF.PRECISE tests
        AddTestCase(sheetData, row++, "ERF.PRECISE", "=ROUND(ERF.PRECISE(0.5), 4)", "ERF.Precise of 0.5");
        AddTestCase(sheetData, row++, "ERF.PRECISE", "=ROUND(ERF.PRECISE(1), 4)", "ERF.Precise of 1");

        // ERFC tests
        AddTestCase(sheetData, row++, "ERFC", "=ROUND(ERFC(0), 4)", "Complementary error function of 0");
        AddTestCase(sheetData, row++, "ERFC", "=ROUND(ERFC(1), 4)", "Complementary error function of 1");

        // ERFC.PRECISE tests
        AddTestCase(sheetData, row++, "ERFC.PRECISE", "=ROUND(ERFC.PRECISE(0.5), 4)", "ERFC.Precise of 0.5");
        AddTestCase(sheetData, row++, "ERFC.PRECISE", "=ROUND(ERFC.PRECISE(1), 4)", "ERFC.Precise of 1");

        // BESSELI tests
        AddTestCase(sheetData, row++, "BESSELI", "=ROUND(BESSELI(1, 0), 4)", "Bessel I function order 0");
        AddTestCase(sheetData, row++, "BESSELI", "=ROUND(BESSELI(1, 1), 4)", "Bessel I function order 1");

        // BESSELJ tests
        AddTestCase(sheetData, row++, "BESSELJ", "=ROUND(BESSELJ(1, 0), 4)", "Bessel J function order 0");
        AddTestCase(sheetData, row++, "BESSELJ", "=ROUND(BESSELJ(1, 1), 4)", "Bessel J function order 1");

        // BESSELK tests
        AddTestCase(sheetData, row++, "BESSELK", "=ROUND(BESSELK(1, 0), 4)", "Bessel K function order 0");
        AddTestCase(sheetData, row++, "BESSELK", "=ROUND(BESSELK(1, 1), 4)", "Bessel K function order 1");

        // BESSELY tests
        AddTestCase(sheetData, row++, "BESSELY", "=ROUND(BESSELY(1, 0), 4)", "Bessel Y function order 0");
        AddTestCase(sheetData, row++, "BESSELY", "=ROUND(BESSELY(1, 1), 4)", "Bessel Y function order 1");

        // COMPLEX tests
        AddTestCase(sheetData, row++, "COMPLEX", "=COMPLEX(3, 4)", "Complex number 3+4i");
        AddTestCase(sheetData, row++, "COMPLEX", "=COMPLEX(5, -2)", "Complex number 5-2i");
        AddTestCase(sheetData, row++, "COMPLEX", "=COMPLEX(0, 1)", "Complex number i");

        // IMREAL tests
        AddTestCase(sheetData, row++, "IMREAL", "=IMREAL(\"3+4i\")", "Real part of 3+4i");
        AddTestCase(sheetData, row++, "IMREAL", "=IMREAL(\"5-2i\")", "Real part of 5-2i");

        // IMAGINARY tests
        AddTestCase(sheetData, row++, "IMAGINARY", "=IMAGINARY(\"3+4i\")", "Imaginary part of 3+4i");
        AddTestCase(sheetData, row++, "IMAGINARY", "=IMAGINARY(\"5-2i\")", "Imaginary part of 5-2i");

        // IMABS tests
        AddTestCase(sheetData, row++, "IMABS", "=IMABS(\"3+4i\")", "Absolute value of 3+4i");
        AddTestCase(sheetData, row++, "IMABS", "=IMABS(\"5-12i\")", "Absolute value of 5-12i");

        // IMARGUMENT tests
        AddTestCase(sheetData, row++, "IMARGUMENT", "=ROUND(IMARGUMENT(\"1+i\"), 4)", "Argument of 1+i");
        AddTestCase(sheetData, row++, "IMARGUMENT", "=ROUND(IMARGUMENT(\"1-i\"), 4)", "Argument of 1-i");

        // IMCONJUGATE tests
        AddTestCase(sheetData, row++, "IMCONJUGATE", "=IMCONJUGATE(\"3+4i\")", "Conjugate of 3+4i");
        AddTestCase(sheetData, row++, "IMCONJUGATE", "=IMCONJUGATE(\"5-2i\")", "Conjugate of 5-2i");

        // IMSUM tests
        AddTestCase(sheetData, row++, "IMSUM", "=IMSUM(\"3+4i\", \"1+2i\")", "Sum of complex numbers");
        AddTestCase(sheetData, row++, "IMSUM", "=IMSUM(\"5-2i\", \"3+i\")", "Sum of complex numbers 2");

        // IMSUB tests
        AddTestCase(sheetData, row++, "IMSUB", "=IMSUB(\"3+4i\", \"1+2i\")", "Difference of complex numbers");
        AddTestCase(sheetData, row++, "IMSUB", "=IMSUB(\"5-2i\", \"3+i\")", "Difference of complex numbers 2");

        // IMPRODUCT tests
        AddTestCase(sheetData, row++, "IMPRODUCT", "=IMPRODUCT(\"3+4i\", \"1+2i\")", "Product of complex numbers");
        AddTestCase(sheetData, row++, "IMPRODUCT", "=IMPRODUCT(\"2+i\", \"3-i\")", "Product of complex numbers 2");

        // IMDIV tests
        AddTestCase(sheetData, row++, "IMDIV", "=IMDIV(\"4+2i\", \"2\")", "Division of complex numbers");
        AddTestCase(sheetData, row++, "IMDIV", "=IMDIV(\"1+i\", \"1-i\")", "Division of complex numbers 2");

        // IMPOWER tests
        AddTestCase(sheetData, row++, "IMPOWER", "=IMPOWER(\"2+i\", 2)", "Complex power squared");
        AddTestCase(sheetData, row++, "IMPOWER", "=IMPOWER(\"i\", 3)", "Complex power cubed");

        // IMSQRT tests
        AddTestCase(sheetData, row++, "IMSQRT", "=IMSQRT(\"3+4i\")", "Square root of complex");
        AddTestCase(sheetData, row++, "IMSQRT", "=IMSQRT(\"-1\")", "Square root of -1");

        // IMEXP tests
        AddTestCase(sheetData, row++, "IMEXP", "=IMEXP(\"i\")", "Exponential of i");
        AddTestCase(sheetData, row++, "IMEXP", "=IMEXP(\"1+i\")", "Exponential of 1+i");

        // IMLN tests
        AddTestCase(sheetData, row++, "IMLN", "=IMLN(\"1\")", "Natural log of 1");
        AddTestCase(sheetData, row++, "IMLN", "=IMLN(\"i\")", "Natural log of i");

        // IMLOG10 tests
        AddTestCase(sheetData, row++, "IMLOG10", "=IMLOG10(\"10\")", "Log10 of 10");
        AddTestCase(sheetData, row++, "IMLOG10", "=IMLOG10(\"1+i\")", "Log10 of 1+i");

        // IMLOG2 tests
        AddTestCase(sheetData, row++, "IMLOG2", "=IMLOG2(\"2\")", "Log2 of 2");
        AddTestCase(sheetData, row++, "IMLOG2", "=IMLOG2(\"4+0i\")", "Log2 of 4");

        // IMSIN tests
        AddTestCase(sheetData, row++, "IMSIN", "=IMSIN(\"0\")", "Sin of 0");
        AddTestCase(sheetData, row++, "IMSIN", "=IMSIN(\"i\")", "Sin of i");

        // IMCOS tests
        AddTestCase(sheetData, row++, "IMCOS", "=IMCOS(\"0\")", "Cos of 0");
        AddTestCase(sheetData, row++, "IMCOS", "=IMCOS(\"i\")", "Cos of i");

        // IMTAN tests
        AddTestCase(sheetData, row++, "IMTAN", "=IMTAN(\"0\")", "Tan of 0");
        AddTestCase(sheetData, row++, "IMTAN", "=IMTAN(\"1+i\")", "Tan of 1+i");

        // IMSEC tests
        AddTestCase(sheetData, row++, "IMSEC", "=IMSEC(\"0\")", "Sec of 0");
        AddTestCase(sheetData, row++, "IMSEC", "=IMSEC(\"1+i\")", "Sec of 1+i");

        // IMCSC tests
        AddTestCase(sheetData, row++, "IMCSC", "=IMCSC(\"1\")", "Csc of 1");
        AddTestCase(sheetData, row++, "IMCSC", "=IMCSC(\"1+i\")", "Csc of 1+i");

        // IMCOT tests
        AddTestCase(sheetData, row++, "IMCOT", "=IMCOT(\"1\")", "Cot of 1");
        AddTestCase(sheetData, row++, "IMCOT", "=IMCOT(\"1+i\")", "Cot of 1+i");

        // IMSECH tests
        AddTestCase(sheetData, row++, "IMSECH", "=IMSECH(\"0\")", "Sech of 0");
        AddTestCase(sheetData, row++, "IMSECH", "=IMSECH(\"1+i\")", "Sech of 1+i");

        // IMCSCH tests
        AddTestCase(sheetData, row++, "IMCSCH", "=IMCSCH(\"1\")", "Csch of 1");
        AddTestCase(sheetData, row++, "IMCSCH", "=IMCSCH(\"1+i\")", "Csch of 1+i");

        // IMSINH tests
        AddTestCase(sheetData, row++, "IMSINH", "=IMSINH(\"0\")", "Sinh of 0");
        AddTestCase(sheetData, row++, "IMSINH", "=IMSINH(\"1+i\")", "Sinh of 1+i");

        // IMCOSH tests
        AddTestCase(sheetData, row++, "IMCOSH", "=IMCOSH(\"0\")", "Cosh of 0");
        AddTestCase(sheetData, row++, "IMCOSH", "=IMCOSH(\"1+i\")", "Cosh of 1+i");

        // BITAND tests
        AddTestCase(sheetData, row++, "BITAND", "=BITAND(5, 3)", "Bitwise AND 5 and 3");
        AddTestCase(sheetData, row++, "BITAND", "=BITAND(15, 7)", "Bitwise AND 15 and 7");

        // BITOR tests
        AddTestCase(sheetData, row++, "BITOR", "=BITOR(5, 3)", "Bitwise OR 5 and 3");
        AddTestCase(sheetData, row++, "BITOR", "=BITOR(8, 4)", "Bitwise OR 8 and 4");

        // BITXOR tests
        AddTestCase(sheetData, row++, "BITXOR", "=BITXOR(5, 3)", "Bitwise XOR 5 and 3");
        AddTestCase(sheetData, row++, "BITXOR", "=BITXOR(12, 10)", "Bitwise XOR 12 and 10");

        // BITLSHIFT tests
        AddTestCase(sheetData, row++, "BITLSHIFT", "=BITLSHIFT(3, 2)", "Bitwise left shift 3 by 2");
        AddTestCase(sheetData, row++, "BITLSHIFT", "=BITLSHIFT(5, 1)", "Bitwise left shift 5 by 1");

        // BITRSHIFT tests
        AddTestCase(sheetData, row++, "BITRSHIFT", "=BITRSHIFT(12, 2)", "Bitwise right shift 12 by 2");
        AddTestCase(sheetData, row++, "BITRSHIFT", "=BITRSHIFT(10, 1)", "Bitwise right shift 10 by 1");


        SortCellsInRows(sheetData);
        worksheetPart.Worksheet.Save();
    }

    private static void CreateDatabaseFunctionsSheet(WorkbookPart workbookPart, Sheets sheets, uint sheetId)
    {
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        worksheetPart.Worksheet = new Worksheet(sheetData);

        sheets.AppendChild(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = "Database",
        });

        // Create database table (A1:C5)
        AddTextCell(sheetData, "A1", "Name");
        AddTextCell(sheetData, "B1", "Dept");
        AddTextCell(sheetData, "C1", "Salary");

        AddTextCell(sheetData, "A2", "Alice");
        AddTextCell(sheetData, "B2", "Sales");
        AddCell(sheetData, "C2", "50000");

        AddTextCell(sheetData, "A3", "Bob");
        AddTextCell(sheetData, "B3", "IT");
        AddCell(sheetData, "C3", "60000");

        AddTextCell(sheetData, "A4", "Carol");
        AddTextCell(sheetData, "B4", "Sales");
        AddCell(sheetData, "C4", "55000");

        AddTextCell(sheetData, "A5", "Dave");
        AddTextCell(sheetData, "B5", "IT");
        AddCell(sheetData, "C5", "65000");

        // Criteria range (E1:E2)
        AddTextCell(sheetData, "E1", "Dept");
        AddTextCell(sheetData, "E2", "Sales");

        var row = 7;

        // DSUM tests
        AddTestCase(sheetData, row++, "DSUM", "=DSUM(A1:C5, \"Salary\", E1:E2)", "DSUM sales salaries");

        // DCOUNT tests
        AddTestCase(sheetData, row++, "DCOUNT", "=DCOUNT(A1:C5, \"Salary\", E1:E2)", "DCOUNT sales records");

        // DCOUNTA tests
        AddTestCase(sheetData, row++, "DCOUNTA", "=DCOUNTA(A1:C5, \"Name\", E1:E2)", "DCOUNTA sales names");

        // DAVERAGE tests
        AddTestCase(sheetData, row++, "DAVERAGE", "=DAVERAGE(A1:C5, \"Salary\", E1:E2)", "DAVERAGE sales salary");

        // DMAX tests
        AddTestCase(sheetData, row++, "DMAX", "=DMAX(A1:C5, \"Salary\", E1:E2)", "DMAX sales salary");

        // DMIN tests
        AddTestCase(sheetData, row++, "DMIN", "=DMIN(A1:C5, \"Salary\", E1:E2)", "DMIN sales salary");

        SortCellsInRows(sheetData);
        worksheetPart.Worksheet.Save();
    }

    private static void CreateInformationFunctionsSheet(WorkbookPart workbookPart, Sheets sheets, uint sheetId)
    {
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        worksheetPart.Worksheet = new Worksheet(sheetData);

        sheets.AppendChild(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = "Information",
        });

        AddHeader(sheetData, 1);
        var row = 2;

        // Setup test data in column F (avoid conflict with test case columns A-D)
        AddCell(sheetData, "F1", "123");
        AddTextCell(sheetData, "F2", "Hello");

        // ISNUMBER tests
        AddTestCase(sheetData, row++, "ISNUMBER", "=ISNUMBER(123)", "Is number literal");
        AddTestCase(sheetData, row++, "ISNUMBER", "=ISNUMBER(\"text\")", "Is number text");
        AddTestCase(sheetData, row++, "ISNUMBER", "=ISNUMBER(F1)", "Is number cell");

        // ISTEXT tests
        AddTestCase(sheetData, row++, "ISTEXT", "=ISTEXT(\"Hello\")", "Is text literal");
        AddTestCase(sheetData, row++, "ISTEXT", "=ISTEXT(123)", "Is text number");
        AddTestCase(sheetData, row++, "ISTEXT", "=ISTEXT(F2)", "Is text cell");

        // ISEVEN tests
        AddTestCase(sheetData, row++, "ISEVEN", "=ISEVEN(2)", "Is even 2");
        AddTestCase(sheetData, row++, "ISEVEN", "=ISEVEN(3)", "Is even 3");
        AddTestCase(sheetData, row++, "ISEVEN", "=ISEVEN(0)", "Is even 0");

        // ISODD tests
        AddTestCase(sheetData, row++, "ISODD", "=ISODD(3)", "Is odd 3");
        AddTestCase(sheetData, row++, "ISODD", "=ISODD(2)", "Is odd 2");
        AddTestCase(sheetData, row++, "ISODD", "=ISODD(1)", "Is odd 1");

        // ISLOGICAL tests
        AddTestCase(sheetData, row++, "ISLOGICAL", "=ISLOGICAL(TRUE)", "Is logical TRUE");
        AddTestCase(sheetData, row++, "ISLOGICAL", "=ISLOGICAL(FALSE)", "Is logical FALSE");
        AddTestCase(sheetData, row++, "ISLOGICAL", "=ISLOGICAL(1)", "Is logical number");

        // ISNONTEXT tests
        AddTestCase(sheetData, row++, "ISNONTEXT", "=ISNONTEXT(123)", "Is nontext number");
        AddTestCase(sheetData, row++, "ISNONTEXT", "=ISNONTEXT(\"Hello\")", "Is nontext text");
        AddTestCase(sheetData, row++, "ISNONTEXT", "=ISNONTEXT(TRUE)", "Is nontext logical");

        // TYPE tests
        AddTestCase(sheetData, row++, "TYPE", "=TYPE(123)", "Type of number");
        AddTestCase(sheetData, row++, "TYPE", "=TYPE(\"Hello\")", "Type of text");
        AddTestCase(sheetData, row++, "TYPE", "=TYPE(TRUE)", "Type of logical");

        // N tests
        AddTestCase(sheetData, row++, "N", "=N(123)", "N of number");
        AddTestCase(sheetData, row++, "N", "=N(\"Hello\")", "N of text");
        AddTestCase(sheetData, row++, "N", "=N(TRUE)", "N of TRUE");

        SortCellsInRows(sheetData);
        worksheetPart.Worksheet.Save();
    }

    private static void CreateErrorHandlingFunctionsSheet(WorkbookPart workbookPart, Sheets sheets, uint sheetId)
    {
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        worksheetPart.Worksheet = new Worksheet(sheetData);

        sheets.AppendChild(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = "ErrorHandling",
        });

        AddHeader(sheetData, 1);
        var row = 2;

        // IFERROR tests
        AddTestCase(sheetData, row++, "IFERROR", "=IFERROR(10/2, \"Error\")", "IFERROR valid");
        AddTestCase(sheetData, row++, "IFERROR", "=IFERROR(10/0, \"DivByZero\")", "IFERROR division by zero");
        AddTestCase(sheetData, row++, "IFERROR", "=IFERROR(VLOOKUP(99, F5:G7, 2, FALSE), \"Not Found\")", "IFERROR lookup fail", ("F5", "1"), ("G5", "A"), ("F6", "2"), ("G6", "B"), ("F7", "3"), ("G7", "C"));

        // ISERROR tests
        AddTestCase(sheetData, row++, "ISERROR", "=ISERROR(10/2)", "ISERROR valid");
        AddTestCase(sheetData, row++, "ISERROR", "=ISERROR(10/0)", "ISERROR division error");
        AddTestCase(sheetData, row++, "ISERROR", "=ISERROR(SQRT(-1))", "ISERROR sqrt negative");

        // ISNA tests
        AddTestCase(sheetData, row++, "ISNA", "=ISNA(10)", "ISNA value");
        AddTestCase(sheetData, row++, "ISNA", "=ISNA(VLOOKUP(99, F5:G7, 2, FALSE))", "ISNA lookup fail");

        // ISERR tests
        AddTestCase(sheetData, row++, "ISERR", "=ISERR(10/2)", "ISERR valid");
        AddTestCase(sheetData, row++, "ISERR", "=ISERR(10/0)", "ISERR division error");

        // ISBLANK tests
        AddCell(sheetData, "H1", "10");
        AddTextCell(sheetData, "H2", "");
        AddTestCase(sheetData, row++, "ISBLANK", "=ISBLANK(H1)", "ISBLANK not blank");
        AddTestCase(sheetData, row++, "ISBLANK", "=ISBLANK(H3)", "ISBLANK empty cell");

        SortCellsInRows(sheetData);
        worksheetPart.Worksheet.Save();
    }


    private static void CreateForecastingFunctionsSheet(WorkbookPart workbookPart, Sheets sheets, uint sheetId)
    {
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        worksheetPart.Worksheet = new Worksheet(sheetData);

        sheets.AppendChild(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = "Forecasting",
        });

        AddHeader(sheetData, 1);
        var row = 2;

        // Setup time series data for forecasting tests in high rows to avoid conflicts
        // Known Y values (sales data)
        AddCell(sheetData, "F100", "100");
        AddCell(sheetData, "F101", "110");
        AddCell(sheetData, "F102", "125");
        AddCell(sheetData, "F103", "140");
        AddCell(sheetData, "F104", "160");

        // Known X values (time periods)
        AddCell(sheetData, "G100", "1");
        AddCell(sheetData, "G101", "2");
        AddCell(sheetData, "G102", "3");
        AddCell(sheetData, "G103", "4");
        AddCell(sheetData, "G104", "5");

        // Timeline dates for ETS functions (using Excel date serial numbers)
        AddCell(sheetData, "H100", "44927"); // 2023-01-01
        AddCell(sheetData, "H101", "44958"); // 2023-02-01
        AddCell(sheetData, "H102", "44986"); // 2023-03-01
        AddCell(sheetData, "H103", "45017"); // 2023-04-01
        AddCell(sheetData, "H104", "45047"); // 2023-05-01

        // FORECAST tests - Linear forecast
        AddTestCase(sheetData, row++, "FORECAST", "=ROUND(FORECAST(6, F100:F104, G100:G104), 2)", "Forecast for period 6");
        AddTestCase(sheetData, row++, "FORECAST", "=ROUND(FORECAST(7, F100:F104, G100:G104), 2)", "Forecast for period 7");
        AddTestCase(sheetData, row++, "FORECAST", "=ROUND(FORECAST(3.5, F100:F104, G100:G104), 2)", "Forecast interpolation");

        // FORECAST.LINEAR tests (same as FORECAST)
        AddTestCase(sheetData, row++, "FORECAST.LINEAR", "=ROUND(FORECAST.LINEAR(6, F100:F104, G100:G104), 2)", "Linear forecast period 6");
        AddTestCase(sheetData, row++, "FORECAST.LINEAR", "=ROUND(FORECAST.LINEAR(8, F100:F104, G100:G104), 2)", "Linear forecast period 8");

        // FORECAST.ETS tests - Exponential smoothing forecast
        AddTestCase(sheetData, row++, "FORECAST.ETS", "=ROUND(FORECAST.ETS(45078, F100:F104, H100:H104, 1, 1, 1), 2)", "ETS forecast next month");
        AddTestCase(sheetData, row++, "FORECAST.ETS", "=ROUND(FORECAST.ETS(45108, F100:F104, H100:H104), 2)", "ETS forecast 2 months");

        // FORECAST.ETS.CONFINT tests - Confidence interval
        AddTestCase(sheetData, row++, "FORECAST.ETS.CONFINT", "=ROUND(FORECAST.ETS.CONFINT(45078, F100:F104, H100:H104, 0.95, 1, 1, 1), 2)", "ETS confidence interval 95%");
        AddTestCase(sheetData, row++, "FORECAST.ETS.CONFINT", "=ROUND(FORECAST.ETS.CONFINT(45078, F100:F104, H100:H104, 0.90), 2)", "ETS confidence interval 90%");

        // FORECAST.ETS.SEASONALITY tests - Detect seasonality
        AddTestCase(sheetData, row++, "FORECAST.ETS.SEASONALITY", "=FORECAST.ETS.SEASONALITY(F100:F104, H100:H104, 1, 1)", "ETS seasonality detection");
        AddTestCase(sheetData, row++, "FORECAST.ETS.SEASONALITY", "=FORECAST.ETS.SEASONALITY(F100:F104, H100:H104)", "ETS seasonality auto");

        // FORECAST.ETS.STAT tests - Statistical metrics
        AddTestCase(sheetData, row++, "FORECAST.ETS.STAT", "=ROUND(FORECAST.ETS.STAT(F100:F104, H100:H104, 1, 1, 1, 1), 4)", "ETS alpha parameter");
        AddTestCase(sheetData, row++, "FORECAST.ETS.STAT", "=ROUND(FORECAST.ETS.STAT(F100:F104, H100:H104, 2), 4)", "ETS beta parameter");
        AddTestCase(sheetData, row++, "FORECAST.ETS.STAT", "=ROUND(FORECAST.ETS.STAT(F100:F104, H100:H104, 3), 4)", "ETS gamma parameter");
        AddTestCase(sheetData, row++, "FORECAST.ETS.STAT", "=ROUND(FORECAST.ETS.STAT(F100:F104, H100:H104, 8), 4)", "ETS MASE metric");

        // TREND tests - Linear trend extrapolation
        AddTestCase(sheetData, row++, "TREND", "=ROUND(INDEX(TREND(F100:F104, G100:G104, 6), 1), 2)", "Trend for new X value 6");
        AddTestCase(sheetData, row++, "TREND", "=ROUND(INDEX(TREND(F100:F104, G100:G104), 3), 2)", "Trend fitted value period 3");
        AddTestCase(sheetData, row++, "TREND", "=ROUND(INDEX(TREND(F100:F104), 4), 2)", "Trend auto X values");

        // GROWTH tests - Exponential growth
        AddTestCase(sheetData, row++, "GROWTH", "=ROUND(INDEX(GROWTH(F100:F104, G100:G104, 6), 1), 2)", "Growth for new X value 6");
        AddTestCase(sheetData, row++, "GROWTH", "=ROUND(INDEX(GROWTH(F100:F104, G100:G104), 3), 2)", "Growth fitted value period 3");
        AddTestCase(sheetData, row++, "GROWTH", "=ROUND(INDEX(GROWTH(F100:F104), 4), 2)", "Growth auto X values");

        // LINEST tests - Linear regression statistics
        AddTestCase(sheetData, row++, "LINEST", "=ROUND(INDEX(LINEST(F100:F104, G100:G104), 1), 4)", "LINEST slope");
        AddTestCase(sheetData, row++, "LINEST", "=ROUND(INDEX(LINEST(F100:F104, G100:G104), 2), 4)", "LINEST intercept");
        AddTestCase(sheetData, row++, "LINEST", "=ROUND(INDEX(LINEST(F100:F104, G100:G104, TRUE, TRUE), 1, 1), 4)", "LINEST slope with stats");
        AddTestCase(sheetData, row++, "LINEST", "=ROUND(INDEX(LINEST(F100:F104, G100:G104, FALSE), 1), 4)", "LINEST zero intercept");

        // LOGEST tests - Exponential regression statistics
        AddTestCase(sheetData, row++, "LOGEST", "=ROUND(INDEX(LOGEST(F100:F104, G100:G104), 1), 4)", "LOGEST base coefficient");
        AddTestCase(sheetData, row++, "LOGEST", "=ROUND(INDEX(LOGEST(F100:F104, G100:G104), 2), 4)", "LOGEST constant");
        AddTestCase(sheetData, row++, "LOGEST", "=ROUND(INDEX(LOGEST(F100:F104, G100:G104, TRUE, TRUE), 1, 1), 4)", "LOGEST with statistics");
        AddTestCase(sheetData, row++, "LOGEST", "=ROUND(INDEX(LOGEST(F100:F104, G100:G104, FALSE), 1), 4)", "LOGEST b=1");

        SortCellsInRows(sheetData);
        worksheetPart.Worksheet.Save();
    }

    private static void CreateCubeFunctionsSheet(WorkbookPart workbookPart, Sheets sheets, uint sheetId)
    {
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        worksheetPart.Worksheet = new Worksheet(sheetData);

        sheets.AppendChild(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = "Cube",
        });

        AddHeader(sheetData, 1);
        var row = 2;

        // Note: Cube functions require external OLAP connections
        // These test cases verify the functions return appropriate errors when connection is unavailable
        // Test data setup - connection strings (will fail without actual OLAP server)
        AddTextCell(sheetData, "F100", "Provider=MSOLAP;Data Source=localhost;Initial Catalog=Adventure Works DW");
        AddTextCell(sheetData, "F101", "[Measures].[Internet Sales Amount]");
        AddTextCell(sheetData, "F102", "[Product].[Product Categories].[Category].[Bikes]");
        AddTextCell(sheetData, "F103", "[Date].[Fiscal].[Fiscal Year].[FY 2023]");

        // CUBEVALUE tests - Retrieve aggregated value from cube
        AddTestCase(sheetData, row++, "CUBEVALUE", "=IFERROR(CUBEVALUE(F100, F101), \"#CONNECT_ERROR\")", "CUBEVALUE single measure", ("F100", "Provider=MSOLAP"), ("F101", "[Measures].[Sales]"));
        AddTestCase(sheetData, row++, "CUBEVALUE", "=IFERROR(CUBEVALUE(F100, F101, F102), \"#CONNECT_ERROR\")", "CUBEVALUE with dimension");
        AddTestCase(sheetData, row++, "CUBEVALUE", "=IFERROR(CUBEVALUE(F100, F101, F102, F103), \"#CONNECT_ERROR\")", "CUBEVALUE multiple dimensions");
        AddTestCase(sheetData, row++, "CUBEVALUE", "=ISERROR(CUBEVALUE(\"InvalidConn\", \"[Measures].[Sales]\"))", "CUBEVALUE connection error check");

        // CUBEMEMBER tests - Define member or tuple from cube
        AddTestCase(sheetData, row++, "CUBEMEMBER", "=IFERROR(CUBEMEMBER(F100, F102), \"#CONNECT_ERROR\")", "CUBEMEMBER basic");
        AddTestCase(sheetData, row++, "CUBEMEMBER", "=IFERROR(CUBEMEMBER(F100, F102, \"Bikes Category\"), \"#CONNECT_ERROR\")", "CUBEMEMBER with caption");
        AddTestCase(sheetData, row++, "CUBEMEMBER", "=IFERROR(CUBEMEMBER(F100, \"[Product].[All Products]\"), \"#CONNECT_ERROR\")", "CUBEMEMBER all products");
        AddTestCase(sheetData, row++, "CUBEMEMBER", "=ISERROR(CUBEMEMBER(\"InvalidConn\", F102))", "CUBEMEMBER error validation");

        // CUBEMEMBERPROPERTY tests - Return property value of member
        AddTestCase(sheetData, row++, "CUBEMEMBERPROPERTY", "=IFERROR(CUBEMEMBERPROPERTY(F100, F102, \"MEMBER_CAPTION\"), \"#CONNECT_ERROR\")", "CUBEMEMBERPROPERTY caption");
        AddTestCase(sheetData, row++, "CUBEMEMBERPROPERTY", "=IFERROR(CUBEMEMBERPROPERTY(F100, F102, \"MEMBER_UNIQUE_NAME\"), \"#CONNECT_ERROR\")", "CUBEMEMBERPROPERTY unique name");
        AddTestCase(sheetData, row++, "CUBEMEMBERPROPERTY", "=IFERROR(CUBEMEMBERPROPERTY(F100, F102, \"LEVEL_NUMBER\"), \"#CONNECT_ERROR\")", "CUBEMEMBERPROPERTY level number");
        AddTestCase(sheetData, row++, "CUBEMEMBERPROPERTY", "=ISERROR(CUBEMEMBERPROPERTY(\"InvalidConn\", F102, \"CAPTION\"))", "CUBEMEMBERPROPERTY error check");

        // CUBERANKEDMEMBER tests - Return nth member in set
        AddTestCase(sheetData, row++, "CUBERANKEDMEMBER", "=IFERROR(CUBERANKEDMEMBER(F100, \"[Product].[Product].Members\", 1), \"#CONNECT_ERROR\")", "CUBERANKEDMEMBER 1st member");
        AddTestCase(sheetData, row++, "CUBERANKEDMEMBER", "=IFERROR(CUBERANKEDMEMBER(F100, \"[Product].[Product].Members\", 5, \"Top 5\"), \"#CONNECT_ERROR\")", "CUBERANKEDMEMBER with caption");
        AddTestCase(sheetData, row++, "CUBERANKEDMEMBER", "=IFERROR(CUBERANKEDMEMBER(F100, \"[Date].[Calendar].Members\", 10), \"#CONNECT_ERROR\")", "CUBERANKEDMEMBER 10th member");
        AddTestCase(sheetData, row++, "CUBERANKEDMEMBER", "=ISERROR(CUBERANKEDMEMBER(\"InvalidConn\", \"[Product].Members\", 1))", "CUBERANKEDMEMBER error check");

        // CUBESET tests - Define calculated set of members
        AddTestCase(sheetData, row++, "CUBESET", "=IFERROR(CUBESET(F100, \"[Product].[Category].Members\"), \"#CONNECT_ERROR\")", "CUBESET basic");
        AddTestCase(sheetData, row++, "CUBESET", "=IFERROR(CUBESET(F100, \"[Product].[Category].Members\", \"All Categories\"), \"#CONNECT_ERROR\")", "CUBESET with caption");
        AddTestCase(sheetData, row++, "CUBESET", "=IFERROR(CUBESET(F100, \"TopCount([Product].[Product].Members, 10)\", \"Top 10\", 1), \"#CONNECT_ERROR\")", "CUBESET sorted ascending");
        AddTestCase(sheetData, row++, "CUBESET", "=IFERROR(CUBESET(F100, \"[Date].[Fiscal].Members\", \"Dates\", 2, \"[Measures].[Sales]\"), \"#CONNECT_ERROR\")", "CUBESET sorted descending");
        AddTestCase(sheetData, row++, "CUBESET", "=ISERROR(CUBESET(\"InvalidConn\", \"[Product].Members\"))", "CUBESET error validation");

        // CUBESETCOUNT tests - Return number of items in set
        AddTextCell(sheetData, "G100", "#CONNECT_ERROR");
        AddTestCase(sheetData, row++, "CUBESETCOUNT", "=IFERROR(CUBESETCOUNT(CUBESET(F100, \"[Product].[Category].Members\")), \"#CONNECT_ERROR\")", "CUBESETCOUNT basic");
        AddTestCase(sheetData, row++, "CUBESETCOUNT", "=IF(G100=\"#CONNECT_ERROR\", \"#CONNECT_ERROR\", CUBESETCOUNT(G100))", "CUBESETCOUNT from cell reference");
        AddTestCase(sheetData, row++, "CUBESETCOUNT", "=IFERROR(CUBESETCOUNT(CUBESET(F100, \"[Date].[Calendar].Members\")), \"#CONNECT_ERROR\")", "CUBESETCOUNT calendar");
        AddTestCase(sheetData, row++, "CUBESETCOUNT", "=ISERROR(CUBESETCOUNT(\"InvalidSet\"))", "CUBESETCOUNT error check");

        // CUBEKPIMEMBER tests - Return KPI property
        AddTestCase(sheetData, row++, "CUBEKPIMEMBER", "=IFERROR(CUBEKPIMEMBER(F100, \"Sales Growth\", 1), \"#CONNECT_ERROR\")", "CUBEKPIMEMBER value");
        AddTestCase(sheetData, row++, "CUBEKPIMEMBER", "=IFERROR(CUBEKPIMEMBER(F100, \"Sales Growth\", 2), \"#CONNECT_ERROR\")", "CUBEKPIMEMBER goal");
        AddTestCase(sheetData, row++, "CUBEKPIMEMBER", "=IFERROR(CUBEKPIMEMBER(F100, \"Sales Growth\", 3), \"#CONNECT_ERROR\")", "CUBEKPIMEMBER status");
        AddTestCase(sheetData, row++, "CUBEKPIMEMBER", "=IFERROR(CUBEKPIMEMBER(F100, \"Revenue KPI\", 4, \"Trend\"), \"#CONNECT_ERROR\")", "CUBEKPIMEMBER trend with caption");
        AddTestCase(sheetData, row++, "CUBEKPIMEMBER", "=IFERROR(CUBEKPIMEMBER(F100, \"Profit Margin\", 5), \"#CONNECT_ERROR\")", "CUBEKPIMEMBER weight");
        AddTestCase(sheetData, row++, "CUBEKPIMEMBER", "=ISERROR(CUBEKPIMEMBER(\"InvalidConn\", \"KPI\", 1))", "CUBEKPIMEMBER error check");

        SortCellsInRows(sheetData);
        worksheetPart.Worksheet.Save();
    }

    private static void AddHeader(SheetData sheetData, uint rowIndex)
    {
        var row = new Row { RowIndex = rowIndex };
        sheetData.AppendChild(row);

        row.AppendChild(new Cell
        {
            CellReference = $"A{rowIndex}",
            DataType = CellValues.String,
            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("Function"),
        });

        row.AppendChild(new Cell
        {
            CellReference = $"B{rowIndex}",
            DataType = CellValues.String,
            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("Formula"),
        });

        row.AppendChild(new Cell
        {
            CellReference = $"C{rowIndex}",
            DataType = CellValues.String,
            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("Expected (Excel)"),
        });

        row.AppendChild(new Cell
        {
            CellReference = $"D{rowIndex}",
            DataType = CellValues.String,
            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("Description"),
        });
    }

    private static void AddTestCase(SheetData sheetData, int rowIndex, string functionName, string formula, string description, params (string CellRef, string Value)[] setupData)
    {
        // Add setup data first
        foreach (var (cellRef, value) in setupData)
        {
            AddCell(sheetData, cellRef, value);
        }

        // Get or create the row for this test case
        var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == (uint)rowIndex);
        if (row == null)
        {
            row = new Row { RowIndex = (uint)rowIndex };
            sheetData.AppendChild(row);
        }

        // Column A: Function name
        row.AppendChild(new Cell
        {
            CellReference = $"A{rowIndex}",
            DataType = CellValues.String,
            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(functionName),
        });

        // Column B: Formula (as text for reference)
        row.AppendChild(new Cell
        {
            CellReference = $"B{rowIndex}",
            DataType = CellValues.String,
            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(formula),
        });

        // Column C: Formula (actual formula for Excel to calculate)
        row.AppendChild(new Cell
        {
            CellReference = $"C{rowIndex}",
            CellFormula = new CellFormula(formula),
            // No CellValue - Excel will calculate
        });

        // Column D: Description
        row.AppendChild(new Cell
        {
            CellReference = $"D{rowIndex}",
            DataType = CellValues.String,
            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(description),
        });
    }

    private static void AddCell(SheetData sheetData, string cellRef, string value)
    {
        var rowIndex = GetRowIndex(cellRef);
        var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
        if (row == null)
        {
            row = new Row { RowIndex = rowIndex };
            sheetData.AppendChild(row);
        }

        row.AppendChild(new Cell
        {
            CellReference = cellRef,
            DataType = CellValues.Number,
            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(value),
        });
    }

    private static void AddTextCell(SheetData sheetData, string cellRef, string value)
    {
        var rowIndex = GetRowIndex(cellRef);
        var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
        if (row == null)
        {
            row = new Row { RowIndex = rowIndex };
            sheetData.AppendChild(row);
        }

        row.AppendChild(new Cell
        {
            CellReference = cellRef,
            DataType = CellValues.String,
            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(value),
        });
    }

    private static uint GetRowIndex(string reference)
    {
        return uint.Parse(new string(reference.Where(char.IsDigit).ToArray()), CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Sorts cells within each row by cell reference (required by Excel).
    /// Excel requires cells to be in alphabetical order within a row (A, B, C... not C, A, B).
    /// </summary>
    private static void SortCellsInRows(SheetData sheetData)
    {
        foreach (var row in sheetData.Elements<Row>())
        {
            var cells = row.Elements<Cell>().ToList();

            // Sort by column index (extracted from cell reference like "A1" -> 1, "B1" -> 2)
            var sortedCells = cells.OrderBy(c =>
            {
                var cellRef = c.CellReference?.Value ?? string.Empty;
                return GetColumnIndex(cellRef);
            }).ToList();

            // Remove all cells
            row.RemoveAllChildren<Cell>();

            // Add back in sorted order
            foreach (var cell in sortedCells)
            {
                row.AppendChild(cell);
            }
        }
    }

    private static int GetColumnIndex(string cellReference)
    {
        if (string.IsNullOrEmpty(cellReference))
        {
            return 0;
        }

        int index = 0;
        foreach (char c in cellReference)
        {
            if (char.IsLetter(c))
            {
                index = (index * 26) + (char.ToUpperInvariant(c) - 'A' + 1);
            }
            else
            {
                break; // Stop at the first digit
            }
        }

        return index;
    }
}
