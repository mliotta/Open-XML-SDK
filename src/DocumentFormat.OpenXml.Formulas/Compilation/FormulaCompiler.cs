// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Parsing;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

/// <summary>
/// Compiles formula ASTs into executable lambda expressions.
/// </summary>
public class FormulaCompiler
{
    private readonly ParameterExpression _ctxParam = Expression.Parameter(typeof(CellContext), "ctx");

    /// <summary>
    /// Compiles a formula AST into a lambda expression.
    /// </summary>
    /// <param name="ast">The formula AST.</param>
    /// <returns>A compiled lambda expression.</returns>
    public Expression<Func<CellContext, CellValue>> Compile(FormulaNode ast)
    {
        var body = CompileNode(ast);
        return Expression.Lambda<Func<CellContext, CellValue>>(body, _ctxParam);
    }

    private Expression CompileNode(FormulaNode node)
    {
        return node switch
        {
            BinaryOpNode bin => CompileBinaryOp(bin),
            UnaryOpNode un => CompileUnaryOp(un),
            FunctionCallNode func => CompileFunction(func),
            CellReferenceNode cell => CompileCellRef(cell),
            SheetReferenceNode sheet => CompileSheetRef(sheet),
            RangeNode range => throw new CompilationException("Range nodes must be inside function calls"),
            LiteralNode lit => CompileLiteral(lit),
            _ => throw new CompilationException($"Unsupported node type: {node.GetType().Name}"),
        };
    }

    private Expression CompileLiteral(LiteralNode node)
    {
        if (node.Value is double d)
        {
            return Expression.Constant(CellValue.FromNumber(d));
        }

        if (node.Value is string s)
        {
            return Expression.Constant(CellValue.FromString(s));
        }

        if (node.Value is bool b)
        {
            return Expression.Constant(CellValue.FromBool(b));
        }

        if (node.Value is CellValue cv)
        {
            return Expression.Constant(cv);
        }

        throw new CompilationException($"Unsupported literal type: {node.Value.GetType().Name}");
    }

    private Expression CompileBinaryOp(BinaryOpNode node)
    {
        var left = CompileNode(node.Left);
        var right = CompileNode(node.Right);

        // Handle string concatenation separately
        if (node.Operator == BinaryOperator.Concat)
        {
            return CompileConcat(left, right);
        }

        // Get NumericValue from CellValue
        var leftValue = Expression.Property(left, nameof(CellValue.NumericValue));
        var rightValue = Expression.Property(right, nameof(CellValue.NumericValue));

        Expression result = node.Operator switch
        {
            BinaryOperator.Add => Expression.Add(leftValue, rightValue),
            BinaryOperator.Subtract => Expression.Subtract(leftValue, rightValue),
            BinaryOperator.Multiply => Expression.Multiply(leftValue, rightValue),
            BinaryOperator.Divide => CompileSafeDivide(leftValue, rightValue),
            BinaryOperator.Power => CompilePower(leftValue, rightValue),
            BinaryOperator.GreaterThan => Expression.GreaterThan(leftValue, rightValue),
            BinaryOperator.LessThan => Expression.LessThan(leftValue, rightValue),
            BinaryOperator.GreaterThanOrEqual => Expression.GreaterThanOrEqual(leftValue, rightValue),
            BinaryOperator.LessThanOrEqual => Expression.LessThanOrEqual(leftValue, rightValue),
            BinaryOperator.Equals => Expression.Equal(leftValue, rightValue),
            BinaryOperator.NotEqual => Expression.NotEqual(leftValue, rightValue),
            _ => throw new CompilationException($"Unsupported operator: {node.Operator}"),
        };

        // For comparison operators, convert bool to CellValue
        if (node.Operator == BinaryOperator.GreaterThan ||
            node.Operator == BinaryOperator.LessThan ||
            node.Operator == BinaryOperator.GreaterThanOrEqual ||
            node.Operator == BinaryOperator.LessThanOrEqual ||
            node.Operator == BinaryOperator.Equals ||
            node.Operator == BinaryOperator.NotEqual)
        {
            var fromBoolMethod = typeof(CellValue).GetMethod(nameof(CellValue.FromBool), BindingFlags.Public | BindingFlags.Static);
            if (fromBoolMethod == null)
            {
                throw new CompilationException($"Method {nameof(CellValue.FromBool)} not found");
            }

            return Expression.Call(fromBoolMethod, result);
        }

        // For arithmetic operators (except divide, which is already wrapped), convert double to CellValue
        if (node.Operator == BinaryOperator.Divide)
        {
            // Already returns CellValue from CompileSafeDivide
            return result;
        }

        var fromNumberMethod = typeof(CellValue).GetMethod(nameof(CellValue.FromNumber), BindingFlags.Public | BindingFlags.Static);
        if (fromNumberMethod == null)
        {
            throw new CompilationException($"Method {nameof(CellValue.FromNumber)} not found");
        }

        return Expression.Call(fromNumberMethod, result);
    }

    private Expression CompileSafeDivide(Expression left, Expression right)
    {
        // if (right == 0) return Error("#DIV/0!") else return CellValue.FromNumber(left / right)
        var zero = Expression.Constant(0.0);

        var errorMethod = typeof(CellValue).GetMethod(nameof(CellValue.Error), BindingFlags.Public | BindingFlags.Static);
        if (errorMethod == null)
        {
            throw new CompilationException($"Method {nameof(CellValue.Error)} not found");
        }

        var divByZeroError = Expression.Call(errorMethod, Expression.Constant("#DIV/0!"));

        var division = Expression.Divide(left, right);

        var fromNumberMethod = typeof(CellValue).GetMethod(nameof(CellValue.FromNumber), BindingFlags.Public | BindingFlags.Static);
        if (fromNumberMethod == null)
        {
            throw new CompilationException($"Method {nameof(CellValue.FromNumber)} not found");
        }

        var divisionResult = Expression.Call(fromNumberMethod, division);

        return Expression.Condition(
            Expression.Equal(right, zero),
            divByZeroError,
            divisionResult);
    }

    private Expression CompileCellRef(CellReferenceNode node)
    {
        // ctx.GetCell("A1")
        var getCellMethod = typeof(CellContext).GetMethod(nameof(CellContext.GetCell), BindingFlags.Public | BindingFlags.Instance);
        if (getCellMethod == null)
        {
            throw new CompilationException($"Method {nameof(CellContext.GetCell)} not found");
        }

        return Expression.Call(_ctxParam, getCellMethod, Expression.Constant(node.Reference));
    }

    private Expression CompileFunction(FunctionCallNode node)
    {
        // Do lookup at COMPILE TIME
        if (!FunctionRegistry.TryGetFunction(node.FunctionName, out var function))
        {
            throw new UnsupportedFunctionException(node.FunctionName);
        }

        // Special handling for functions that take ranges
        var compiledArgs = new List<Expression>();

        foreach (var arg in node.Arguments)
        {
            if (arg is RangeNode range)
            {
                // Expand range to array of CellValues
                var getRangeMethod = typeof(CellContext).GetMethod(nameof(CellContext.GetRange), BindingFlags.Public | BindingFlags.Instance);
                if (getRangeMethod == null)
                {
                    throw new CompilationException($"Method {nameof(CellContext.GetRange)} not found");
                }

                var rangeExpr = Expression.Call(
                    _ctxParam,
                    getRangeMethod,
                    Expression.Constant(range.Start),
                    Expression.Constant(range.End));

                // Convert IEnumerable<CellValue> to CellValue[]
                var toArrayMethod = typeof(Enumerable).GetMethod(nameof(Enumerable.ToArray));
                if (toArrayMethod == null)
                {
                    throw new CompilationException($"Method {nameof(Enumerable.ToArray)} not found");
                }

                var arrayExpr = Expression.Call(toArrayMethod.MakeGenericMethod(typeof(CellValue)), rangeExpr);
                compiledArgs.Add(arrayExpr);
            }
            else
            {
                // Single value - wrap in array
                var valueExpr = CompileNode(arg);
                var arrayExpr = Expression.NewArrayInit(typeof(CellValue), valueExpr);
                compiledArgs.Add(arrayExpr);
            }
        }

        // Combine all argument arrays into a single array
        Expression argsArray;
        if (compiledArgs.Count == 0)
        {
            argsArray = Expression.NewArrayInit(typeof(CellValue));
        }
        else if (compiledArgs.Count == 1 && node.Arguments[0] is RangeNode)
        {
            // Single range argument - use as-is
            argsArray = compiledArgs[0];
        }
        else
        {
            // Multiple arguments or mix - concatenate arrays
            argsArray = ConcatenateArrays(compiledArgs);
        }

        // Get the static function reference
        var funcConstant = Expression.Constant(function);
        var executeMethod = typeof(IFunctionImplementation).GetMethod(nameof(IFunctionImplementation.Execute));
        if (executeMethod == null)
        {
            throw new CompilationException($"Method {nameof(IFunctionImplementation.Execute)} not found");
        }

        // Call: function.Execute(ctx, args)
        return Expression.Call(funcConstant, executeMethod, _ctxParam, argsArray);
    }

    private static Expression ConcatenateArrays(List<Expression> arrays)
    {
        // Call runtime helper that does single allocation + Array.Copy
        // This avoids LINQ Concat's intermediate enumerables and works on all target frameworks
        var arrayOfArraysExpr = Expression.NewArrayInit(typeof(CellValue[]), arrays);

        var helperMethod = typeof(FormulaCompiler).GetMethod(
            nameof(ConcatenateArraysAtRuntime),
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);

        if (helperMethod == null)
        {
            throw new CompilationException($"Method {nameof(ConcatenateArraysAtRuntime)} not found");
        }

        return Expression.Call(helperMethod, arrayOfArraysExpr);
    }

    /// <summary>
    /// Runtime helper that efficiently concatenates arrays with single allocation.
    /// </summary>
    private static CellValue[] ConcatenateArraysAtRuntime(params CellValue[][] arrays)
    {
        // Calculate total length
        var totalLength = 0;
        foreach (var arr in arrays)
        {
            totalLength += arr.Length;
        }

        // Allocate result array once
        var result = new CellValue[totalLength];

        // Copy all arrays using Array.Copy
        var offset = 0;
        foreach (var arr in arrays)
        {
            Array.Copy(arr, 0, result, offset, arr.Length);
            offset += arr.Length;
        }

        return result;
    }

    private Expression CompileUnaryOp(UnaryOpNode node)
    {
        var operand = CompileNode(node.Operand);

        return node.Operator switch
        {
            UnaryOperator.Negate => CompileNegate(operand),
            UnaryOperator.Plus => operand, // Unary plus is a no-op
            UnaryOperator.Percent => CompilePercent(operand),
            _ => throw new CompilationException($"Unsupported unary operator: {node.Operator}"),
        };
    }

    private Expression CompileNegate(Expression operand)
    {
        var numValue = Expression.Property(operand, nameof(CellValue.NumericValue));
        var negated = Expression.Negate(numValue);

        var fromNumberMethod = typeof(CellValue).GetMethod(nameof(CellValue.FromNumber), BindingFlags.Public | BindingFlags.Static);
        if (fromNumberMethod == null)
        {
            throw new CompilationException($"Method {nameof(CellValue.FromNumber)} not found");
        }

        return Expression.Call(fromNumberMethod, negated);
    }

    private Expression CompilePercent(Expression operand)
    {
        var numValue = Expression.Property(operand, nameof(CellValue.NumericValue));
        var hundred = Expression.Constant(100.0);
        var percent = Expression.Divide(numValue, hundred);

        var fromNumberMethod = typeof(CellValue).GetMethod(nameof(CellValue.FromNumber), BindingFlags.Public | BindingFlags.Static);
        if (fromNumberMethod == null)
        {
            throw new CompilationException($"Method {nameof(CellValue.FromNumber)} not found");
        }

        return Expression.Call(fromNumberMethod, percent);
    }

    private Expression CompilePower(Expression left, Expression right)
    {
        var powMethod = typeof(System.Math).GetMethod(nameof(System.Math.Pow), new[] { typeof(double), typeof(double) });
        if (powMethod == null)
        {
            throw new CompilationException($"Method {nameof(System.Math.Pow)} not found");
        }

        var power = Expression.Call(powMethod, left, right);

        var fromNumberMethod = typeof(CellValue).GetMethod(nameof(CellValue.FromNumber), BindingFlags.Public | BindingFlags.Static);
        if (fromNumberMethod == null)
        {
            throw new CompilationException($"Method {nameof(CellValue.FromNumber)} not found");
        }

        return Expression.Call(fromNumberMethod, power);
    }

    private Expression CompileConcat(Expression left, Expression right)
    {
        // Get StringValue from both CellValues
        var leftString = Expression.Property(left, nameof(CellValue.StringValue));
        var rightString = Expression.Property(right, nameof(CellValue.StringValue));

        var concatMethod = typeof(string).GetMethod(nameof(string.Concat), new[] { typeof(string), typeof(string) });
        if (concatMethod == null)
        {
            throw new CompilationException($"Method {nameof(string.Concat)} not found");
        }

        var concatenated = Expression.Call(concatMethod, leftString, rightString);

        var fromStringMethod = typeof(CellValue).GetMethod(nameof(CellValue.FromString), BindingFlags.Public | BindingFlags.Static);
        if (fromStringMethod == null)
        {
            throw new CompilationException($"Method {nameof(CellValue.FromString)} not found");
        }

        return Expression.Call(fromStringMethod, concatenated);
    }

    private Expression CompileSheetRef(SheetReferenceNode node)
    {
        // For now, throw exception - cross-sheet references require access to other worksheets
        // This will be implemented in Phase 2
        throw new CompilationException($"Cross-sheet references are not yet supported: {node.SheetName}!{node.CellReference}");
    }
}
