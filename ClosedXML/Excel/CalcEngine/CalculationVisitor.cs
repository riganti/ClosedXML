﻿using ClosedXML.Excel.CalcEngine.Exceptions;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using AnyValue = OneOf.OneOf<ClosedXML.Excel.CalcEngine.Logical, ClosedXML.Excel.CalcEngine.Number1, ClosedXML.Excel.CalcEngine.Text, ClosedXML.Excel.CalcEngine.Error1, ClosedXML.Excel.CalcEngine.Array, ClosedXML.Excel.CalcEngine.Reference>;

namespace ClosedXML.Excel.CalcEngine
{
    internal class CalculationVisitor : IFormulaVisitor<CalcContext, AnyValue>
    {
        private readonly FunctionRegistry _functions;

        public CalculationVisitor(FunctionRegistry functions)
        {
            _functions = functions;
        }

        public AnyValue Visit(CalcContext context, ScalarNode node)
        {
            return node.Value;
        }

        public AnyValue Visit(CalcContext context, ErrorExpression node)
        {
            return new Error1(node.ErrorType);
        }

        public AnyValue Visit(CalcContext context, UnaryExpression node)
        {
            var arg = node.Expression.Accept(context, this);
            if (context.UseImplicitIntersection)
            {
                arg = arg.ImplicitIntersection(context);
            }

            return node.Operation switch
            {
                UnaryOp.Add => arg.UnaryPlus(),
                UnaryOp.Subtract => arg.UnaryMinus(context),
                UnaryOp.Percentage => arg.UnaryPercent(context),
                UnaryOp.SpillRange => throw new NotImplementedException(),
                UnaryOp.ImplicitIntersection => throw new NotImplementedException("Excel 2016 implicit intersection is different from @ intersection of E2019+"),
                _ => throw new NotSupportedException($"Unknown operator {node.Operation}.")
            };
        }

        public AnyValue Visit(CalcContext context, BinaryExpression node)
        {
            var leftArg = node.LeftExpression.Accept(context, this);
            var rightArg = node.RightExpression.Accept(context, this);

            switch (node.Operation)
            {
                case BinaryOp.Range: return leftArg.ReferenceRange(rightArg);
                case BinaryOp.Union: return leftArg.ReferenceUnion(rightArg);
                case BinaryOp.Intersection: throw new NotImplementedException();
            };

            if (context.UseImplicitIntersection)
            {
                leftArg = leftArg.ImplicitIntersection(context);
                rightArg = rightArg.ImplicitIntersection(context);
            }

            return node.Operation switch
            {
                BinaryOp.Concat => leftArg.Concat(rightArg, context),
                BinaryOp.Add => leftArg.BinaryPlus(rightArg, context),
                BinaryOp.Sub => leftArg.BinaryMinus(rightArg, context),
                BinaryOp.Mult => leftArg.BinaryMult(rightArg, context),
                BinaryOp.Div => leftArg.BinaryDiv(rightArg, context),
                BinaryOp.Exp => leftArg.BinaryExp(rightArg, context),
                BinaryOp.Lt => leftArg.IsLessThan(rightArg, context),
                BinaryOp.Lte => leftArg.IsLessThanOrEqual(rightArg, context),
                BinaryOp.Eq => leftArg.IsEqual(rightArg, context),
                BinaryOp.Neq => leftArg.IsNotEqual(rightArg, context),
                BinaryOp.Gte => leftArg.IsGreaterThanOrEqual(rightArg, context),
                BinaryOp.Gt => leftArg.IsGreaterThan(rightArg, context),
                _ => throw new NotSupportedException($"Unknown operator {node.Operation}.")
            };
        }

        // Value is which args can be range, from 0. null - all args
        private static readonly Dictionary<string, IReadOnlyList<int>> rangeFunctions = new(StringComparer.OrdinalIgnoreCase)
        {
            { "AND" , null },
            { "NETWORKDAYS", new List<int> { 2 } },
            { "WORKDAY", new List<int> { 2 } },
            { "AVERAGE", null },
            { "AVERAGEA", null },
            { "CONCAT", null },
            { "CONCATENATE", null }, // TODO: Remove after switch to new engine. This function actually doesn't acccept ranges, but it's legacy implementation has a check and there is a test.
            { "COUNT", null },
            { "COUNTA", null },
            { "COUNTBLANK", null},
            { "COUNTIF", new List<int>{ 0 } },
            { "COUNTIFs", Enumerable.Range(0, 255).Where(x => x % 2 == 0).ToList() },
            { "DEVSQ", null },
            { "GEOMEAN", null },
            { "HLOOKUP", new List<int>{ 1 } },
            { "INDEX", new List<int> { 0, 1 } },
            { "MATCH", new List<int> { 1 } },
            { "MAX", null },
            { "MAXA", null },
            { "MEDIAN", null },
            { "MIN", null },
            { "MINA", null },
            { "MINVERSE", null },
            { "MMULT", null },
            { "SERIESSUM", new List<int>{ 3 } }, // Yay, this function is part of ECMA-376, but isn't in the list of functions that allow range.
            { "STDEV", null },
            { "STDEVA", null },
            { "STDEVP", null },
            { "STDEVPA", null },
            { "SUM", null },
            { "SUMIF", new List<int> { 0, 2 } },
            { "SUMIFS", new List<int> { 0,1,3,5,7,9} },
            { "SUMPRODUCT", null },
            { "TEXTJOIN", Enumerable.Range(2, 253).ToList() },
            { "VAR", null },
            { "VARA", null },
            { "VARP", null },
            { "VARPA", null },
            { "VLOOKUP", new List<int>{ 1 } },
        };


        public AnyValue Visit(CalcContext context, FunctionExpression node)
        {
            var args = new AnyValue?[node.Parameters.Count];
            for (var i = 0; i < node.Parameters.Count; ++i)
            {
                var paramNode = node.Parameters[i];
                if (paramNode is not EmptyArgumentNode)
                    args[i] = node.Parameters[i].Accept(context, this);
                else
                    args[i] = null;
            }

            var hasRangeParam = rangeFunctions.TryGetValue(node.Name, out var rangeParamIdx);
            for (var i = 0; i < args.Length; ++i)
            {
                if (!hasRangeParam || (rangeParamIdx is not null && !rangeParamIdx.Contains(i)))
                {
                    args[i] = args[i]?.ImplicitIntersection(context);
                }
            }

            if (!_functions.TryGetFunc(node.Name, out FormulaFunction function))
            {
                if (!_functions.TryGetFunc(node.Name, out FunctionDefinition legacyFunction))
                    return Error1.Name;

                // This creates a some of overhead, but all legacy functions will be migrated in near future
                var adaptedArgs = new List<Expression>(args.Length);
                foreach (var arg in args)
                {
                    Expression adaptedArg = arg.HasValue ? arg.Value.Match(
                        logical => new Expression(logical.Value),
                        number => new Expression(number.Value),
                        text => new Expression(text.Value),
                        error => new Expression(error.Type),
                        array => throw new NotSupportedException("Legacy CalcEngine couldn't work with arrays and neither will adapter."),
                        range =>
                        {
                            if (range.Areas.Count != 1)
                                throw new NotSupportedException($"Legacy XObjectExpression could only work with single area, reference has {range.Areas.Count}.");

                            var area = range.Areas.Single();
                            if (area.Worksheet is not null)
                            {
                                return new XObjectExpression(new CellRangeReference(area.Worksheet.Range(area), context.CalcEngine));
                            }
                            else
                            {
                                return new XObjectExpression(new CellRangeReference(context.Worksheet.Range(area), context.CalcEngine));
                            }
                        })
                        : new EmptyValueExpression();
                    adaptedArgs.Add(adaptedArg);
                }
                try
                {
                    var result = legacyFunction.Function(adaptedArgs);
                    return result switch
                    {
                        bool logic => AnyValue.FromT0(new Logical(logic)),
                        double number => AnyValue.FromT1(new Number1(number)),
                        string text => AnyValue.FromT2(new Text(text)),
                        int number => AnyValue.FromT1(new Number1(number)), /* date mostly */
                        long number => AnyValue.FromT1(new Number1(number)),
                        DateTime date => AnyValue.FromT1(new Number1(date.ToOADate())),
                        TimeSpan time => AnyValue.FromT1(new Number1(time.ToSerialDateTime())),
                        double[,] array => AnyValue.FromT4(new NumberArray(array)),
                        _ => throw new NotImplementedException($"Got a result from some function type {result?.GetType().Name ?? "null"} with value {result}.")
                    };
                }
                catch (CalcEngineException)
                {
                    //TODO: Map errors
                    throw;
                }
            }

            return function.CallFunction(context, args);
        }

        public AnyValue Visit(CalcContext context, ReferenceNode node)
        {
            XLWorksheet worksheet;
            if (node.Prefix is not null)
            {
                if (node.Prefix.File is not null)
                    throw new NotImplementedException("References from other files are not yet implemented.");

                if (node.Prefix.FirstSheet is not null || node.Prefix.LastSheet is not null)
                    throw new NotImplementedException("3D references are not yet implemented.");

                var sheet = node.Prefix.Sheet;
                if (!context.Workbook.TryGetWorksheet(sheet, out var worksheet1))
                    return Error1.Ref;
                worksheet = (XLWorksheet)worksheet1;
            }
            else
            {
                worksheet = null;
            }

            if (node.Type == ReferenceItemType.Cell || node.Type == ReferenceItemType.HRange || node.Type == ReferenceItemType.VRange)
                return new Reference(new XLRangeAddress(worksheet, node.Address));

            var rangeName = node.Address;
            worksheet ??= context.Worksheet;
            var found = worksheet.NamedRanges.TryGetValue(rangeName, out var namedRange)
                    || context.Workbook.NamedRanges.TryGetValue(rangeName, out namedRange);
            if (!found)
                return Error1.Name;

            if (!namedRange.IsValid)
                return Error1.Ref;

            if (namedRange.Ranges.Count != 1)
                throw new NotImplementedException("Range to reference conversion not implemented.");

            var range = namedRange.Ranges.Single();
            return new Reference((XLRangeAddress)range.RangeAddress);
        }

        public AnyValue Visit(CalcContext context, NotSupportedNode node)
        {
            throw new NotImplementedException($"Evaluation of {node.FeatureName} is not implemented.");
        }

        public AnyValue Visit(CalcContext context, StructuredReferenceNode node)
        {
            throw new NotImplementedException($"Evaluation of structured references is not implemented.");
        }

        #region Never visited nodes

        public AnyValue Visit(CalcContext context, PrefixNode node) => throw new InvalidOperationException();

        public AnyValue Visit(CalcContext context, FileNode node) => throw new NotImplementedException();

        public AnyValue Visit(CalcContext context, EmptyArgumentNode node) => throw new InvalidOperationException();

        #endregion

    }

    /// <summary>
    /// An adapter for legacy function implementations.
    /// </summary>
    internal class Expression : IComparable<Expression>
    {
        private readonly object _value;

        public Expression(object value)
        {
            _value = value;
        }

        public virtual object Evaluate()
        {
            if (_value is ExpressionErrorType error)
                ThrowApplicableException(error);
            return _value;
        }


        //---------------------------------------------------------------------------

        #region ** implicit converters

        public static implicit operator string(Expression x)
        {
            if (x._value is ExpressionErrorType error)
                ThrowApplicableException(error);

            var v = x.Evaluate();

            if (v == null)
                return string.Empty;

            if (v is bool b)
                return b.ToString().ToUpper();

            return v.ToString();
        }

        public static implicit operator double(Expression x)
        {
            if (x._value is ExpressionErrorType error)
                ThrowApplicableException(error);

            // evaluate
            var v = x.Evaluate();

            // handle doubles
            if (v is double dbl)
            {
                return dbl;
            }

            // handle booleans
            if (v is bool b)
            {
                return b ? 1 : 0;
            }

            // handle dates
            if (v is DateTime dt)
            {
                return dt.ToOADate();
            }

            if (v is TimeSpan ts)
            {
                return ts.TotalDays;
            }

            // handle string
            if (v is string s && double.TryParse(s, out var doubleValue))
            {
                return doubleValue;
            }

            // handle nulls
            if (v == null || v is string)
            {
                return 0;
            }

            // handle everything else
            CultureInfo _ci = Thread.CurrentThread.CurrentCulture;
            return (double)Convert.ChangeType(v, typeof(double), _ci);
        }

        public static implicit operator bool(Expression x)
        {
            if (x._value is ExpressionErrorType error)
                ThrowApplicableException(error);

            // evaluate
            var v = x.Evaluate();

            // handle booleans
            if (v is bool b)
            {
                return b;
            }

            // handle nulls
            if (v == null)
            {
                return false;
            }

            // handle doubles
            if (v is double dbl)
            {
                return dbl != 0;
            }

            // handle everything else
            return (double)Convert.ChangeType(v, typeof(double)) != 0;
        }

        public static implicit operator DateTime(Expression x)
        {
            if (x._value is ExpressionErrorType error)
                ThrowApplicableException(error);

            // evaluate
            var v = x.Evaluate();

            // handle dates
            if (v is DateTime dt)
            {
                return dt;
            }

            if (v is TimeSpan ts)
            {
                return new DateTime().Add(ts);
            }

            // handle numbers
            if (v.IsNumber())
            {
                return DateTime.FromOADate((double)x);
            }

            // handle everything else
            CultureInfo _ci = Thread.CurrentThread.CurrentCulture;
            return (DateTime)Convert.ChangeType(v, typeof(DateTime), _ci);
        }

        #endregion ** implicit converters

        //---------------------------------------------------------------------------

        #region ** IComparable<Expression>

        public int CompareTo(Expression other)
        {
            // get both values
            var c1 = this.Evaluate() as IComparable;
            var c2 = other.Evaluate() as IComparable;

            // handle nulls
            if (c1 == null && c2 == null)
            {
                return 0;
            }
            if (c2 == null)
            {
                return -1;
            }
            if (c1 == null)
            {
                return +1;
            }

            // make sure types are the same
            if (c1.GetType() != c2.GetType())
            {
                try
                {
                    if (c1 is DateTime)
                        c2 = ((DateTime)other);
                    else if (c2 is DateTime)
                        c1 = ((DateTime)this);
                    else
                        c2 = Convert.ChangeType(c2, c1.GetType()) as IComparable;
                }
                catch (InvalidCastException) { return -1; }
                catch (FormatException) { return -1; }
                catch (OverflowException) { return -1; }
                catch (ArgumentNullException) { return -1; }
            }

            // String comparisons should be case insensitive
            if (c1 is string s1 && c2 is string s2)
                return StringComparer.OrdinalIgnoreCase.Compare(s1, s2);
            else
                return c1.CompareTo(c2);
        }

        #endregion ** IComparable<Expression>


        private static void ThrowApplicableException(ExpressionErrorType errorType)
        {
            switch (errorType)
            {
                // TODO: include last token in exception message
                case ExpressionErrorType.CellReference:
                    throw new CellReferenceException();
                case ExpressionErrorType.CellValue:
                    throw new CellValueException();
                case ExpressionErrorType.DivisionByZero:
                    throw new DivisionByZeroException();
                case ExpressionErrorType.NameNotRecognized:
                    throw new NameNotRecognizedException();
                case ExpressionErrorType.NoValueAvailable:
                    throw new NoValueAvailableException();
                case ExpressionErrorType.NullValue:
                    throw new NullValueException();
                case ExpressionErrorType.NumberInvalid:
                    throw new NumberException();
            }
        }
    }


    /// <summary>
    /// Expression that represents an omitted parameter.
    /// </summary>
    internal class EmptyValueExpression : Expression
    {
        public EmptyValueExpression() : base(null)
        {
        }
    }
}