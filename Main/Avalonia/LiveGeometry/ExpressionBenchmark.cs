using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Avalonia.Controls;
using DynamicGeometry;
using Drawing = DynamicGeometry.Drawing;

namespace LiveGeometry;

/// <summary>
/// What each <see cref="ExpressionStrategy"/> costs: every drawing of the gallery is loaded
/// under each (with the time its expressions took to compile, out of the load), then its
/// expressions are evaluated over and over, its function graphs sampled, and the whole
/// drawing recalculated as a drag does; the values are compared between the strategies.
/// "LiveGeometry.Desktop.exe --bench-expressions <file>" writes the table to the file, the
/// browser build at "/?bench=expressions" to the console (webauto console).
/// </summary>
public static class ExpressionBenchmark
{
    // each expression is evaluated this many times, each drawing recalculated this many
    // times, a function graph's function sampled at this many x
    const int EvaluationRounds = 200;
    const int RecalculationRounds = 20;
    const int FunctionSamples = 1000;

    class Result
    {
        public ExpressionStrategy Strategy;
        public TimeSpan Load;
        public TimeSpan Compile;
        public TimeSpan Evaluation;
        public TimeSpan Functions;
        public TimeSpan Recalculation;
        public int ExpressionCount;
        public int FunctionCount;
        public int Mismatches;
        public double LargestDifference;
    }

    public static void Run(Action<string> log)
    {
        var strategies = (ExpressionStrategy[])Enum.GetValues(typeof(ExpressionStrategy));
        var previousStrategy = Compiler.Instance.Strategy;
        bool previousSuppress = PointBase.SuppressAutoLabelPoints;
        var results = new List<Result>();

        // the expressions' values under the first strategy, by drawing: the others must agree
        Dictionary<string, double[]> reference = null;
        try
        {
            PointBase.SuppressAutoLabelPoints = true;

            // a pass that counts for nothing: the first strategy measured paid for the JIT
            // of everything a load runs (the browser's runtime warms up too)
            Compiler.Instance.Strategy = strategies[0];
            foreach (var item in GalleryCatalog.Items)
            {
                Measure(item, new Result(), new Dictionary<string, double[]>());
            }

            foreach (var strategy in strategies)
            {
                Compiler.Instance.Strategy = strategy;
                var result = new Result { Strategy = strategy };
                var values = new Dictionary<string, double[]>();
                foreach (var item in GalleryCatalog.Items)
                {
                    Measure(item, result, values);
                }

                if (reference == null)
                {
                    reference = values;
                }
                else
                {
                    Compare(reference, values, result);
                }

                results.Add(result);
            }
        }
        finally
        {
            Compiler.Instance.Strategy = previousStrategy;
            PointBase.SuppressAutoLabelPoints = previousSuppress;
        }

        var first = results[0];
        log($"Expression strategies over {GalleryCatalog.Items.Count} gallery drawings: {first.ExpressionCount} expressions, {first.FunctionCount} functions; "
            + $"times in ms, evaluation is {EvaluationRounds} rounds of every expression, functions {FunctionSamples} samples of each, recalculation {RecalculationRounds} of every drawing.");
        log(Row("strategy", "load", "compile", "share", "evaluate", "functions", "recalculate", "mismatches"));
        foreach (var result in results)
        {
            log(Row(
                result.Strategy.ToString(),
                Milliseconds(result.Load),
                Milliseconds(result.Compile),
                (100 * result.Compile.TotalMilliseconds / System.Math.Max(1, result.Load.TotalMilliseconds)).ToString("0", CultureInfo.InvariantCulture) + "%",
                Milliseconds(result.Evaluation),
                Milliseconds(result.Functions),
                Milliseconds(result.Recalculation),
                result.Strategy == first.Strategy ? "-" : result.Mismatches + (result.Mismatches > 0 ? " (largest " + result.LargestDifference.ToString("0.###e0", CultureInfo.InvariantCulture) + ")" : "")));
        }
    }

    static void Measure(GalleryItem item, Result result, Dictionary<string, double[]> values)
    {
        var element = XElement.Parse(item.LoadText());

        // a surface nothing shows, as a gallery tile has
        var canvas = new Canvas { Width = 560, Height = 380 };
        Compiler.Instance.CompileTime = TimeSpan.Zero;
        var stopwatch = Stopwatch.StartNew();
        var drawing = new Drawing(canvas);
        drawing.AddFromXml(element);
        result.Load += stopwatch.Elapsed;
        result.Compile += Compiler.Instance.CompileTime;

        // the delegates a recalculation runs: the expressions of the points by coordinates
        // and the figures by equation, the functions of the graphs
        var expressions = drawing.Figures
            .OfType<IExpressionOwner>()
            .SelectMany(owner => owner.Expressions)
            .Where(expression => expression != null && expression.IsValid)
            .Select(expression => expression.Value)
            .ToArray();
        var functions = drawing.Figures
            .OfType<FunctionGraph>()
            .Select(graph => graph.Function)
            .Where(function => function != null)
            .ToArray();
        result.ExpressionCount += expressions.Length;
        result.FunctionCount += functions.Length;

        var sample = new double[expressions.Length];
        stopwatch.Restart();
        for (int round = 0; round < EvaluationRounds; round++)
        {
            for (int i = 0; i < expressions.Length; i++)
            {
                sample[i] = expressions[i]();
            }
        }

        result.Evaluation += stopwatch.Elapsed;
        values[item.Slug] = sample;

        stopwatch.Restart();
        foreach (var function in functions)
        {
            for (int i = 0; i < FunctionSamples; i++)
            {
                function(-5 + 10.0 * i / FunctionSamples);
            }
        }

        result.Functions += stopwatch.Elapsed;

        stopwatch.Restart();
        for (int round = 0; round < RecalculationRounds; round++)
        {
            drawing.Recalculate();
        }

        result.Recalculation += stopwatch.Elapsed;
    }

    static void Compare(Dictionary<string, double[]> reference, Dictionary<string, double[]> values, Result result)
    {
        foreach (var pair in values)
        {
            if (!reference.TryGetValue(pair.Key, out var expected) || expected.Length != pair.Value.Length)
            {
                result.Mismatches++;
                continue;
            }

            for (int i = 0; i < expected.Length; i++)
            {
                double actual = pair.Value[i];
                if (double.IsNaN(expected[i]) && double.IsNaN(actual))
                {
                    continue;
                }

                double difference = System.Math.Abs(actual - expected[i]) / System.Math.Max(1, System.Math.Abs(expected[i]));
                if (!(difference <= 1e-12))
                {
                    result.Mismatches++;
                    result.LargestDifference = System.Math.Max(result.LargestDifference, double.IsNaN(difference) ? double.PositiveInfinity : difference);
                }
            }
        }
    }

    static string Milliseconds(TimeSpan time)
    {
        return time.TotalMilliseconds.ToString("0", CultureInfo.InvariantCulture);
    }

    static string Row(params string[] cells)
    {
        var row = new StringBuilder();
        int[] widths = { 16, 8, 8, 6, 9, 10, 12, 12 };
        for (int i = 0; i < cells.Length; i++)
        {
            row.Append(i == 0 ? cells[i].PadRight(widths[i]) : cells[i].PadLeft(widths[i]));
        }

        return row.ToString();
    }
}
