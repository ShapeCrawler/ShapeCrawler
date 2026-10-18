using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using ShapeCrawler.Extensions;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Texts;

internal static class TextHighlighting
{
    internal static void Validate(string text, Color color)
    {
        ArgumentException.ThrowIfNullOrEmpty(text);
        ValidateColor(color);
    }

    internal static void ValidateColor(Color color)
    {
        if (!float.IsFinite(color.Alpha) || color.Alpha is < 0 or > 255)
        {
            throw new ArgumentOutOfRangeException(nameof(color), "Highlight alpha must be between 0 and 255.");
        }
    }

    internal static int HighlightText(OpenXmlElement body, string text, Color color, bool ignoreCase)
    {
        Validate(text, color);
        return Highlight(body, color, original =>
        {
            var ranges = new List<(int Start, int End)>();
            var comparison = ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            var offset = 0;
            while (offset <= original.Length - text.Length)
            {
                var index = original.IndexOf(text, offset, comparison);
                if (index < 0)
                {
                    break;
                }

                offset = index + text.Length;
                ranges.Add((index, offset));
            }

            return ranges;
        });
    }

    internal static int HighlightRegex(OpenXmlElement body, Regex regex, Color color)
    {
        ArgumentNullException.ThrowIfNull(regex);
        ValidateColor(color);
        return Highlight(body, color, original => [.. regex.Matches(original).Cast<Match>()
            .Where(match => match.Length > 0).Select(match => (match.Index, match.Index + match.Length))]);
    }

    private static int Highlight(OpenXmlElement body, Color color, Func<string, List<(int Start, int End)>> find)
    {
        var plans = new List<HighlightPlan>();
        var count = 0;
        foreach (var paragraph in body.Elements<A.Paragraph>())
        {
            var runs = new List<A.Run>();
            foreach (var child in paragraph.ChildElements)
            {
                if (child is A.Run { Text: not null } run)
                {
                    runs.Add(run);
                }
                else
                {
                    count += Prepare(runs, find, plans);
                    runs.Clear();
                }
            }

            count += Prepare(runs, find, plans);
        }

        // Finish matching before changing this text box, including regex timeouts and Unicode checks.
        foreach (var plan in plans)
        {
            Apply(plan, color);
        }

        return count;
    }

    private static int Prepare(List<A.Run> runs, Func<string, List<(int Start, int End)>> find, List<HighlightPlan> plans)
    {
        if (runs.Count == 0)
        {
            return 0;
        }

        var original = string.Concat(runs.Select(run => run.Text!.Text));
        var matches = find(original);
        var ranges = new List<(int Start, int End)>();
        foreach (var match in matches.OrderBy(match => match.Start))
        {
            if (ranges.Count > 0 && match.Start <= ranges[^1].End)
            {
                ranges[^1] = (ranges[^1].Start, Math.Max(ranges[^1].End, match.End));
            }
            else
            {
                ranges.Add(match);
            }
        }

        foreach (var (start, end) in ranges)
        {
            foreach (var boundary in new[] { start, end })
            {
                if (boundary > 0 && boundary < original.Length && char.IsHighSurrogate(original[boundary - 1])
                    && char.IsLowSurrogate(original[boundary]))
                {
                    throw new ArgumentException("A highlight match cannot split a Unicode surrogate pair.");
                }
            }
        }

        if (ranges.Count > 0)
        {
            plans.Add(new HighlightPlan([.. runs], ranges));
        }

        return matches.Count;
    }

    private static void Apply(HighlightPlan plan, Color color)
    {
        var absolute = 0;
        var matchIndex = 0;
        foreach (var run in plan.Runs)
        {
            var text = run.Text!.Text;
            var start = absolute;
            absolute += text.Length;
            if (text.Length == 0)
            {
                continue;
            }

            while (matchIndex < plan.Ranges.Count && plan.Ranges[matchIndex].End <= start)
            {
                matchIndex++;
            }

            if (matchIndex == plan.Ranges.Count || plan.Ranges[matchIndex].Start >= absolute)
            {
                continue;
            }

            var pieces = new List<(string Text, bool Selected)>();
            var cursor = 0;
            for (var index = matchIndex; index < plan.Ranges.Count && plan.Ranges[index].Start < absolute; index++)
            {
                var from = Math.Max(start, plan.Ranges[index].Start) - start;
                var to = Math.Min(absolute, plan.Ranges[index].End) - start;
                if (cursor < from)
                {
                    pieces.Add((text[cursor..from], false));
                }

                pieces.Add((text[from..to], true));
                cursor = to;
            }

            if (cursor < text.Length)
            {
                pieces.Add((text[cursor..], false));
            }

            var template = (A.Run)run.CloneNode(true);
            OpenXmlElement previous = run;
            for (var index = 0; index < pieces.Count; index++)
            {
                var piece = index == 0 ? run : (A.Run)template.CloneNode(true);
                piece.Text!.Text = pieces[index].Text;
                if (pieces[index].Selected)
                {
                    piece.RunProperties ??= new A.RunProperties();
                    piece.RunProperties.AddAHighlight(color);
                }

                if (index > 0)
                {
                    previous.InsertAfterSelf(piece);
                }

                previous = piece;
            }
        }
    }

    private sealed record HighlightPlan(A.Run[] Runs, List<(int Start, int End)> Ranges);
}
