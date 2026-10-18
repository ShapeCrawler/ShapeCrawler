using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Texts;

internal static class TextReplacement
{
    internal static void Validate(string oldText, string newText)
    {
        ArgumentException.ThrowIfNullOrEmpty(oldText);
        ArgumentNullException.ThrowIfNull(newText);
        XmlConvert.VerifyXmlChars(newText);
    }

    internal static int Replace(OpenXmlElement body, string oldText, string newText, bool ignoreCase)
    {
        Validate(oldText, newText);
        var count = 0;
        foreach (var paragraph in body.Elements<A.Paragraph>())
        {
            var runs = new List<A.Run>();
            foreach (var child in paragraph.ChildElements.ToArray())
            {
                if (child is A.Run { Text: not null } run)
                {
                    runs.Add(run);
                }
                else
                {
                    count += ReplaceRuns(runs, oldText, newText, ignoreCase);
                    runs.Clear();
                }
            }

            count += ReplaceRuns(runs, oldText, newText, ignoreCase);
        }

        return count;
    }

    private static int ReplaceRuns(List<A.Run> runs, string oldText, string newText, bool ignoreCase)
    {
        var texts = runs.Select(run => run.Text!.Text).ToArray();
        var original = string.Concat(texts);
        var matches = new List<int>();
        var comparison = ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
        var offset = 0;
        while (offset <= original.Length - oldText.Length)
        {
            var match = original.IndexOf(oldText, offset, comparison);
            if (match < 0)
            {
                break;
            }

            matches.Add(match);
            offset = match + oldText.Length;
        }

        var changed = new HashSet<A.Run>();
        // Work backwards so earlier matches retain their original offsets.
        foreach (var match in matches.AsEnumerable().Reverse())
        {
            var (first, start) = Locate(texts, match);
            var (last, end) = Locate(texts, match + oldText.Length - 1);
            var prefix = runs[first].Text!.Text[..start];
            var suffix = runs[last].Text!.Text[(end + 1)..];
            runs[first].Text!.Text = prefix + newText + (first == last ? suffix : string.Empty);
            changed.Add(runs[first]);
            if (first != last)
            {
                for (var i = first + 1; i < last; i++)
                {
                    runs[i].Text!.Text = string.Empty;
                }

                runs[last].Text!.Text = suffix;
                changed.Add(runs[last]);
            }
        }

        foreach (var run in changed)
        {
            InsertLineBreaks(run);
        }

        return matches.Count;
    }

    private static (int Run, int Offset) Locate(string[] texts, int offset)
    {
        for (var i = 0; i < texts.Length; i++)
        {
            if (offset < texts[i].Length)
            {
                return (i, offset);
            }

            offset -= texts[i].Length;
        }

        throw new InvalidOperationException("The text match lies outside its runs.");
    }

    private static void InsertLineBreaks(A.Run run)
    {
        var text = run.Text!.Text;
        if (!text.Contains('\r') && !text.Contains('\n'))
        {
            return;
        }

        var lines = text.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n').Split('\n');
        run.Text.Text = lines[0];
        OpenXmlElement previous = run;
        foreach (var line in lines.Skip(1))
        {
            var lineBreak = new A.Break();
            var next = new A.Run();
            if (run.RunProperties is not null)
            {
                lineBreak.RunProperties = (A.RunProperties)run.RunProperties.CloneNode(true);
                next.RunProperties = (A.RunProperties)run.RunProperties.CloneNode(true);
            }

            next.Text = new A.Text(line);
            previous.InsertAfterSelf(lineBreak);
            lineBreak.InsertAfterSelf(next);
            previous = next;
        }
    }
}
