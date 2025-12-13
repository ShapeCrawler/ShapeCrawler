using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace ShapeCrawler.Texts;

/// <summary>
///     Represents Markdown text.
/// </summary>
internal sealed class MarkdownText(
    string markdownText,
    IParagraphCollection paragraphs,
    Func<AutofitType> getAutofitType,
    Action<string> shrinkFont,
    Action applyResize)
{
    /// <summary>
    ///     Applies markdown-formatted text to the paragraphs.
    /// </summary>
    internal void ApplyTo()
    {
        var lines = Regex.Split(markdownText, "\r\n|\r|\n", RegexOptions.None, TimeSpan.FromMilliseconds(1000));
        if (IsList(lines))
        {
            this.RenderList(lines);
        }
        else
        {
            this.RenderRegular(markdownText);
        }

        applyResize();
    }

    private static bool IsList(string[] lines) =>
        lines.Any(l => l.TrimStart().StartsWith("- ", StringComparison.CurrentCulture));

    private void RenderList(string[] lines)
    {
        var paragraphsList = paragraphs.ToArray();
        var firstPara = paragraphsList.FirstOrDefault();
        if (firstPara == null)
        {
            return;
        }

        foreach (var p in paragraphsList.Skip(1))
        {
            p.Remove();
        }

        foreach (var portion in firstPara.Portions.ToArray())
        {
            portion.Remove();
        }

        int paraIndex = 0;
        foreach (var rawLine in lines)
        {
            if (string.IsNullOrWhiteSpace(rawLine))
            {
                continue;
            }

            var line = rawLine.TrimStart();
            if (!line.StartsWith("- ", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var content = line[2..];
            if (paraIndex > 0)
            {
                paragraphs.Add();
            }

            var paragraph = paragraphs[paraIndex];
            foreach (var portion in paragraph.Portions.ToArray())
            {
                portion.Remove();
            }

            paragraph.Portions.AddText(content);
            paragraph.Bullet.Type = BulletType.Character;
            paragraph.Bullet.Character = "•";
            paraIndex++;
        }
    }

    private void RenderRegular(string text)
    {
        var paragraphsList = paragraphs.ToArray();
        var portionPara = paragraphsList.FirstOrDefault(p => p.Portions.Any()) ?? paragraphsList.First();

        // Clear other paragraphs
        foreach (var p in paragraphsList.Where(p => p != portionPara))
        {
            p.Remove();
        }

        foreach (var portion in portionPara.Portions.ToList())
        {
            portion.Remove();
        }

        const string markdownPattern = @"(\*\*(?<bold>[^\*]+)\*\*)|(?<regular>[^\*]+)";
        var matches = Regex.Matches(
            text,
            markdownPattern,
            RegexOptions.Singleline | RegexOptions.IgnoreCase,
            TimeSpan.FromMilliseconds(1000));
        foreach (Match match in matches)
        {
            if (match.Groups["bold"].Success)
            {
                portionPara.Portions.AddText(match.Groups["bold"].Value);
                portionPara.Portions.Last().Font!.IsBold = true;
            }
            else if (match.Groups["regular"].Success)
            {
                portionPara.Portions.AddText(match.Groups["regular"].Value);
                portionPara.Portions.Last().Font!.IsBold = false;
            }
        }

        if (getAutofitType() == AutofitType.Shrink)
        {
            shrinkFont(text);
        }
    }
}