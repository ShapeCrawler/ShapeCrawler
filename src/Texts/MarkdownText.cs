using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace ShapeCrawler.Texts;

/// <summary>
///     Represents markdown-formatted text that can be applied to paragraphs.
/// </summary>
internal sealed class MarkdownText
{
    private readonly string markdownText;
    private readonly IParagraphCollection paragraphs;
    private readonly Func<AutofitType> getAutofitType;
    private readonly Action<string> shrinkFont;
    private readonly Action applyResize;

    internal MarkdownText(
        string markdownText,
        IParagraphCollection paragraphs,
        Func<AutofitType> getAutofitType,
        Action<string> shrinkFont,
        Action applyResize)
    {
        this.markdownText = markdownText;
        this.paragraphs = paragraphs;
        this.getAutofitType = getAutofitType;
        this.shrinkFont = shrinkFont;
        this.applyResize = applyResize;
    }

    /// <summary>
    ///     Applies markdown-formatted text to the paragraphs.
    /// </summary>
    internal void ApplyTo()
    {
        var lines = Regex.Split(this.markdownText, "\r\n|\r|\n", RegexOptions.None, TimeSpan.FromMilliseconds(1000));
        if (IsList(lines))
        {
            this.RenderList(lines);
        }
        else
        {
            this.RenderRegular(this.markdownText);
        }

        this.applyResize();
    }

    private static bool IsList(string[] lines) =>
        lines.Any(l => l.TrimStart().StartsWith("- ", StringComparison.CurrentCulture));

    private void RenderList(string[] lines)
    {
        var paragraphsList = this.paragraphs.ToList();
        var firstPara = paragraphsList.FirstOrDefault();
        if (firstPara == null)
        {
            return;
        }

        foreach (var p in paragraphsList.Skip(1))
        {
            p.Remove();
        }

        foreach (var portion in firstPara.Portions.ToList())
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
                this.paragraphs.Add();
            }

            var paragraph = this.paragraphs[paraIndex];
            foreach (var portion in paragraph.Portions.ToList())
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
        var paragraphsList = this.paragraphs.ToList();
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

        if (this.getAutofitType() == AutofitType.Shrink)
        {
            this.shrinkFont(text);
        }
    }
}