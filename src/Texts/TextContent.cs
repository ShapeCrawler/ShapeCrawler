using System;
using System.Linq;
using ShapeCrawler.Paragraphs;

namespace ShapeCrawler.Texts;

/// <summary>
///     Represents a plain text content.
/// </summary>
internal sealed class TextContent(
    string text,
    IParagraphCollection paragraphs,
    Func<AutofitType> getAutofitType,
    Action<string> shrinkFont,
    Action applyResize)
{
    /// <summary>
    ///     Applies the text content to the paragraphs.
    /// </summary>
    internal void ApplyTo()
    {
        var paragraphsList = paragraphs.ToArray();
        var firstParagraph = paragraphsList.FirstOrDefault();

        // Store LatinName from first portion if available
        string? latinNameToPreserve = GetLatinNameToPreserve(firstParagraph);

        // Store font color hex from first portion if available
        string? colorHexToPreserve = GetFontColorHexToPreserve(firstParagraph);

        // Clear existing content and ensure we have a first paragraph
        firstParagraph = this.PrepareContainer(firstParagraph, paragraphsList);

        // Add new text with preserved font
        var paragraphLines = text.Split([Environment.NewLine], StringSplitOptions.None);
        this.AddToParagraphs(paragraphLines, firstParagraph, latinNameToPreserve);
        if (colorHexToPreserve != null)
        {
            for (int i = 0; i < paragraphLines.Length; i++)
            {
                var portion = paragraphs[i].Portions.Last();
                portion.Font!.Color.Set(colorHexToPreserve);
            }
        }

        this.ApplyFormatting();
    }

    private static string? GetLatinNameToPreserve(IParagraph? firstParagraph)
    {
        var firstPortion = firstParagraph?.Portions.FirstOrDefault();
        return firstPortion?.Font!.LatinName;
    }

    private static string? GetFontColorHexToPreserve(IParagraph? firstParagraph)
    {
        var firstPortion = firstParagraph?.Portions.FirstOrDefault();
        return firstPortion?.Font?.Color.Hex;
    }

    private static void ApplyLatinNameIfNeeded(IParagraphPortion portion, string? latinNameToPreserve)
    {
        if (latinNameToPreserve != null && portion.Font != null)
        {
            portion.Font.LatinName = latinNameToPreserve;
        }
    }

    private IParagraph PrepareContainer(IParagraph? firstParagraph, IParagraph[] paragraphsList)
    {
        if (firstParagraph == null)
        {
            paragraphs.Add();
            return paragraphs.First();
        }

        foreach (var paragraph in paragraphsList.Skip(1))
        {
            paragraph.Remove();
        }

        // Keep the first text run as a formatting template, including inherited font size.
        var firstTextPortion = firstParagraph.Portions.FirstOrDefault() as TextParagraphPortion;
        foreach (var portion in firstParagraph.Portions.Skip(firstTextPortion == null ? 0 : 1).ToList())
        {
            portion.Remove();
        }

        return firstParagraph;
    }

    private void AddToParagraphs(string[] paragraphLines, IParagraph firstParagraph, string? latinNameToPreserve)
    {
        if (paragraphLines.Length <= 0)
        {
            return;
        }

        // Add first line to the first paragraph
        var templatePortion = firstParagraph.Portions.FirstOrDefault();
        firstParagraph.Portions.AddText(paragraphLines[0]);
        templatePortion?.Remove();
        ApplyLatinNameIfNeeded(firstParagraph.Portions.Last(), latinNameToPreserve);

        // Add remaining lines as new paragraphs
        for (int i = 1; i < paragraphLines.Length; i++)
        {
            paragraphs.Add();
            var emptyPortions = paragraphs[i].Portions.ToArray();
            paragraphs[i].Portions.AddText(paragraphLines[i]);
            foreach (var emptyPortion in emptyPortions)
            {
                emptyPortion.Remove();
            }

            ApplyLatinNameIfNeeded(paragraphs[i].Portions.Last(), latinNameToPreserve);
        }
    }

    private void ApplyFormatting()
    {
        if (getAutofitType() == AutofitType.Shrink)
        {
            shrinkFont(text);
        }

        applyResize();
    }
}
