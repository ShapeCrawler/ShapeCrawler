using System;
using System.Linq;

namespace ShapeCrawler.Texts;

/// <summary>
///     Represents the plain text content of a textbox with formatting preservation.
/// </summary>
internal sealed class TextContent
{
    private readonly string text;
    private readonly IParagraphCollection paragraphs;
    private readonly Func<AutofitType> getAutofitType;
    private readonly Action<string> shrinkFont;
    private readonly Action applyResize;

    internal TextContent(
        string text,
        IParagraphCollection paragraphs,
        Func<AutofitType> getAutofitType,
        Action<string> shrinkFont,
        Action applyResize)
    {
        this.text = text;
        this.paragraphs = paragraphs;
        this.getAutofitType = getAutofitType;
        this.shrinkFont = shrinkFont;
        this.applyResize = applyResize;
    }

    /// <summary>
    ///     Applies the text content to the paragraphs.
    /// </summary>
    internal void ApplyTo()
    {
        var paragraphsList = this.paragraphs.ToList();
        var firstParagraph = paragraphsList.FirstOrDefault();

        // Store LatinName from first portion if available
        string? latinNameToPreserve = GetLatinNameToPreserve(firstParagraph);

        // Store font color hex from first portion if available
        string? colorHexToPreserve = GetFontColorHexToPreserve(firstParagraph);

        // Clear existing content and ensure we have a first paragraph
        firstParagraph = this.PrepareContainer(firstParagraph, paragraphsList);

        // Add new text with preserved font
        var paragraphLines = this.text.Split([Environment.NewLine], StringSplitOptions.None);
        this.AddToParagraphs(paragraphLines, firstParagraph, latinNameToPreserve);
        if (colorHexToPreserve != null)
        {
            for (int i = 0; i < paragraphLines.Length; i++)
            {
                var portion = this.paragraphs[i].Portions.Last();
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

    private IParagraph PrepareContainer(IParagraph? firstParagraph, System.Collections.Generic.List<IParagraph> paragraphsList)
    {
        if (firstParagraph == null)
        {
            this.paragraphs.Add();
            return this.paragraphs.First();
        }
        
        foreach (var paragraph in paragraphsList.Skip(1))
        {
            paragraph.Remove();
        }

        foreach (var portion in firstParagraph.Portions.ToList())
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
        firstParagraph.Portions.AddText(paragraphLines[0]);
        ApplyLatinNameIfNeeded(firstParagraph.Portions.Last(), latinNameToPreserve);

        // Add remaining lines as new paragraphs
        for (int i = 1; i < paragraphLines.Length; i++)
        {
            this.paragraphs.Add();
            this.paragraphs[i].Portions.AddText(paragraphLines[i]);
            ApplyLatinNameIfNeeded(this.paragraphs[i].Portions.Last(), latinNameToPreserve);
        }
    }

    private void ApplyFormatting()
    {
        if (this.getAutofitType() == AutofitType.Shrink)
        {
            this.shrinkFont(this.text);
        }

        this.applyResize();
    }
}