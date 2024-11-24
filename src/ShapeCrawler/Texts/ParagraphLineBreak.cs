using System;
using ShapeCrawler.Exceptions;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Texts;

internal sealed record ParagraphLineBreak : IParagraphPortion
{
    private readonly A.Break aBreak;
    
    internal ParagraphLineBreak(A.Break aBreak)
    {
        this.aBreak = aBreak;
    }

    public string? Text { get; set; } = Environment.NewLine;

    public ITextPortionFont? Font { get; }

    public IHyperlink? Link => null;

    public Color TextHighlightColor
    {
        get => throw new SCException("New Line portion does not support text highlight color.");
        set => throw new SCException("New Line portion does not support text highlight color.");
    }

    public void Remove()
    {
        this.aBreak.Remove();
    }
}