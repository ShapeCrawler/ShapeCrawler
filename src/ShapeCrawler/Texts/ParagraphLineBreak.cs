using System;
using ShapeCrawler.Drawing;
using ShapeCrawler.Exceptions;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Texts;

internal sealed class ParagraphLineBreak : IParagraphPortion
{
    private readonly A.Break aBreak;
    
    internal ParagraphLineBreak(A.Break aBreak, Action onRemovedHandler)
    {
        this.aBreak = aBreak;
        this.Removed += onRemovedHandler;
    }
    
    private event Action Removed;

    public string? Text { get; set; } = Environment.NewLine;

    public ITextPortionFont? Font { get; }

    public string? Hyperlink
    {
        get => null; 
        set => throw new SCException("New Line portion does not support hyperlink.");
    }

    public SCColor? TextHighlightColor
    {
        get => null; 
        set => throw new SCException("New Line portion does not support text highlight color.");
    }

    public void Remove()
    {
        this.aBreak.Remove();
        this.Removed?.Invoke();
    }
}