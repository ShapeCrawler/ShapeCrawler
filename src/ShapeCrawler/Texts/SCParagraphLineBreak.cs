﻿using System;
using ShapeCrawler.Drawing;
using ShapeCrawler.Exceptions;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Texts;

internal sealed class SCParagraphLineBreak : IParagraphPortion
{
    private readonly A.Break aBreak;
    
    internal SCParagraphLineBreak(A.Break aBreak, Action onRemovedHandler)
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

    public IField? Field { get; }

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