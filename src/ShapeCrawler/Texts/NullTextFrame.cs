﻿using System;

namespace ShapeCrawler.Shapes;

internal class NullTextFrame : ITextFrame
{
    private const string error = $"The shape is not a text holder. Use {nameof(IShape.IsTextHolder)} method to check it.";
    public IParagraphCollection Paragraphs => throw new Exception(error);

    public string Text
    {
        get => throw new Exception(error); 
        set => throw new Exception(error);
    }

    public AutofitType AutofitType
    {
        get => throw new Exception(error); 
        set=> throw new Exception(error);
    }
    public double LeftMargin { 
        get => throw new Exception(error); 
        set => throw new Exception(error);
    }

    public double RightMargin
    {
        get => throw new Exception(error); 
        set => throw new Exception(error);
    }

    public double TopMargin
    {
        get => throw new Exception(error); 
        set => throw new Exception(error);
    }

    public double BottomMargin
    {
        get => throw new Exception(error); 
        set => throw new Exception(error);
    }
    public bool TextWrapped => throw new Exception(error);
}