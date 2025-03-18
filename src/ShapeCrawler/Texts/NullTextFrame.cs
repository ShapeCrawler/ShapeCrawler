using System;

namespace ShapeCrawler.Texts;

internal readonly struct NullTextFrame : ITextBox
{
    private const string Error = $"The shape is not a text holder. Use {nameof(IShape.IsTextHolder)} method to check it.";
    
    public IParagraphs Paragraphs => throw new Exception(Error);

    public string Text
    {
        get => throw new Exception(Error); 
        set => throw new Exception(Error);
    }

    public AutofitType AutofitType
    {
        get => throw new Exception(Error); 
        set => throw new Exception(Error);
    }
    
    public decimal LeftMargin 
    { 
        get => throw new Exception(Error); 
        set => throw new Exception(Error);
    }

    public decimal RightMargin
    {
        get => throw new Exception(Error); 
        set => throw new Exception(Error);
    }

    public decimal TopMargin
    {
        get => throw new Exception(Error); 
        set => throw new Exception(Error);
    }

    public decimal BottomMargin
    {
        get => throw new Exception(Error); 
        set => throw new Exception(Error);
    }

    public bool TextWrapped => throw new Exception(Error);
    
    public string SdkXPath => throw new Exception(Error);

    public TextVerticalAlignment VerticalAlignment 
    {
        get => throw new Exception(Error);
        set => throw new Exception(Error);
    }
}