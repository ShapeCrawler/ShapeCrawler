using System;

namespace ShapeCrawler.Texts;

internal class NullTextFrame : ITextFrame
{
    private const string error = $"The shape is not a text holder. Use {nameof(IShape.IsTextHolder)} method to check it.";
    
    public IParagraphs Paragraphs => throw new Exception(error);

    public string Text
    {
        get => throw new Exception(error); 
        set => throw new Exception(error);
    }

    public AutofitType AutofitType
    {
        get => throw new Exception(error); 
        set => throw new Exception(error);
    }
    
    public decimal LeftMargin 
    { 
        get => throw new Exception(error); 
        set => throw new Exception(error);
    }

    public decimal RightMargin
    {
        get => throw new Exception(error); 
        set => throw new Exception(error);
    }

    public decimal TopMargin
    {
        get => throw new Exception(error); 
        set => throw new Exception(error);
    }

    public decimal BottomMargin
    {
        get => throw new Exception(error); 
        set => throw new Exception(error);
    }
    
    public bool TextWrapped => throw new Exception(error);
    
    public string SDKXPath => throw new Exception(error);
}