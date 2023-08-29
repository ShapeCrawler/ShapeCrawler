using ShapeCrawler.Exceptions;

namespace ShapeCrawler.AutoShapes;

internal sealed record NullTextFrame : ITextFrame
{
    private const string Error = $"The AutoShape is not a text holder. Use {nameof(IShape.IsTextHolder)} property to check if a AutoShape is a text holder.";

    public IParagraphCollection Paragraphs => throw new SCException(Error);

    public string Text
    {
        get => throw new SCException(Error);
        set => throw new SCException(Error);
    }
    public SCAutofitType AutofitType { get; set; }
    public double LeftMargin { get; set; }
    public double RightMargin { get; set; }
    public double TopMargin { get; set; }
    public double BottomMargin { get; set; }
    public bool TextWrapped => throw new SCException(Error);
}