namespace ShapeCrawler;

internal sealed class SCSlideSize
{
    internal SCSlideSize(int slideWidth, int slideHeight)
    {
        this.Width = slideWidth;
        this.Height = slideHeight;
    }

    internal int Width { get; }

    internal int Height { get; }
}