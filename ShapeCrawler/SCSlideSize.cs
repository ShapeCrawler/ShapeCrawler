namespace ShapeCrawler
{
    internal class SCSlideSize
    {
        public SCSlideSize(int slideWidth, int slideHeight)
        {
            this.Width = slideWidth;
            this.Height = slideHeight;
        }

        public int Width { get; }

        public int Height { get; }
    }
}