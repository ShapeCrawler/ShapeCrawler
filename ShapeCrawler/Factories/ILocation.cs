namespace ShapeCrawler.Factories
{
    internal interface ILocation
    {
        int X { get; }

        int Y { get; }

        int Width { get; }

        int Height { get; }

        void SetX(int x);

        void SetY(int y);

        void SetWidth(int w);

        void SetHeight(int h);
    }
}