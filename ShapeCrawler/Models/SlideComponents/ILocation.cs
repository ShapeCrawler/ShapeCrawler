namespace ShapeCrawler.Models.SlideComponents
{
    /// <summary>
    /// Represents a shape location and size data.
    /// </summary>
    public interface ILocation
    {
        long X { get; }
      
        long Y { get; }
      
        long Width { get; }
      
        long Height { get; }
       
        void SetX(long x);
       
        void SetY(long y);
       
        void SetWidth(long w);
       
        void SetHeight(long h);
    }
}