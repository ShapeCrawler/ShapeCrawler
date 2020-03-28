namespace SlideDotNet.Models.SlideComponents
{
    public interface IInnerTransform
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