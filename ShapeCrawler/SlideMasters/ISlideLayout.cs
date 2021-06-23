namespace ShapeCrawler.SlideMasters
{
    /// <summary>
    ///     Represents a Slide Layout.
    /// </summary>
    public interface ISlideLayout : IBaseSlide
    {
        /// <summary>
        ///     Gets parent Slide Master.
        /// </summary>
        ISlideMaster ParentSlideMaster { get; }
    }
}