namespace ShapeCrawler.Drawing
{
    internal interface IImageContainer : IRemovable
    {
        SCPresentation ParentPresentation { get; }
    }
}
