namespace ShapeCrawler
{
    public interface IRemovable // TODO: make internal
    {
        bool IsRemoved { get; set; }

        void ThrowIfRemoved();
    }
}