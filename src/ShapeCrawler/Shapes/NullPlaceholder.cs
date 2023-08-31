using ShapeCrawler.Exceptions;

namespace ShapeCrawler.Shapes;

internal class NullPlaceholder : IPlaceholder
{
    private readonly string error;

    internal NullPlaceholder()
        : this($"The shape is not a placeholder. Use {nameof(IShape.IsPlaceholder)} property to check if shape is a placeholder.")
    {
    }

    internal NullPlaceholder(string error)
    {
        this.error = error;
    }

    public SCPlaceholderType Type => throw new SCException(this.error);
}