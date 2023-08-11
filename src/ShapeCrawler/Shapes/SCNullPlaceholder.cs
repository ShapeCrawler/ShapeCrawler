using ShapeCrawler.Exceptions;

namespace ShapeCrawler.Shapes;

internal class SCNullPlaceholder : IPlaceholder
{
    private readonly string error;

    internal SCNullPlaceholder(string error)
    {
        this.error = error;
    }

    public SCPlaceholderType Type => throw new SCException(this.error);
}