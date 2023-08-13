using System.IO;
using ShapeCrawler.Exceptions;

namespace ShapeCrawler.Shapes;

internal record SCNullShapeFill : IShapeFill
{
    private const string errorText = "This Auto Shape type is not capable of having fill.";
    public SCFillType Type => throw new SCException(errorText);
    public IImage Picture => throw new SCException(errorText);
    public string Color => throw new SCException(errorText);
    public double AlphaPercentage => throw new SCException(errorText);
    public double LuminanceModulationPercentage => throw new SCException(errorText);
    public double LuminanceOffsetPercentage => throw new SCException(errorText);

    public void SetPicture(Stream image)
    {
        throw new SCException(errorText);
    }

    public void SetColor(string hex)
    {
        throw new SCException(errorText);
    }
}