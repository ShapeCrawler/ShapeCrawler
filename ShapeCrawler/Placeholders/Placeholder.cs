using System;
using ShapeCrawler.Shared;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Placeholders;

internal abstract class Placeholder : IPlaceholder
{
    protected Placeholder(P.PlaceholderShape pPlaceholderShape)
    {
        this.PPlaceholderShape = pPlaceholderShape;
    }

    public SCPlaceholderType Type => this.GetPlaceholderType();

    internal P.PlaceholderShape PPlaceholderShape { get; }

    protected internal Shape ReferencedShape => this.ReferencedShapeLazy.Value;

    protected ResettableLazy<Shape> ReferencedShapeLazy { get; set; }

    #region Private Methods

    private SCPlaceholderType GetPlaceholderType()
    {
        var pPlaceholderValue = this.PPlaceholderShape.Type;
        if (pPlaceholderValue == null)
        {
            return SCPlaceholderType.Custom;
        }

        if (pPlaceholderValue == P.PlaceholderValues.Title)
        {
            return SCPlaceholderType.Title;
        }

        if (pPlaceholderValue == P.PlaceholderValues.CenteredTitle)
        {
            return SCPlaceholderType.CenteredTitle;
        }

        return (SCPlaceholderType)Enum.Parse(typeof(SCPlaceholderType), pPlaceholderValue.Value.ToString());
    }

    #endregion Private Methods
}