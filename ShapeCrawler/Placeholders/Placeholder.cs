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

    internal abstract ResettableLazy<Shape?> ReferencedShape { get; }

    #region Private Methods

    private SCPlaceholderType GetPlaceholderType()
    {
        var pPlaceholderValue = this.PPlaceholderShape.Type;
        if (pPlaceholderValue == null)
        {
            return SCPlaceholderType.Content;
        }

        if (pPlaceholderValue == P.PlaceholderValues.Title)
        {
            return SCPlaceholderType.Title;
        }

        if (pPlaceholderValue == P.PlaceholderValues.CenteredTitle)
        {
            return SCPlaceholderType.CenteredTitle;
        }

        if (pPlaceholderValue == P.PlaceholderValues.Body)
        {
            return SCPlaceholderType.Text;
        }

        if (pPlaceholderValue == P.PlaceholderValues.Diagram)
        {
            return SCPlaceholderType.SmartArt;
        }

        if (pPlaceholderValue == P.PlaceholderValues.ClipArt)
        {
            return SCPlaceholderType.OnlineImage;
        }
        
        return (SCPlaceholderType)Enum.Parse(typeof(SCPlaceholderType), pPlaceholderValue.Value.ToString());
    }

    #endregion Private Methods
}