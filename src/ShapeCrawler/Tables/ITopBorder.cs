using DocumentFormat.OpenXml;
using ShapeCrawler.Units;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a top border of a table cell.
/// </summary>
public interface ITopBorder
{
    /// <summary>
    ///     Gets or sets border width in points.
    /// </summary>
    float Width { get; set; }
}

internal class TopBorder : ITopBorder
{
    private readonly DocumentFormat.OpenXml.Drawing.TableCellProperties aTableCellProperties;

    internal TopBorder(DocumentFormat.OpenXml.Drawing.TableCellProperties aTableCellProperties)
    {
        this.aTableCellProperties = aTableCellProperties;
    }

    public float Width
    {
        get => this.GetWidth();
        set => this.UpdateWidth(value);
    }

    private void UpdateWidth(float points)
    {
        if (this.aTableCellProperties.TopBorderLineProperties is null)
        {
            var aSolidFill = new DocumentFormat.OpenXml.Drawing.SolidFill
            {
                SchemeColor = new DocumentFormat.OpenXml.Drawing.SchemeColor { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.Text1 }
            };
            this.aTableCellProperties.TopBorderLineProperties = new DocumentFormat.OpenXml.Drawing.TopBorderLineProperties();
            this.aTableCellProperties.TopBorderLineProperties.AppendChild(aSolidFill);
        }
        
        var emus = new Points(points).AsEmus();
        this.aTableCellProperties.TopBorderLineProperties.Width = new Int32Value((int)emus);
    }

    private float GetWidth()
    {
        if (this.aTableCellProperties.TopBorderLineProperties is null)
        {
            return 1; // default value
        }

        var emus = this.aTableCellProperties.TopBorderLineProperties!.Width!.Value;
        
        return new Emus(emus).AsPoints();
    }
}