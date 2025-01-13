using DocumentFormat.OpenXml;
using ShapeCrawler.Units;

namespace ShapeCrawler.Tables;

internal class TopBorder : IBorder
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
        
        var emus = new Points((decimal)points).AsEmus();
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