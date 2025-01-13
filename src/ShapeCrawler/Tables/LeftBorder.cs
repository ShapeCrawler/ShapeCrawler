using DocumentFormat.OpenXml;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Tables;

internal class LeftBorder : IBorder
{
    private readonly A.TableCellProperties aTableCellProperties;

    internal LeftBorder(A.TableCellProperties aTableCellProperties)
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
        if (this.aTableCellProperties.LeftBorderLineProperties is null)
        {
            var aSolidFill = new A.SolidFill
            {
                SchemeColor = new A.SchemeColor { Val = A.SchemeColorValues.Text1 }
            };
            this.aTableCellProperties.LeftBorderLineProperties = new A.LeftBorderLineProperties();
            this.aTableCellProperties.LeftBorderLineProperties.AppendChild(aSolidFill);
        }
        
        var emus = new Points((decimal)points).AsEmus();
        this.aTableCellProperties.LeftBorderLineProperties!.Width = new Int32Value((int)emus);
    }

    private float GetWidth()
    {
        if (this.aTableCellProperties.LeftBorderLineProperties is null)
        {
            return 1; // default value
        }

        var emus = this.aTableCellProperties.LeftBorderLineProperties!.Width!.Value;
        
        return new Emus(emus).AsPoints();
    }
}