using DocumentFormat.OpenXml;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Tables;

internal class RightBorder : IBorder
{
    private readonly A.TableCellProperties aTableCellProperties;

    internal RightBorder(A.TableCellProperties aTableCellProperties)
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
        if (this.aTableCellProperties.RightBorderLineProperties is null)
        {
            var aSolidFill = new A.SolidFill
            {
                SchemeColor = new A.SchemeColor { Val = A.SchemeColorValues.Text1 }
            };
            this.aTableCellProperties.RightBorderLineProperties = new A.RightBorderLineProperties();
            this.aTableCellProperties.RightBorderLineProperties.AppendChild(aSolidFill);
        }
        
        var emus = new Points(points).AsEmus();
        this.aTableCellProperties.RightBorderLineProperties!.Width = new Int32Value((int)emus);
    }

    private float GetWidth()
    {
        if (this.aTableCellProperties.RightBorderLineProperties is null)
        {
            return 1; // default value
        }

        var emus = this.aTableCellProperties.RightBorderLineProperties!.Width!.Value;
        
        return new Emus(emus).AsPoints();
    }
}