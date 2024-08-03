using DocumentFormat.OpenXml;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Tables;

internal class BottomBorder : IBorder
{
    private readonly A.TableCellProperties aTableCellProperties;

    internal BottomBorder(A.TableCellProperties aTableCellProperties)
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
        if (this.aTableCellProperties.BottomBorderLineProperties is null)
        {
            var aSolidFill = new A.SolidFill
            {
                SchemeColor = new A.SchemeColor { Val = A.SchemeColorValues.Text1 }
            };
            this.aTableCellProperties.BottomBorderLineProperties = new A.BottomBorderLineProperties();
            this.aTableCellProperties.BottomBorderLineProperties.AppendChild(aSolidFill);
        }
        
        var emus = new Points(points).AsEmus();
        this.aTableCellProperties.BottomBorderLineProperties!.Width = new Int32Value((int)emus);
    }

    private float GetWidth()
    {
        if (this.aTableCellProperties.BottomBorderLineProperties is null)
        {
            return 1; // default value
        }

        var emus = this.aTableCellProperties.BottomBorderLineProperties!.Width!.Value;
        
        return new Emus(emus).AsPoints();
    }
}