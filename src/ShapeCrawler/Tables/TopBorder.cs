using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Tables;

internal class TopBorder(A.TableCellProperties aTableCellProperties): IBorder
{
    public decimal Width
    {
        get => this.GetWidth();
        set => this.UpdateWidth(value);
    }

    public string? Color { get => this.GetColor(); set => this.SetColor(value!); }

    private string? GetColor()
    {
        return aTableCellProperties.TopBorderLineProperties?.GetFirstChild<SolidFill>()?.RgbColorModelHex?.Val;
    }

    private void SetColor(string color)
    {
        aTableCellProperties.TopBorderLineProperties ??= new A.TopBorderLineProperties
        {
            Width = (Int32Value)new Points(1).AsEmus()
        };

        var solidFill = aTableCellProperties.TopBorderLineProperties.GetFirstChild<A.SolidFill>();

        if (solidFill is null)
        {
            solidFill = new A.SolidFill();
            aTableCellProperties.TopBorderLineProperties.AppendChild(solidFill);
        }

        solidFill.RgbColorModelHex ??= new A.RgbColorModelHex();

        solidFill.RgbColorModelHex.Val = new HexBinaryValue(color);
    }

    private void UpdateWidth(decimal points)
    {
        if (aTableCellProperties.TopBorderLineProperties is null)
        {
            var aSolidFill = new A.SolidFill
            {
                SchemeColor = new A.SchemeColor { Val = SchemeColorValues.Text1 }
            };
            aTableCellProperties.TopBorderLineProperties = new A.TopBorderLineProperties();
            aTableCellProperties.TopBorderLineProperties.AppendChild(aSolidFill);
        }
        
        var emus = new Points(points).AsEmus();
        aTableCellProperties.TopBorderLineProperties.Width = new Int32Value((int)emus);
    }

    private decimal GetWidth()
    {
        if (aTableCellProperties.TopBorderLineProperties is null)
        {
            return 1; // default value
        }

        var emus = aTableCellProperties.TopBorderLineProperties!.Width!.Value;
        
        return new Emus(emus).AsPoints();
    }
}