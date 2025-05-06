using DocumentFormat.OpenXml;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Tables;

internal class BottomBorder(A.TableCellProperties aTableCellProperties) : IBorder
{
    public decimal Width
    {
        get
        {
            if (aTableCellProperties.BottomBorderLineProperties is null)
            {
                return 1; // default value
            }

            var emus = aTableCellProperties.BottomBorderLineProperties!.Width!.Value;

            return new Emus(emus).AsPoints();
        }
        set => this.UpdateWidth(value);
    }

    public string? Color { get => this.GetColor(); set => this.SetColor(value!); }

    private string? GetColor() => aTableCellProperties.BottomBorderLineProperties?.GetFirstChild<A.SolidFill>()
        ?.RgbColorModelHex?.Val;

    private void SetColor(string color)
    {
        aTableCellProperties.BottomBorderLineProperties ??= new A.BottomBorderLineProperties
        {
            Width = (Int32Value)new Points(1).AsEmus()
        };

        var aSolidFill = aTableCellProperties.BottomBorderLineProperties.GetFirstChild<A.SolidFill>();

        if (aSolidFill is null)
        {
            aSolidFill = new A.SolidFill();
            aTableCellProperties.BottomBorderLineProperties.AppendChild(aSolidFill);
        }

        aSolidFill.RgbColorModelHex ??= new A.RgbColorModelHex();

        aSolidFill.RgbColorModelHex.Val = new HexBinaryValue(color);
    }

    private void UpdateWidth(decimal points)
    {
        if (aTableCellProperties.BottomBorderLineProperties is null)
        {
            var aSolidFill = new A.SolidFill { RgbColorModelHex = new() { Val = "000000" } };

            aTableCellProperties.BottomBorderLineProperties = new A.BottomBorderLineProperties();
            aTableCellProperties.BottomBorderLineProperties.AppendChild(aSolidFill);
        }

        var emus = new Points(points).AsEmus();
        aTableCellProperties.BottomBorderLineProperties!.Width = new Int32Value((int)emus);
    }
}