using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
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

    public string? Color { get => this.GetColor(); set => this.SetColor(value!); }

    private string? GetColor()
    {
        return this.aTableCellProperties?.RightBorderLineProperties?.GetFirstChild<SolidFill>()?.RgbColorModelHex?.Val;
    }

    private void SetColor(string color)
    {
        this.aTableCellProperties.RightBorderLineProperties ??= new A.RightBorderLineProperties
        {
            Width = (Int32Value)new Points(1).AsEmus()
        };

        var solidFill = this.aTableCellProperties.RightBorderLineProperties.GetFirstChild<A.SolidFill>();

        if (solidFill is null)
        {
            solidFill = new A.SolidFill();
            this.aTableCellProperties.RightBorderLineProperties.AppendChild(solidFill);
        }

        solidFill.RgbColorModelHex ??= new A.RgbColorModelHex();

        solidFill.RgbColorModelHex.Val = new HexBinaryValue(color);
    }

    private void UpdateWidth(float points)
    {
        if (this.aTableCellProperties.RightBorderLineProperties is null)
        {
            var aSolidFill = new A.SolidFill
            {
                RgbColorModelHex = new A.RgbColorModelHex { Val = "000000" }
            };

            this.aTableCellProperties.RightBorderLineProperties = new A.RightBorderLineProperties();
            this.aTableCellProperties.RightBorderLineProperties.AppendChild(aSolidFill);
        }
        
        var emus = new Points((decimal)points).AsEmus();
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