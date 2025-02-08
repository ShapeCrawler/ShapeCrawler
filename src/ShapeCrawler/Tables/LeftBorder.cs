using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
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

    public string? Color { get => this.GetColor(); set => this.SetColor(value!); }

    private string? GetColor()
    {
        return this.aTableCellProperties?.LeftBorderLineProperties?.GetFirstChild<SolidFill>()?.RgbColorModelHex?.Val;
    }

    private void SetColor(string color)
    {
        this.aTableCellProperties.LeftBorderLineProperties ??= new A.LeftBorderLineProperties
        {
            Width = (Int32Value)new Points(1).AsEmus()
        };

        var solidFill = this.aTableCellProperties.LeftBorderLineProperties.GetFirstChild<A.SolidFill>();

        if (solidFill is null)
        {
            solidFill = new A.SolidFill();
            this.aTableCellProperties.LeftBorderLineProperties.AppendChild(solidFill);
        }

        solidFill.RgbColorModelHex ??= new A.RgbColorModelHex();

        solidFill.RgbColorModelHex.Val = new HexBinaryValue(color);
    }

    private void UpdateWidth(float points)
    {
        if (this.aTableCellProperties.LeftBorderLineProperties is null)
        {
            var solidFill = new A.SolidFill
            {
                RgbColorModelHex = new A.RgbColorModelHex { Val = "000000" } // black by default 
            };

            this.aTableCellProperties.LeftBorderLineProperties = new A.LeftBorderLineProperties();
            this.aTableCellProperties.LeftBorderLineProperties.AppendChild(solidFill);
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