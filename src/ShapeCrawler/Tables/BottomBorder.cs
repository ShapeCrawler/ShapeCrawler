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

    public decimal Width
    {
        get
        {
            if (this.aTableCellProperties.BottomBorderLineProperties is null)
            {
                return 1; // default value
            }

            var emus = this.aTableCellProperties.BottomBorderLineProperties!.Width!.Value;
        
            return new Emus(emus).AsPoints();
        }
        set => this.UpdateWidth(value);
    }

    public string? Color { get => this.GetColor(); set => this.SetColor(value!); }

    private string? GetColor()
    {
        return this.aTableCellProperties.BottomBorderLineProperties?.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val;
    }

    private void SetColor(string color)
    {
        this.aTableCellProperties.BottomBorderLineProperties ??= new A.BottomBorderLineProperties
        {
            Width = (Int32Value)new Points(1).AsEmus()
        };

        var aSolidFill = this.aTableCellProperties.BottomBorderLineProperties.GetFirstChild<A.SolidFill>();

        if (aSolidFill is null)
        {
            aSolidFill = new A.SolidFill();
            this.aTableCellProperties.BottomBorderLineProperties.AppendChild(aSolidFill);
        }

        aSolidFill.RgbColorModelHex ??= new A.RgbColorModelHex();

        aSolidFill.RgbColorModelHex.Val = new HexBinaryValue(color);
    }

    private void UpdateWidth(decimal points)
    {
        if (this.aTableCellProperties.BottomBorderLineProperties is null)
        {
            var aSolidFill = new A.SolidFill
            {
                RgbColorModelHex = new() { Val = "000000" }
            };

            this.aTableCellProperties.BottomBorderLineProperties = new A.BottomBorderLineProperties();
            this.aTableCellProperties.BottomBorderLineProperties.AppendChild(aSolidFill);
        }
        
        var emus = new Points(points).AsEmus();
        this.aTableCellProperties.BottomBorderLineProperties!.Width = new Int32Value((int)emus);
    }
}