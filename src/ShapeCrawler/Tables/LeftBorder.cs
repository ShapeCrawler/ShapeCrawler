using DocumentFormat.OpenXml;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Tables;

internal class LeftBorder : IBorder
{
    private readonly A.TableCellProperties aTableCellProperties;
    private A.SolidFill? aSolidFill;

    internal LeftBorder(A.TableCellProperties aTableCellProperties)
    {
        this.aTableCellProperties = aTableCellProperties;

        if(this.aTableCellProperties.LeftBorderLineProperties is not null)
        {
            this.aSolidFill = this.aTableCellProperties.LeftBorderLineProperties.GetFirstChild<A.SolidFill>();
        }
        else
        {
            this.aSolidFill = null;
        }
    }

    public float Width
    {
        get => this.GetWidth();
        set => this.UpdateWidth(value);
    }

    public string? Color { get => this.GetColor(); set => this.SetColor(value!); }

    private string? GetColor()
    {
        if (this.aSolidFill is null || this.aSolidFill.RgbColorModelHex is null)
        {
            return null;
        }

        return this.aSolidFill.RgbColorModelHex.Val;
    }

    private void SetColor(string color)
    {
        this.aTableCellProperties.LeftBorderLineProperties ??= new A.LeftBorderLineProperties
        {
            Width = new Int32Value(12700) // 1 * 12700 => emu to point
        };

        this.aSolidFill ??= this.aTableCellProperties.LeftBorderLineProperties.GetFirstChild<A.SolidFill>();

        if (this.aSolidFill is null)
        {
            this.aSolidFill = new A.SolidFill();
            this.aTableCellProperties.LeftBorderLineProperties.AppendChild(this.aSolidFill);
        }

        this.aSolidFill.RgbColorModelHex ??= new A.RgbColorModelHex();

        this.aSolidFill.RgbColorModelHex.Val = new HexBinaryValue(color);
    }

    private void UpdateWidth(float points)
    {
        if (this.aTableCellProperties.LeftBorderLineProperties is null)
        {
            this.aSolidFill = new A.SolidFill
            {
                RgbColorModelHex = new A.RgbColorModelHex { Val = "000000" } // black by default 
            };

            this.aTableCellProperties.LeftBorderLineProperties = new A.LeftBorderLineProperties();
            this.aTableCellProperties.LeftBorderLineProperties.AppendChild(this.aSolidFill);
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