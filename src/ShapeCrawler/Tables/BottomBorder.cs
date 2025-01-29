using DocumentFormat.OpenXml;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Tables;

internal class BottomBorder : IBorder
{
    private readonly A.TableCellProperties aTableCellProperties;
    private A.SolidFill? aSolidFill;

    internal BottomBorder(A.TableCellProperties aTableCellProperties)
    {
        this.aTableCellProperties = aTableCellProperties;

        if (this.aTableCellProperties.BottomBorderLineProperties is not null)
        {
            this.aSolidFill = this.aTableCellProperties.BottomBorderLineProperties.GetFirstChild<A.SolidFill>();
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
        this.aTableCellProperties.BottomBorderLineProperties ??= new A.BottomBorderLineProperties
        {
            Width = new Int32Value(12700) // 1 * 12700 => emu to point
        };

        this.aSolidFill ??= this.aTableCellProperties.BottomBorderLineProperties.GetFirstChild<A.SolidFill>();

        if (this.aSolidFill is null)
        {
            this.aSolidFill = new A.SolidFill();
            this.aTableCellProperties.BottomBorderLineProperties.AppendChild(this.aSolidFill);
        }

        this.aSolidFill.RgbColorModelHex ??= new A.RgbColorModelHex();

        this.aSolidFill.RgbColorModelHex.Val = new HexBinaryValue(color);
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
        
        var emus = new Points((decimal)points).AsEmus();
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