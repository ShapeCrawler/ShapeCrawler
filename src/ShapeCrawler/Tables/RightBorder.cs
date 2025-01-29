using DocumentFormat.OpenXml;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Tables;

internal class RightBorder : IBorder
{
    private readonly A.TableCellProperties aTableCellProperties;
    private A.SolidFill? aSolidFill;

    internal RightBorder(A.TableCellProperties aTableCellProperties)
    {
        this.aTableCellProperties = aTableCellProperties;

        if (this.aTableCellProperties.RightBorderLineProperties is not null)
        {
            this.aSolidFill = this.aTableCellProperties.RightBorderLineProperties.GetFirstChild<A.SolidFill>();
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
        this.aTableCellProperties.RightBorderLineProperties ??= new A.RightBorderLineProperties
        {
            Width = new Int32Value(12700) // 1 * 12700 => emu to point
        };

        this.aSolidFill ??= this.aTableCellProperties.RightBorderLineProperties.GetFirstChild<A.SolidFill>();

        if (this.aSolidFill is null)
        {
            this.aSolidFill = new A.SolidFill();
            this.aTableCellProperties.RightBorderLineProperties.AppendChild(this.aSolidFill);
        }

        this.aSolidFill.RgbColorModelHex ??= new A.RgbColorModelHex();

        this.aSolidFill.RgbColorModelHex.Val = new HexBinaryValue(color);
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