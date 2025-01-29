using DocumentFormat.OpenXml;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Tables;

internal class TopBorder : IBorder
{
    private readonly DocumentFormat.OpenXml.Drawing.TableCellProperties aTableCellProperties;
    private A.SolidFill? aSolidFill;

    internal TopBorder(DocumentFormat.OpenXml.Drawing.TableCellProperties aTableCellProperties)
    {
        this.aTableCellProperties = aTableCellProperties;

        if (this.aTableCellProperties.TopBorderLineProperties is not null)
        {
            this.aSolidFill = this.aTableCellProperties.TopBorderLineProperties.GetFirstChild<A.SolidFill>();
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
        this.aTableCellProperties.TopBorderLineProperties ??= new A.TopBorderLineProperties
        {
            Width = new Int32Value(12700) // 1 * 12700 => emu to point
        };

        this.aSolidFill ??= this.aTableCellProperties.TopBorderLineProperties.GetFirstChild<A.SolidFill>();

        if (this.aSolidFill is null)
        {
            this.aSolidFill = new A.SolidFill();
            this.aTableCellProperties.TopBorderLineProperties.AppendChild(this.aSolidFill);
        }

        this.aSolidFill.RgbColorModelHex ??= new A.RgbColorModelHex();

        this.aSolidFill.RgbColorModelHex.Val = new HexBinaryValue(color);
    }

    private void UpdateWidth(float points)
    {
        if (this.aTableCellProperties.TopBorderLineProperties is null)
        {
            var aSolidFill = new DocumentFormat.OpenXml.Drawing.SolidFill
            {
                SchemeColor = new DocumentFormat.OpenXml.Drawing.SchemeColor { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.Text1 }
            };
            this.aTableCellProperties.TopBorderLineProperties = new DocumentFormat.OpenXml.Drawing.TopBorderLineProperties();
            this.aTableCellProperties.TopBorderLineProperties.AppendChild(aSolidFill);
        }
        
        var emus = new Points((decimal)points).AsEmus();
        this.aTableCellProperties.TopBorderLineProperties.Width = new Int32Value((int)emus);
    }

    private float GetWidth()
    {
        if (this.aTableCellProperties.TopBorderLineProperties is null)
        {
            return 1; // default value
        }

        var emus = this.aTableCellProperties.TopBorderLineProperties!.Width!.Value;
        
        return new Emus(emus).AsPoints();
    }
}