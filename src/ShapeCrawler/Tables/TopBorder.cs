using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Tables;

internal class TopBorder : IBorder
{
    private readonly DocumentFormat.OpenXml.Drawing.TableCellProperties aTableCellProperties;
  
    internal TopBorder(DocumentFormat.OpenXml.Drawing.TableCellProperties aTableCellProperties)
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
        return this.aTableCellProperties?.TopBorderLineProperties?.GetFirstChild<SolidFill>()?.RgbColorModelHex?.Val;
    }

    private void SetColor(string color)
    {
        this.aTableCellProperties.TopBorderLineProperties ??= new A.TopBorderLineProperties
        {
            Width = (Int32Value)new Points(1).AsEmus()
        };

        var solidFill = this.aTableCellProperties.TopBorderLineProperties.GetFirstChild<A.SolidFill>();

        if (solidFill is null)
        {
            solidFill = new A.SolidFill();
            this.aTableCellProperties.TopBorderLineProperties.AppendChild(solidFill);
        }

        solidFill.RgbColorModelHex ??= new A.RgbColorModelHex();

        solidFill.RgbColorModelHex.Val = new HexBinaryValue(color);
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