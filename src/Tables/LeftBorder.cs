using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Tables;

internal class LeftBorder(A.TableCellProperties aTableCellProperties): IBorder
{
    public decimal Width
    {
        get
        {
            if (aTableCellProperties.LeftBorderLineProperties is null)
            {
                return 1; // default value
            }

            var emus = aTableCellProperties.LeftBorderLineProperties!.Width!.Value;

            return new Emus(emus).AsPoints();
        }

        set
        {
            if (aTableCellProperties.LeftBorderLineProperties is null)
            {
                var solidFill = new A.SolidFill
                {
                    RgbColorModelHex = new A.RgbColorModelHex { Val = "000000" } // black by default 
                };

                aTableCellProperties.LeftBorderLineProperties = new A.LeftBorderLineProperties();
                aTableCellProperties.LeftBorderLineProperties.AppendChild(solidFill);
            }

            var emus = new Points(value).AsEmus();
            aTableCellProperties.LeftBorderLineProperties!.Width = new Int32Value((int)emus);
        }
    }

    public string? Color
    {
        get => aTableCellProperties.LeftBorderLineProperties?.GetFirstChild<SolidFill>()?.RgbColorModelHex?.Val;
        set
        {
            aTableCellProperties.LeftBorderLineProperties ??= new A.LeftBorderLineProperties
            {
                Width = (Int32Value)new Points(1).AsEmus()
            };

            var solidFill = aTableCellProperties.LeftBorderLineProperties.GetFirstChild<A.SolidFill>();

            if (solidFill is null)
            {
                solidFill = new A.SolidFill();
                aTableCellProperties.LeftBorderLineProperties.AppendChild(solidFill);
            }

            solidFill.RgbColorModelHex ??= new A.RgbColorModelHex();

            solidFill.RgbColorModelHex.Val = new HexBinaryValue(value);
        }
    }
}