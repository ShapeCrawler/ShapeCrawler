using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Tables;

internal class RightBorder(A.TableCellProperties aTableCellProperties): IBorder
{
    public decimal Width
    {
        get
        {
            if (aTableCellProperties.RightBorderLineProperties is null)
            {
                return 1; // default value
            }

            var emus = aTableCellProperties.RightBorderLineProperties!.Width!.Value;

            return new Emus(emus).AsPoints();
        }

        set
        {
            if (aTableCellProperties.RightBorderLineProperties is null)
            {
                var aSolidFill = new A.SolidFill { RgbColorModelHex = new A.RgbColorModelHex { Val = "000000" } };

                aTableCellProperties.RightBorderLineProperties = new A.RightBorderLineProperties();
                aTableCellProperties.RightBorderLineProperties.AppendChild(aSolidFill);
            }

            var emus = new Points(value).AsEmus();
            aTableCellProperties.RightBorderLineProperties!.Width = new Int32Value((int)emus);
        }
    }

    public string? Color
    {
        get => aTableCellProperties.RightBorderLineProperties?.GetFirstChild<SolidFill>()
            ?.RgbColorModelHex?.Val;
        set
        {
            aTableCellProperties.RightBorderLineProperties ??= new A.RightBorderLineProperties
            {
                Width = (Int32Value)new Points(1).AsEmus()
            };

            var solidFill = aTableCellProperties.RightBorderLineProperties.GetFirstChild<A.SolidFill>();
            if (solidFill is null)
            {
                solidFill = new A.SolidFill();
                aTableCellProperties.RightBorderLineProperties.AppendChild(solidFill);
            }

            solidFill.RgbColorModelHex ??= new A.RgbColorModelHex();
            solidFill.RgbColorModelHex.Val = new HexBinaryValue(value);
        }
    }
}