using DocumentFormat.OpenXml;
using ShapeCrawler.Tables;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Texts;

internal sealed class TextBoxMargins(OpenXmlElement textBody)
{
    internal decimal Left
    {
        get
        {
            return new LeftRightMargin(textBody.GetFirstChild<A.BodyProperties>()!.LeftInset).Value;
        }

        set
        {
            var bodyProperties = textBody.GetFirstChild<A.BodyProperties>()!;
            var emu = new Points(value).AsEmus();
            bodyProperties.LeftInset = new Int32Value((int)emu);
        }
    }

    internal decimal Right
    {
        get => new LeftRightMargin(textBody.GetFirstChild<A.BodyProperties>()!.RightInset).Value;
        set
        {
            var bodyProperties = textBody.GetFirstChild<A.BodyProperties>()!;
            var emu = new Points(value).AsEmus();
            bodyProperties.RightInset = new Int32Value((int)emu);
        }
    }

    internal decimal Top
    {
        get => new TopBottomMargin(textBody.GetFirstChild<A.BodyProperties>()!.TopInset).Value;
        set
        {
            var bodyProperties = textBody.GetFirstChild<A.BodyProperties>()!;
            var emu = new Points(value).AsEmus();
            bodyProperties.TopInset = new Int32Value((int)emu);
        }
    }

    internal decimal Bottom
    {
        get => new TopBottomMargin(textBody.GetFirstChild<A.BodyProperties>()!.BottomInset).Value;
        set
        {
            var bodyProperties = textBody.GetFirstChild<A.BodyProperties>()!;
            var emu = new Points(value).AsEmus();
            bodyProperties.BottomInset = new Int32Value((int)emu);
        }
    }
}