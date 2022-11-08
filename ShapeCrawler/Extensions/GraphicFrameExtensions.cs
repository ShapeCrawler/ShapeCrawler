using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Extensions;

internal static class GraphicFrameExtensions
{
    internal static A.Table GetATable(this P.GraphicFrame pGraphicFrame)
    {
        return pGraphicFrame.GetFirstChild<A.Graphic>()!.GraphicData!.GetFirstChild<A.Table>()!;
    }
}