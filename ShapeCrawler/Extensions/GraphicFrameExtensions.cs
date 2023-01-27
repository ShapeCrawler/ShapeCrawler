using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Extensions;

internal static class GraphicFrameExtensions
{
    internal static A.Table GetATable(this P.GraphicFrame pGraphicFrame)
    {
        return pGraphicFrame.GetFirstChild<A.Graphic>() !.GraphicData!.GetFirstChild<A.Table>() !;
    }
    
    internal static void AddAXfrm(this P.GraphicFrame pGraphicFrame, long xEmu, long yEmu, long wEmu, long hEmu)
    {
        var pXfrm = pGraphicFrame.Transform;
        pXfrm?.Remove();
        
        pXfrm = new P.Transform();
        pGraphicFrame.Append(pXfrm);

        var aOff = new A.Offset
        {
            X = xEmu,
            Y = yEmu
        };
        pXfrm.Append(aOff);
        
        var aExt = new A.Extents
        {
            Cx = wEmu,
            Cy = hEmu
        };
        pXfrm.Append(aExt);
    }
}