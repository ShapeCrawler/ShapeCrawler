using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.Enums;
using ShapeCrawler.Models;
using ShapeCrawler.SlideMaster;

// ReSharper disable CheckNamespace

namespace ShapeCrawler
{
    public class OleObjectSc : BaseShape
    {
        public OleObjectSc(SlideMasterSc slideMasterSc, GraphicFrame pGraphicFrame) : base(pGraphicFrame)
        {
            
        }

        public override long Width => throw new System.NotImplementedException();

        public override long Height => throw new System.NotImplementedException();
        public override GeometryType GeometryType { get; }

        public override long X => throw new System.NotImplementedException();

        public override long Y => throw new System.NotImplementedException();
    }
}