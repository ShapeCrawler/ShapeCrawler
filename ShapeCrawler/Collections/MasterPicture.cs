using System;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMaster;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Collections
{
    internal class MasterPicture : MasterShape, IShape
    {
        public MasterPicture(SlideMasterSc slideMaster, P.Picture pPicture) : base(slideMaster, pPicture)
        {
        }

        public string Name { get; }
        public bool Hidden { get; }
    }
}