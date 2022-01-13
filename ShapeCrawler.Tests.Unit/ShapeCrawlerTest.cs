using System;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ShapeCrawler.Tests.Unit
{
    public abstract class ShapeCrawlerTest
    {
        protected IAutoShape GetAutoShape(string presentation, int slideNumber, int shapeId)
        {
            var scPresentation = this.GetPresentation(presentation);
            var slide = scPresentation.Slides.First(s => s.Number == slideNumber);
            var shape = slide.Shapes.First(sp => sp.Id == shapeId);

            return (IAutoShape) shape;
        }

        protected ISlide GetSlide(string presentation, int slideNumber)
        {
            var scPresentation = this.GetPresentation(presentation);

            return scPresentation.Slides[slideNumber - 1];
        }
        
        private SCPresentation GetPresentation(string fileName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var path = assembly.GetManifestResourceNames().First(r => r.EndsWith(fileName, StringComparison.Ordinal));
            var stream = assembly.GetManifestResourceStream(path);
            var mStream = new MemoryStream();
            stream.CopyTo(mStream);
            
            return SCPresentation.Open(mStream, true);
        }
    }
}