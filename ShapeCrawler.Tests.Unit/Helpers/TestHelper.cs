using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ShapeCrawler.Tests.Unit.Helpers
{
    public static class TestHelper
    {
        static TestHelper()
        {
            var bm = new Bitmap(100, 100);
            if (bm.HorizontalResolution == 0)
            {
                // Set default resolution
                bm.SetResolution(96, 96);
            }

            HorizontalResolution = bm.HorizontalResolution;
            VerticalResolution = bm.VerticalResolution;
        }

        public static IParagraph GetParagraph(SCPresentation presentation, SlideElementQuery paragraphRequest)
        {
            IAutoShape autoShape = presentation.Slides[paragraphRequest.SlideIndex]
                .Shapes.First(sp => sp.Id == paragraphRequest.ShapeId) as IAutoShape;
            return autoShape.TextBox.Paragraphs[paragraphRequest.ParagraphIndex];
        }

        public static IParagraph GetParagraph(MemoryStream presentationStream, SlideElementQuery paragraphRequest)
        {
            SCPresentation presentation = SCPresentation.Open(presentationStream, false);
            IAutoShape autoShape = presentation.Slides[paragraphRequest.SlideIndex]
                .Shapes.First(sp => sp.Id == paragraphRequest.ShapeId) as IAutoShape;
            return autoShape.TextBox.Paragraphs[paragraphRequest.ParagraphIndex];
        }

        public static IPortion GetParagraphPortion(SCPresentation presentation, SlideElementQuery elementRequest)
        {
            IAutoShape autoShape = (IAutoShape)presentation.Slides[elementRequest.SlideIndex].Shapes.First(sp => sp.Id == elementRequest.ShapeId);
            
            return autoShape.TextBox.Paragraphs[elementRequest.ParagraphIndex].Portions[elementRequest.PortionIndex];
        }

        public static MemoryStream ToResizeableStream(this byte[] byteArray)
        {
            var stream = new MemoryStream();
            stream.Write(byteArray, 0, byteArray.Length);

            return stream;
        }

        public static readonly float HorizontalResolution;
        
        public static readonly float VerticalResolution;

        public static IAutoShape GetAutoShape(string presentationName, int slideNumber, int shapeId)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var path = assembly.GetManifestResourceNames().First(r => r.EndsWith(presentationName, StringComparison.Ordinal));
            var stream = assembly.GetManifestResourceStream(path);
            var mStream = new MemoryStream();
            stream.CopyTo(mStream);
            var presentation = SCPresentation.Open(mStream, true);
            var slide = presentation.Slides.First(s => s.Number == slideNumber);
            var shape = slide.Shapes.First(sp => sp.Id == shapeId);

            return (IAutoShape) shape;
        }
    }
}