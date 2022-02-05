using System;
using System.IO;
using System.Linq;
using System.Reflection;
using ClosedXML.Excel;

namespace ShapeCrawler.Tests.Unit
{
    public abstract class ShapeCrawlerTest
    {
        protected T GetShape<T>(Stream presentation, int slideNumber, int shapeId)
        {
            var scPresentation = SCPresentation.Open(presentation, false);
            
            var slide = scPresentation.Slides[slideNumber - 1];
            var shape = slide.Shapes.First(sp => sp.Id == shapeId);

            return (T) shape;
        }
        
        protected T GetShape<T>(string presentation, int slideNumber, int shapeId)
        {
            var scPresentation = GetPresentationFromAssembly(presentation);
            var slide = scPresentation.Slides[slideNumber - 1];
            var shape = slide.Shapes.First(sp => sp.Id == shapeId);

            return (T) shape;
        }
        
        protected IAutoShape GetAutoShape(string presentation, int slideNumber, int shapeId)
        {
            var scPresentation = GetPresentationFromAssembly(presentation);
            var slide = scPresentation.Slides.First(s => s.Number == slideNumber);
            var shape = slide.Shapes.First(sp => sp.Id == shapeId);

            return (IAutoShape) shape;
        }

        protected T GetWorksheetCellValue<T>(byte[] workbookByteArray, string cellAddress)
        {
            var stream = new MemoryStream(workbookByteArray);
            var xlWorkbook = new XLWorkbook(stream);
            var cellValue = xlWorkbook.Worksheets.First().Cell(cellAddress).Value;

            return (T)cellValue;
        }

        protected static Stream GetPptxStream(string fileName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var path = assembly.GetManifestResourceNames().First(r => r.EndsWith(fileName, StringComparison.Ordinal));
            var stream = assembly.GetManifestResourceStream(path);
            var mStream = new MemoryStream();
            stream!.CopyTo(mStream);

            return mStream;
        }

        private IPresentation GetPresentationFromAssembly(string fileName)
        {
            var stream = GetPptxStream(fileName);
            
            return SCPresentation.Open(stream, true);
        }
    }
}