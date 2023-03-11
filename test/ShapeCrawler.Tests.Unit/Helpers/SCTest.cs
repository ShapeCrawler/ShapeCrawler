using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using ClosedXML.Excel;
using ShapeCrawler.Tests.Shared;

namespace ShapeCrawler.Tests.Unit.Helpers;

public abstract class SCTest
{
    public static List<string> HelperAssets = new()
    {
        "autoshape-grouping.pptx",
        "001.pptx",
        "table-case001.pptx",
        "autoshape-case005_text-frame.pptx",
        "009_table.pptx"
    };

    protected static T GetShape<T>(string presentation, int slideNumber, int shapeId)
    {
        var scPresentation = GetPresentationFromAssembly(presentation);
        var slide = scPresentation.Slides[slideNumber - 1];
        var shape = slide.Shapes.First(sp => sp.Id == shapeId);

        return (T)shape;
    }

    protected static T GetWorksheetCellValue<T>(byte[] workbookByteArray, string cellAddress)
    {
        var stream = new MemoryStream(workbookByteArray);
        var xlWorkbook = new XLWorkbook(stream);
        var cellValue = xlWorkbook.Worksheets.First().Cell(cellAddress).Value;

        return (T)cellValue;
    }

    protected static byte[] GetTestBytes(string fileName)
    {
        return GetTestStream(fileName).ToArray();
    }

    protected static MemoryStream GetTestStream(string fileName)
    {
        if (fileName == "009_table.pptx")
        {
            return TestHelperShared.GetStream(fileName);
        }
        var assembly = Assembly.GetExecutingAssembly();
        return assembly.GetResourceStream(fileName);
    }

    protected static IPresentation SaveAndOpenPresentation(IPresentation presentation)
    {
        var stream = new MemoryStream();
        presentation.SaveAs(stream);

        return SCPresentation.Open(stream);
    }

    private static IPresentation GetPresentationFromAssembly(string fileName)
    {
        var stream = GetTestStream(fileName);

        return SCPresentation.Open(stream);
    }
}

public class ValidationError
{
    public ValidationError(string description, string path)
    {
        this.Description = description;
        this.Path = path;
    }

    public string Path { get; }

    public string Description { get; }
}