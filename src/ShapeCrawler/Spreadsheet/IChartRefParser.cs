using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Spreadsheet
{
    public interface IChartRefParser
    {
        IList<double> GetNumbers(DocumentFormat.OpenXml.Drawing.Charts.NumberReference numRef, ChartPart chartPart);

        string GetSingleString(DocumentFormat.OpenXml.Drawing.Charts.StringReference strRef, ChartPart chartPart);
    }
}