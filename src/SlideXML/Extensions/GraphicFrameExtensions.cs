using System.Linq;
using LogicNull.Utilities;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace SlideXML.Extensions
{
    /// <summary>
    /// Contains extension methods for <see cref="P.GraphicFrame "/> class object.
    /// </summary>
    public static class GraphicFrameExtensions
    {
        /// <summary>
        /// Has <see cref="P.GraphicFrame"/> instance chart.
        /// </summary>
        /// <param name="grFrame"></param>        
        public static bool HasChart(this P.GraphicFrame grFrame)
        {
            Check.NotNull(grFrame, nameof(grFrame));

            var grData = grFrame.Descendants<A.GraphicData>().Single();
            var endsWithChart = grData?.Uri?.Value?.EndsWith("chart");

            return endsWithChart != null && endsWithChart != false;
        }
    }
}
