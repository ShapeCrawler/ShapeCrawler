using System;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace SlideDotNet.Extensions
{
    /// <summary>
    /// Contains extension methods for <see cref="P.GraphicFrame "/> class object.
    /// </summary>
    public static class GraphicFrameExtensions
    {
        private const string ChartUri = "http://schemas.openxmlformats.org/drawingml/2006/chart";

        /// <summary>
        /// Has <see cref="P.GraphicFrame"/> instance chart.
        /// </summary>
        /// <param name="grFrame"></param>        
        public static bool HasChart(this P.GraphicFrame grFrame)
        {
            var grData = grFrame.GetFirstChild<A.Graphic>().GetFirstChild<A.GraphicData>();
            
            return grData.Uri.Value.Equals(ChartUri, StringComparison.Ordinal);
        }
    }
}
