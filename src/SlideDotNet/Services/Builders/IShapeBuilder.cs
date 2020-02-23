using System.Collections.Generic;
using DocumentFormat.OpenXml;
using SlideDotNet.Models;
using SlideDotNet.Models.Settings;
using SlideDotNet.Models.SlideComponents;
using SlideDotNet.Models.SlideComponents.Chart;
using SlideXML.Models.SlideComponents;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideDotNet.Services.Builders
{
    /// <summary>
    /// Represents a shape builder.
    /// </summary>
    public interface IShapeBuilder
    {
        /// <summary>
        /// Builds a shape with OLE object content.
        /// </summary>
        ShapeEx WithOle(Location location, IShapeContext spContext, OleObject ole);

        /// <summary>
        /// Builds a shape with picture content.
        /// </summary>
        ShapeEx WithPicture(Location location, IShapeContext spContext, Picture picture);

        /// <summary>
        /// Builds a AutoShape.
        /// </summary>
        ShapeEx WithAutoShape(Location location, IShapeContext spContext);

        /// <summary>
        /// Builds a shape with table content.
        /// </summary>
        ShapeEx WithTable(Location location, IShapeContext spContext, TableEx table);

        /// <summary>
        /// Builds a shape with OLE object content.
        /// </summary>
        ShapeEx WithChart(Location location, IShapeContext spContext, ChartEx chart);

        /// <summary>
        /// Builds a group shape which has grouped shape items.
        /// </summary>
        ShapeEx WithGroup(Location location, IShapeContext spContext, IEnumerable<ShapeEx> groupedShapes);
    }
}
