using System.Collections.Generic;
using ShapeCrawler.Enums;
using ShapeCrawler.Models.SlideComponents;
using ShapeCrawler.Models.SlideComponents.Chart;
using ShapeCrawler.Settings;
using SlideDotNet.Models.TableComponents;
using OleObject = ShapeCrawler.Models.SlideComponents.OleObject;

namespace ShapeCrawler.Factories.Builders
{
    /// <summary>
    /// Represents a shape builder.
    /// </summary>
    public interface IShapeBuilder
    {
        /// <summary>
        /// Builds a shape with OLE object content.
        /// </summary>
        ShapeEx WithOle(ILocation innerTransform, ShapeContext spContext, OleObject ole);

        /// <summary>
        /// Builds a shape with picture content.
        /// </summary>
        ShapeEx WithPicture(ILocation innerTransform, ShapeContext spContext, Picture picture, GeometryType geometry);

        /// <summary>
        /// Builds a AutoShape.
        /// </summary>
        ShapeEx WithAutoShape(ILocation innerTransform, ShapeContext spContext, GeometryType geometry);

        /// <summary>
        /// Builds a shape with table content.
        /// </summary>
        ShapeEx WithTable(ILocation innerTransform, ShapeContext spContext, TableEx table);

        /// <summary>
        /// Builds a shape with OLE object content.
        /// </summary>
        ShapeEx WithChart(ILocation innerTransform, ShapeContext spContext, ChartEx chart);

        /// <summary>
        /// Builds a group shape which has grouped shape items.
        /// </summary>
        ShapeEx WithGroup(ILocation innerTransform, ShapeContext spContext, IList<ShapeEx> groupedShapes);
    }
}
