using System.Collections.Generic;
using ShapeCrawler.Enums;
using ShapeCrawler.Models.Settings;
using ShapeCrawler.Models.SlideComponents;
using ShapeCrawler.Models.SlideComponents.Chart;
using SlideDotNet.Models.TableComponents;
using OleObject = ShapeCrawler.Models.SlideComponents.OleObject;

namespace ShapeCrawler.Services.Builders
{
    /// <summary>
    /// Represents a shape builder.
    /// </summary>
    public interface IShapeBuilder
    {
        /// <summary>
        /// Builds a shape with OLE object content.
        /// </summary>
        Shape WithOle(ILocation innerTransform, IShapeContext spContext, OleObject ole);

        /// <summary>
        /// Builds a shape with picture content.
        /// </summary>
        Shape WithPicture(ILocation innerTransform, IShapeContext spContext, PictureEx picture, GeometryType geometry);

        /// <summary>
        /// Builds a AutoShape.
        /// </summary>
        Shape WithAutoShape(ILocation innerTransform, IShapeContext spContext, GeometryType geometry);

        /// <summary>
        /// Builds a shape with table content.
        /// </summary>
        Shape WithTable(ILocation innerTransform, IShapeContext spContext, TableEx table);

        /// <summary>
        /// Builds a shape with OLE object content.
        /// </summary>
        Shape WithChart(ILocation innerTransform, IShapeContext spContext, ChartEx chart);

        /// <summary>
        /// Builds a group shape which has grouped shape items.
        /// </summary>
        Shape WithGroup(ILocation innerTransform, IShapeContext spContext, IList<Shape> groupedShapes);
    }
}
