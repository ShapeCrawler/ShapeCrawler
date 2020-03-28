using System.Collections.Generic;
using SlideDotNet.Models.Settings;
using SlideDotNet.Models.SlideComponents;
using SlideDotNet.Models.SlideComponents.Chart;
using SlideDotNet.Models.TableComponents;
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
        ShapeEx WithOle(IInnerTransform innerTransform, IShapeContext spContext, OleObject ole);

        /// <summary>
        /// Builds a shape with picture content.
        /// </summary>
        ShapeEx WithPicture(IInnerTransform innerTransform, IShapeContext spContext, PictureEx picture);

        /// <summary>
        /// Builds a AutoShape.
        /// </summary>
        ShapeEx WithAutoShape(IInnerTransform innerTransform, IShapeContext spContext);

        /// <summary>
        /// Builds a shape with table content.
        /// </summary>
        ShapeEx WithTable(IInnerTransform innerTransform, IShapeContext spContext, TableEx table);

        /// <summary>
        /// Builds a shape with OLE object content.
        /// </summary>
        ShapeEx WithChart(IInnerTransform innerTransform, IShapeContext spContext, ChartEx chart);

        /// <summary>
        /// Builds a group shape which has grouped shape items.
        /// </summary>
        ShapeEx WithGroup(IInnerTransform innerTransform, IShapeContext spContext, IList<ShapeEx> groupedShapes);
    }
}
