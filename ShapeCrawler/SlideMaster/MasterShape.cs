using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using ShapeCrawler.Extensions;
using ShapeCrawler.Models;
using ShapeCrawler.Models.Experiment;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.SlideMaster
{
    /// <summary>
    /// Represents a shape on Slide Master.
    /// </summary>
    public class MasterShape : BaseShape
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="MasterShape"/> class.
        /// </summary>
        /// <param name="slideMaster"></param>
        /// <param name="shapeTreeSource">Element of the shape tree.</param>
        public MasterShape(SlideMasterSc slideMaster, OpenXmlCompositeElement shapeTreeSource) 
            : base(slideMaster, shapeTreeSource)
        {
        }

        public override long X => CompositeElement.GetFirstChild<P.ShapeProperties>().Transform2D.Offset.X;

        public override long Y => CompositeElement.GetFirstChild<P.ShapeProperties>().Transform2D.Offset.Y;

        public override long Width => CompositeElement.GetFirstChild<P.ShapeProperties>().Transform2D.Extents.Cx;

        public override long Height => CompositeElement.GetFirstChild<P.ShapeProperties>().Transform2D.Extents.Cy;

        public override GeometryType GeometryType => GetGeometryType();

        /// <summary>
        /// Gets placeholder type. Returns null if the master shape is not a placeholder.
        /// </summary>
        public PlaceholderType? PlaceholderType => GetPlaceholderType();

        #region Private Methods

        private GeometryType GetGeometryType()
        {
            // Get the shape geometry type in SDK format
            PresetGeometry presetGeometry = CompositeElement.GetFirstChild<P.ShapeProperties>().
                                                                GetFirstChild<PresetGeometry>();
            
            if (presetGeometry == null)
            {
                return GeometryType.Custom;
            }

#if NETSTANDARD2_0
            Enum.TryParse(presetGeometry.Preset.Value.ToString(), true, out GeometryType geometryType);
#else
            // Get SDK format into internal type
            GeometryType geometryType = Enum.Parse<GeometryType>(presetGeometry.Preset.Value.ToString());
#endif

            return geometryType;
        }

        private PlaceholderType? GetPlaceholderType()
        {
            P.PlaceholderShape placeholderShape = CompositeElement.GetApplicationNonVisualDrawingProperties().PlaceholderShape;
            if (placeholderShape == null)
            {
                return null;
            }

            // Convert outer sdk placeholder type into library placeholder type
            if (placeholderShape.Type == P.PlaceholderValues.Title ||
                placeholderShape.Type == P.PlaceholderValues.CenteredTitle)
            {
                return ShapeCrawler.PlaceholderType.Title;
            }

            return (PlaceholderType) Enum.Parse(typeof(PlaceholderType), placeholderShape.Type.Value.ToString());
        }

        #endregion Private Methods
    }
}