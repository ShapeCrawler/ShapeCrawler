using DocumentFormat.OpenXml;
using ShapeCrawler.Enums;
using ShapeCrawler.Models;
using System;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler
{
    /// <summary>
    /// Represents an auto shape on a Slide Master.
    /// </summary>
    public class MasterAutoShape : MasterShape
    {
        public MasterAutoShape(OpenXmlCompositeElement pShape) : base(pShape)
        {
        }
    }

    public class MasterShape : BaseShape
    {
        public MasterShape(OpenXmlCompositeElement compositeElement): base(compositeElement)
        {
        }

        public override long X => _compositeElement.GetFirstChild<P.ShapeProperties>().Transform2D.Offset.X;

        public override long Y => _compositeElement.GetFirstChild<P.ShapeProperties>().Transform2D.Offset.Y;

        public override long Width => _compositeElement.GetFirstChild<P.ShapeProperties>().Transform2D.Extents.Cx;

        public override long Height => _compositeElement.GetFirstChild<P.ShapeProperties>().Transform2D.Extents.Cy;
        
        public override GeometryType GeometryType { get; }

        public PlaceholderType? PlaceholderType => GetPlaceholderType();

        private PlaceholderType? GetPlaceholderType()
        {
            P.PlaceholderShape placeholderShape = _compositeElement.GetFirstChild<P.NonVisualShapeProperties>().
                ApplicationNonVisualDrawingProperties.PlaceholderShape;
            if (placeholderShape.Type == null)
            {
                return null;
            }

            // Convert outer sdk placeholder type into library placeholder type
            if (placeholderShape.Type == P.PlaceholderValues.Title || placeholderShape.Type == P.PlaceholderValues.CenteredTitle)
            {
                return Enums.PlaceholderType.Title;
            }
            return (PlaceholderType)Enum.Parse(typeof(PlaceholderType), placeholderShape.Type.Value.ToString());
        }
    }
}