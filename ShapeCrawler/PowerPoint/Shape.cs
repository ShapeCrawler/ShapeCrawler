using System;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Extensions;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Statics;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a base class for shapes located on Slide, Slide Layout or Master Slide.
    /// </summary>
    internal abstract class Shape
    {
        /// <summary>
        ///     Initializes a new instance of the <see cref="Shape"/> class.
        /// </summary>
        protected Shape(OpenXmlCompositeElement pShapeTreeChild)
        {
            this.PShapeTreeChild = pShapeTreeChild;
        }

        #region Public Properties

        public int Id => (int)this.PShapeTreeChild.GetNonVisualDrawingProperties().Id.Value;

        public string Name => this.PShapeTreeChild.GetNonVisualDrawingProperties().Name;

        public bool Hidden => this.DefineHidden();

        public string CustomData
        {
            get => this.GetCustomData();
            set => this.SetCustomData(value);
        }

        /// <summary>
        ///     Gets placeholder. Returns <c>NULL</c> if the shape is not a placeholder.
        /// </summary>
        public abstract IPlaceholder Placeholder { get; }

        /// <summary>
        ///     Gets parent presentation.
        /// </summary>
        public abstract SCPresentation Presentation { get; }

        /// <summary>
        ///     Gets parent Slide Master.
        /// </summary>
        public abstract SCSlideMaster SlideMaster { get; }

        public virtual GeometryType GeometryType => this.GetGeometryType();

        /// <summary>
        ///     Gets y-coordinate of the upper-left corner of the shape.
        /// </summary>
        public long Y
        {
            get => this.GetY();
            set => this.SetY(value);
        }

        /// <summary>
        ///     Gets height of the shape.
        /// </summary>
        public long Height
        {
            get => this.GetHeight();
            set => this.SetHeight(value);
        }

        /// <summary>
        ///     Gets width of the shape.
        /// </summary>
        public long Width
        {
            get => GetWidth();
            set => SetWidth(value);
        }

        /// <summary>
        ///     Gets x-coordinate of the upper-left corner of the shape.
        /// </summary>
        public long X
        {
            get => GetX();
            set => SetX(value);
        }

        #endregion Public Properties

        internal OpenXmlCompositeElement PShapeTreeChild { get; }

        internal abstract ThemePart ThemePart { get; }

        protected void SetCustomData(string value)
        {
            string customDataElement =
                $@"<{ConstantStrings.CustomDataElementName}>{value}</{ConstantStrings.CustomDataElementName}>";
            this.PShapeTreeChild.InnerXml += customDataElement;
        }

        private bool DefineHidden()
        {
            bool? parsedHiddenValue = PShapeTreeChild.GetNonVisualDrawingProperties().Hidden?.Value;
            return parsedHiddenValue != null && parsedHiddenValue == true;
        }

        private string GetCustomData()
        {
            var pattern = @$"<{ConstantStrings.CustomDataElementName}>(.*)<\/{ConstantStrings.CustomDataElementName}>";
            var regex = new Regex(pattern);
            var elementText = regex.Match(PShapeTreeChild.InnerXml).Groups[1];
            if (elementText.Value.Length == 0)
            {
                return null;
            }

            return elementText.Value;
        }

        private void SetX(long value)
        {
            A.Offset aOffset = PShapeTreeChild.Descendants<A.Offset>().FirstOrDefault();
            if (aOffset == null)
            {
                Shape placeholderShape = ((Placeholder) Placeholder).ReferencedShape;
                placeholderShape.X = value;
            }
            else
            {
                aOffset.X = value;
            }
        }

        private long GetX()
        {
            A.Offset aOffset = PShapeTreeChild.Descendants<A.Offset>().FirstOrDefault();
            if (aOffset == null)
            {
                return ((Placeholder) Placeholder).ReferencedShape.X;
            }

            return aOffset.X;
        }


        private void SetY(long value)
        {
            throw new NotImplementedException();
        }

        private long GetY()
        {
            A.Offset aOffset = PShapeTreeChild.Descendants<A.Offset>().FirstOrDefault();
            if (aOffset == null)
            {
                return ((Placeholder) Placeholder).ReferencedShape.Y;
            }

            return aOffset.Y;
        }

        private void SetWidth(long value)
        {
            throw new NotImplementedException();
        }

        private long GetWidth()
        {
            A.Extents aExtents = PShapeTreeChild.Descendants<A.Extents>().FirstOrDefault();
            if (aExtents == null)
            {
                return ((Placeholder) Placeholder).ReferencedShape.Width;
            }

            return aExtents.Cx;
        }

        private void SetHeight(long value)
        {
            throw new NotImplementedException();
        }

        private long GetHeight()
        {
            A.Extents aExtents = PShapeTreeChild.Descendants<A.Extents>().FirstOrDefault();
            if (aExtents == null)
            {
                return ((Placeholder) Placeholder).ReferencedShape.Height;
            }

            return aExtents.Cy;
        }

        private GeometryType GetGeometryType()
        {
            P.ShapeProperties spPr = PShapeTreeChild.Descendants<P.ShapeProperties>().First(); // TODO: optimize
            A.Transform2D transform2D = spPr.Transform2D;
            if (transform2D != null)
            {
                A.PresetGeometry aPresetGeometry = spPr.GetFirstChild<A.PresetGeometry>();

                // Placeholder can have transform on the slide, without having geometry
                if (aPresetGeometry == null)
                {
                    if (spPr.OfType<A.CustomGeometry>().Any())
                    {
                        return GeometryType.Custom;
                    }
                }
                else
                {
                    var name = aPresetGeometry.Preset.Value.ToString();
                    Enum.TryParse(name, true, out GeometryType geometryType);
                    return geometryType;
                }
            }

            Placeholder placeholder = (Placeholder) Placeholder;
            if (placeholder?.ReferencedShape != null)
            {
                return placeholder.ReferencedShape.GeometryType;
            }

            return GeometryType.Rectangle; // return default
        }

        internal void ThrowIfRemoved()
        {
            if (this.IsRemoved)
            {
                throw new ElementIsRemovedException("Shape was removed.");
            }
            else
            {
                
            }
        }

        public bool IsRemoved { get; set; }
    }
}