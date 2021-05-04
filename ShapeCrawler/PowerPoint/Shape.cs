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
    ///     Represents a shape.
    /// </summary>
    internal abstract class Shape
    {
        protected Shape(OpenXmlCompositeElement pShapeTreeChild)
        {
            this.PShapeTreeChild = pShapeTreeChild;
        }

        #region Public Properties

        /// <summary>
        ///     Gets shape identifier.
        /// </summary>
        public int Id => (int)this.PShapeTreeChild.GetNonVisualDrawingProperties().Id.Value;

        /// <summary>
        ///     Gets shape name.
        /// </summary>
        public string Name => this.PShapeTreeChild.GetNonVisualDrawingProperties().Name;

        /// <summary>
        ///     Gets a value indicating whether shape is hidden.
        /// </summary>
        public bool Hidden => this.DefineHidden(); // TODO: the Shape is inherited by LayoutShape, hence do we need this property?

        /// <summary>
        ///     Gets or sets custom data.
        /// </summary>
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
        public abstract SCPresentation ParentPresentation { get; } // TODO: move it on Slide level

        /// <summary>
        ///     Gets parent Slide Master.
        /// </summary>
        public abstract SCSlideMaster ParentSlideMaster { get; }

        /// <summary>
        ///     Gets geometry form type.
        /// </summary>
        public virtual GeometryType GeometryType => this.GetGeometryType();

        /// <summary>
        ///     Gets or sets y-coordinate of the upper-left corner of the shape.
        /// </summary>
        public long Y
        {
            get => this.GetY();
            set => this.SetY(value);
        }

        /// <summary>
        ///     Gets or sets height of the shape.
        /// </summary>
        public long Height
        {
            get => this.GetHeight();
            set => this.SetHeight(value);
        }

        /// <summary>
        ///     Gets or sets width of the shape.
        /// </summary>
        public long Width
        {
            get => this.GetWidth();
            set => this.SetWidth(value);
        }

        /// <summary>
        ///     Gets or sets x-coordinate of the upper-left corner of the shape.
        /// </summary>
        public long X
        {
            get => this.GetXCoordinate();
            set => this.SetXCoordinate(value);
        }

        #endregion Public Properties

        internal OpenXmlCompositeElement PShapeTreeChild { get; }

        internal bool IsRemoved { get; set; }

        protected void SetCustomData(string value)
        {
            string customDataElement =
                $@"<{ConstantStrings.CustomDataElementName}>{value}</{ConstantStrings.CustomDataElementName}>";
            this.PShapeTreeChild.InnerXml += customDataElement;
        }

        private bool DefineHidden()
        {
            bool? parsedHiddenValue = this.PShapeTreeChild.GetNonVisualDrawingProperties().Hidden?.Value;
            return parsedHiddenValue is true;
        }

        private string GetCustomData()
        {
            var pattern = @$"<{ConstantStrings.CustomDataElementName}>(.*)<\/{ConstantStrings.CustomDataElementName}>";
            var regex = new Regex(pattern);
            var elementText = regex.Match(this.PShapeTreeChild.InnerXml).Groups[1];
            if (elementText.Value.Length == 0)
            {
                return null;
            }

            return elementText.Value;
        }

        private void SetXCoordinate(long value)
        {
            A.Offset aOffset = this.PShapeTreeChild.Descendants<A.Offset>().FirstOrDefault();
            if (aOffset == null)
            {
                Shape placeholderShape = ((Placeholder) this.Placeholder).ReferencedShape;
                placeholderShape.X = value;
            }
            else
            {
                aOffset.X = value;
            }
        }

        private long GetXCoordinate()
        {
            A.Offset aOffset = this.PShapeTreeChild.Descendants<A.Offset>().FirstOrDefault();
            if (aOffset == null)
            {
                return ((Placeholder) this.Placeholder).ReferencedShape.X;
            }

            return aOffset.X;
        }


        private void SetY(long value)
        {
            throw new NotImplementedException();
        }

        private long GetY()
        {
            A.Offset aOffset = this.PShapeTreeChild.Descendants<A.Offset>().FirstOrDefault();
            if (aOffset == null)
            {
                return ((Placeholder)this.Placeholder).ReferencedShape.Y;
            }

            return aOffset.Y;
        }

        private void SetWidth(long value)
        {
            throw new NotImplementedException();
        }

        private long GetWidth()
        {
            A.Extents aExtents = this.PShapeTreeChild.Descendants<A.Extents>().FirstOrDefault();
            if (aExtents == null)
            {
                return ((Placeholder)this.Placeholder).ReferencedShape.Width;
            }

            return aExtents.Cx;
        }

        private void SetHeight(long value)
        {
            throw new NotImplementedException();
        }

        private long GetHeight()
        {
            A.Extents aExtents = this.PShapeTreeChild.Descendants<A.Extents>().FirstOrDefault();
            if (aExtents == null)
            {
                return ((Placeholder)this.Placeholder).ReferencedShape.Height;
            }

            return aExtents.Cy;
        }

        private GeometryType GetGeometryType()
        {
            P.ShapeProperties spPr = this.PShapeTreeChild.Descendants<P.ShapeProperties>().First(); // TODO: optimize
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

            Placeholder placeholder = (Placeholder)this.Placeholder;
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
    }
}