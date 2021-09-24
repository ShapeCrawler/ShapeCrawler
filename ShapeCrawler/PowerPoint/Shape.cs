using System;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Extensions;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using ShapeCrawler.Statics;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a shape.
    /// </summary>
    internal abstract class Shape : IRemovable
    {
        protected Shape(OpenXmlCompositeElement sdkPShapeTreeChild, IBaseSlide parentBaseSlide)
        {
            this.SdkPShapeTreeChild = sdkPShapeTreeChild;
            this.ParentBaseSlide = parentBaseSlide;
        }

        #region Public Properties

        /// <summary>
        ///     Gets shape identifier.
        /// </summary>
        public int Id => (int)this.SdkPShapeTreeChild.GetNonVisualDrawingProperties().Id.Value;

        /// <summary>
        ///     Gets shape name.
        /// </summary>
        public string Name => this.SdkPShapeTreeChild.GetNonVisualDrawingProperties().Name;

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
        ///     Gets parent Slide Master.
        /// </summary>
        public abstract SCSlideMaster ParentSlideMaster { get; }

        /// <summary>
        ///     Gets geometry form type.
        /// </summary>
        public virtual GeometryType GeometryType => this.GetGeometryType();

        /// <summary>
        ///     Gets or sets x-coordinate of the upper-left corner of the shape.
        /// </summary>
        public int X
        {
            get => this.GetXCoordinate();
            set => this.SetXCoordinate(value);
        }

        /// <summary>
        ///     Gets or sets y-coordinate of the upper-left corner of the shape.
        /// </summary>
        public int Y
        {
            get => this.GetYCoordinate();
            set => this.SetYCoordinate(value);
        }

        /// <summary>
        ///     Gets or sets height of the shape.
        /// </summary>
        public int Height
        {
            get => this.GetHeight();
            set => this.SetHeight(value);
        }

        /// <summary>
        ///     Gets or sets width of the shape.
        /// </summary>
        public int Width
        {
            get => this.GetWidth();
            set => this.SetWidth(value);
        }

        public IBaseSlide ParentBaseSlide { get; }

        #endregion Public Properties

        internal OpenXmlCompositeElement SdkPShapeTreeChild { get; }

        bool IRemovable.IsRemoved { get; set; }

        public void ThrowIfRemoved()
        {
            if (((IRemovable)this).IsRemoved)
            {
                throw new ElementIsRemovedException("Shape was removed.");
            }

            this.ParentBaseSlide.ThrowIfRemoved();
        }

        private void SetCustomData(string value)
        {
            string customDataElement =
                $@"<{ConstantStrings.CustomDataElementName}>{value}</{ConstantStrings.CustomDataElementName}>";
            this.SdkPShapeTreeChild.InnerXml += customDataElement;
        }

        private string GetCustomData()
        {
            var pattern = @$"<{ConstantStrings.CustomDataElementName}>(.*)<\/{ConstantStrings.CustomDataElementName}>";
            var regex = new Regex(pattern);
            var elementText = regex.Match(this.SdkPShapeTreeChild.InnerXml).Groups[1];
            if (elementText.Value.Length == 0)
            {
                return null;
            }

            return elementText.Value;
        }

        private bool DefineHidden()
        {
            bool? parsedHiddenValue = this.SdkPShapeTreeChild.GetNonVisualDrawingProperties().Hidden?.Value;
            return parsedHiddenValue is true;
        }

        private void SetXCoordinate(int value)
        {
            A.Offset aOffset = this.SdkPShapeTreeChild.Descendants<A.Offset>().FirstOrDefault();
            if (aOffset == null)
            {
                Shape placeholderShape = ((Placeholder) this.Placeholder).ReferencedShape;
                placeholderShape.X = value;
            }
            else
            {
                Bitmap bm = new(100, 100);
                aOffset.X = (long)(value * 914400 / bm.HorizontalResolution);
            }
        }

        private int GetXCoordinate()
        {
            A.Offset aOffset = this.SdkPShapeTreeChild.Descendants<A.Offset>().FirstOrDefault();
            if (aOffset == null)
            {
                return ((Placeholder) this.Placeholder).ReferencedShape.X;
            }

            return PixelConverter.HorizontalEmuToPixel(aOffset.X);
        }

        private void SetYCoordinate(long value)
        {
            A.Offset aOffset = this.SdkPShapeTreeChild.Descendants<A.Offset>().FirstOrDefault();
            if (this.Placeholder is not null)
            {
                throw new PlaceholderCannotBeChangedException();
            }

            Bitmap bm = new(100, 100);
            aOffset.Y = (long)(value * 914400 / bm.VerticalResolution);
        }

        private int GetYCoordinate()
        {
            A.Offset aOffset = this.SdkPShapeTreeChild.Descendants<A.Offset>().FirstOrDefault();
            if (aOffset == null)
            {
                return ((Placeholder)this.Placeholder).ReferencedShape.Y;
            }

            return PixelConverter.VerticalEmuToPixel(aOffset.Y);
        }

        private int GetWidth()
        {
            A.Extents aExtents = this.SdkPShapeTreeChild.Descendants<A.Extents>().FirstOrDefault();
            if (aExtents == null)
            {
                return ((Placeholder)this.Placeholder).ReferencedShape.Width;
            }

            Bitmap bm = new(100, 100);
            return (int)(aExtents.Cx * bm.HorizontalResolution / 914400);
        }

        private void SetWidth(int pixels)
        {
            A.Extents aExtents = this.SdkPShapeTreeChild.Descendants<A.Extents>().FirstOrDefault();
            if (aExtents == null)
            {
                throw new PlaceholderCannotBeChangedException();
            }

            Bitmap bm = new (100, 100);
            aExtents.Cx = (long)(pixels * 914400 / bm.HorizontalResolution);
        }

        private int GetHeight()
        {
            A.Extents aExtents = this.SdkPShapeTreeChild.Descendants<A.Extents>().FirstOrDefault();
            if (aExtents == null)
            {
                return ((Placeholder)this.Placeholder).ReferencedShape.Height;
            }

            Bitmap bm = new(100, 100);
            return (int)(aExtents.Cy * bm.VerticalResolution / 914400);
        }

        private void SetHeight(int pixels)
        {
            A.Extents aExtents = this.SdkPShapeTreeChild.Descendants<A.Extents>().FirstOrDefault();
            if (aExtents == null)
            {
                throw new PlaceholderCannotBeChangedException();
            }

            Bitmap bm = new (100, 100);
            aExtents.Cy = (long)(pixels * 914400 / bm.VerticalResolution);
        }

        private GeometryType GetGeometryType()
        {
            P.ShapeProperties spPr = this.SdkPShapeTreeChild.Descendants<P.ShapeProperties>().First(); // TODO: optimize
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
    }
}