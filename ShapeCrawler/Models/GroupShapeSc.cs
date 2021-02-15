using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using ShapeCrawler.Extensions;
using ShapeCrawler.Factories.Drawing;
using ShapeCrawler.Models;
using ShapeCrawler.Models.SlideComponents;
using ShapeCrawler.Models.Styles;
using ShapeCrawler.Models.Transforms;
using ShapeCrawler.Settings;
using ShapeCrawler.Statics;
using ShapeCrawler.Texts;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
// ReSharper disable CheckNamespace

// ReSharper disable PossibleMultipleEnumeration

namespace ShapeCrawler
{

    /// <summary>
    /// Represents a shape on a slide.
    /// </summary>
    public class GroupShapeSc : IGroupShape
    {
        #region Fields

        internal ShapeContext Context;
        private readonly Lazy<TextBoxSc> _text;
        private readonly Lazy<ShapeFill> _shapeFill;
        private readonly IImageExFactory _imageFactory = new ImageExFactory();
        private bool? _hidden;
        private int _id;
        private string _name;
        private readonly ILocation _innerTransform;

        #endregion Fields

        internal OpenXmlCompositeElement ShapeTreeSource { get; set; }

        #region Public Properties

        /// <summary>
        /// Returns the x-coordinate of the upper-left corner of the shape.
        /// </summary>
        public long X
        {
            get => _innerTransform.X;
            set => _innerTransform.SetX(value);
        }

        /// <summary>
        /// Returns the y-coordinate of the upper-left corner of the shape.
        /// </summary>
        public long Y
        {
            get => _innerTransform.Y;
            set => _innerTransform.SetY(value);
        }

        /// <summary>
        /// Returns the width of the shape.
        /// </summary>
        public long Width
        {
            get => _innerTransform.Width;
            set => _innerTransform.SetWidth(value);
        }

        /// <summary>
        /// Returns the height of the shape.
        /// </summary>
        public long Height
        {
            get => _innerTransform.Height;
            set => _innerTransform.SetHeight(value);
        }

        /// <summary>
        /// Returns shape main content type.
        /// </summary>
        public ShapeContentType ContentType { get; }

        /// <summary>
        /// Returns an element identifier.
        /// </summary>
        public int Id
        {
            get
            {
                InitIdHiddenName();
                return _id;
            }
        }

        /// <summary>
        /// Gets an element name.
        /// </summary>
        public string Name
        {
            get
            {
                InitIdHiddenName();
                return _name;
            }
        }

        /// <summary>
        /// Determines whether the shape is hidden.
        /// </summary>
        public bool Hidden
        {
            get
            {
                InitIdHiddenName();
                return (bool)_hidden;
            }
        }

        /// <summary>
        /// Returns text frame.
        /// </summary>
        /// <remarks>Lazy load.</remarks>
        public TextBoxSc TextBox => _text.Value;

        /// <summary>
        /// Determines whether the shape is placeholder.
        /// </summary>
        public bool IsPlaceholder => Placeholder != null;

        public Placeholder Placeholder
        {
            get
            {
                if (Context.CompositeElement.IsPlaceholder())
                {
                    return new Placeholder();
                }

                return null;
            }
        }

        /// <summary>
        /// Returns placeholder type. Returns null if shape is not a placeholder.
        /// </summary>
        public PlaceholderType? PlaceholderType
        {
            get
            {
                if (!IsPlaceholder)
                {
                    return null;
                }
                
                return Context.PlaceholderService.GetPlaceholderType(Context.CompositeElement);
            }
        }

        /// <summary>
        /// Returns the Shape Fill. Returns <c>NULL</c> if shape is not filled.
        /// </summary>
        public ShapeFill Fill => _shapeFill.Value;

        /// <summary>
        /// Determines whether the shape is grouped.
        /// </summary>
        public bool IsGrouped => _innerTransform is NonPlaceholderGroupedTransform;

        public GeometryType GeometryType => GeometryType.Rectangle;

        public string CustomData
        {
            get => GetCustomData();
            set => SetCustomData(value);
        }

        /// <summary>
        /// Gets parent slide.
        /// </summary>
        public SlideSc Slide { get; internal set; }

        public IReadOnlyCollection<IShape> Shapes { get; }

        #endregion Properties

        #region Constructors

        internal GroupShapeSc(
            ILocation innerTransform,
            ShapeContext spContext,
            List<IShape> groupedShapes,
            OpenXmlCompositeElement shapeTreeSource)
        {
            _innerTransform = innerTransform;
            Context = spContext;
            Shapes = groupedShapes;
            _shapeFill = new Lazy<ShapeFill>(TryGetFill);
            ShapeTreeSource = shapeTreeSource;
        }

        #endregion Constructors

        #region Private Methods

        private void SetCustomData(string value)
        {
            var customDataElement = $@"<{ConstantStrings.CustomDataElementName}>{value}</{ConstantStrings.CustomDataElementName}>";
            Context.CompositeElement.InnerXml += customDataElement;
        }

        private string GetCustomData()
        {
            var pattern = @$"<{ConstantStrings.CustomDataElementName}>(.*)<\/{ConstantStrings.CustomDataElementName}>";
            var regex = new Regex(pattern);
            var elementText = regex.Match(Context.CompositeElement.InnerXml).Groups[1];
            if (elementText.Value.Length == 0)
            {
                return null;
            }

            return elementText.Value;
        }

        private ShapeFill TryGetFill()
        {
            if (ContentType != ShapeContentType.AutoShape)
            {
                return null;
            }
            ImageSc image = _imageFactory.TryFromSdkShape(Context.SlidePart, Context.CompositeElement);
            if (image != null)
            {
                return new ShapeFill(image);
            }

            var xmlShape = (P.Shape) Context.CompositeElement;
            var rgbColorModelHex = xmlShape.ShapeProperties.GetFirstChild<A.SolidFill>()?.RgbColorModelHex;
            return rgbColorModelHex != null ? ShapeFill.FromXmlSolidFill(rgbColorModelHex) : null;
        }

        private void InitIdHiddenName()
        {
            if (_id != 0)
            {
                return;
            }
            var (id, hidden, name) = Context.CompositeElement.GetNvPrValues();
            _id = id;
            _hidden = hidden;
            _name = name;
        }

        #endregion

    }
}