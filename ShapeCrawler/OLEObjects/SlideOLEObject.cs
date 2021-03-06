using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using ShapeCrawler.Extensions;
using ShapeCrawler.Factories;
using ShapeCrawler.OLEObjects;
using ShapeCrawler.Settings;
using ShapeCrawler.Statics;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a shape on a slide.
    /// </summary>
    public class SlideOLEObject : SlideShape, IOLEObject
    {
        #region Constructors

        internal SlideOLEObject(
            OpenXmlCompositeElement shapeTreeChild,
            ILocation innerTransform,
            ShapeContext spContext,
            SlideSc slide) : base(slide, shapeTreeChild)
        {
            ShapeTreeChild = shapeTreeChild;
            _innerTransform = innerTransform;
            Context = spContext;
        }

        #endregion Constructors

        #region Fields

        internal ShapeContext Context;
        private bool? _hidden;
        private int _id;
        private string _name;
        private readonly ILocation _innerTransform;

        internal OpenXmlCompositeElement ShapeTreeChild { get; }
        internal SlideSc Slide { get; }

        #endregion Fields

        #region Public Properties

        /// <summary>
        ///     Returns the x-coordinate of the upper-left corner of the shape.
        /// </summary>
        public long X
        {
            get => _innerTransform.X;
            set => _innerTransform.SetX(value);
        }

        /// <summary>
        ///     Returns the y-coordinate of the upper-left corner of the shape.
        /// </summary>
        public long Y
        {
            get => _innerTransform.Y;
            set => _innerTransform.SetY(value);
        }

        /// <summary>
        ///     Returns the width of the shape.
        /// </summary>
        public long Width
        {
            get => _innerTransform.Width;
            set => _innerTransform.SetWidth(value);
        }

        /// <summary>
        ///     Returns the height of the shape.
        /// </summary>
        public long Height
        {
            get => _innerTransform.Height;
            set => _innerTransform.SetHeight(value);
        }

        /// <summary>
        ///     Returns an element identifier.
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
        ///     Gets an element name.
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
        ///     Determines whether the shape is hidden.
        /// </summary>
        public bool Hidden
        {
            get
            {
                InitIdHiddenName();
                return (bool) _hidden;
            }
        }

        public GeometryType GeometryType => GeometryType.Rectangle;

        public string CustomData
        {
            get => GetCustomData();
            set => SetCustomData(value);
        }

        #endregion Properties

        #region Private Methods

        private void SetCustomData(string value)
        {
            var customDataElement =
                $@"<{ConstantStrings.CustomDataElementName}>{value}</{ConstantStrings.CustomDataElementName}>";
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