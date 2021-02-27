using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using ShapeCrawler.Drawing;
using ShapeCrawler.Extensions;
using ShapeCrawler.Factories;
using ShapeCrawler.Settings;
using ShapeCrawler.Statics;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration

namespace ShapeCrawler
{
    /// <inheritdoc cref="IPicture" />
    public class PictureSc : Shape, IPicture
    {
        #region Constructors

        internal PictureSc(
            SlideSc slide,
            string blipRelateId,
            ILocation innerTransform,
            ShapeContext spContext,
            GeometryType geometryType,
            P.Picture pPicture) : base(pPicture)
        {
            Slide = slide;
            Image = new ImageSc(Slide.SlidePart, blipRelateId);
            _innerTransform = innerTransform;
            Context = spContext;
            GeometryType = geometryType;
        }

        #endregion Constructors

        #region Fields

        private bool? _hidden;
        private int _id;
        private string _name;
        private readonly ILocation _innerTransform;

        internal OpenXmlCompositeElement ShapeTreeSource { get; }
        internal ShapeContext Context { get; }
        internal SlideSc Slide { get; }

        #endregion Fields

        #region Public Properties

        public ImageSc Image { get; }

        public long X
        {
            get => _innerTransform.X;
            set => _innerTransform.SetX(value);
        }

        public long Y
        {
            get => _innerTransform.Y;
            set => _innerTransform.SetY(value);
        }

        public long Width
        {
            get => _innerTransform.Width;
            set => _innerTransform.SetWidth(value);
        }

        public long Height
        {
            get => _innerTransform.Height;
            set => _innerTransform.SetHeight(value);
        }

        public int Id
        {
            get
            {
                InitIdHiddenName();
                return _id;
            }
        }

        public string Name
        {
            get
            {
                InitIdHiddenName();
                return _name;
            }
        }

        public bool Hidden
        {
            get
            {
                InitIdHiddenName();
                return (bool) _hidden;
            }
        }

        public GeometryType GeometryType { get; }

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