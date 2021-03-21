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
            ShapeContext spContext,
            SlideSc slide) : base(slide, shapeTreeChild)
        {
            ShapeTreeChild = shapeTreeChild;
            Context = spContext;
        }

        #endregion Constructors

        #region Fields

        internal ShapeContext Context;
        private int _id;

        internal OpenXmlCompositeElement ShapeTreeChild { get; }

        #endregion Fields

        #region Public Properties

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

        #endregion
    }
}