using System.Collections.Generic;
using DocumentFormat.OpenXml;
using ShapeCrawler.Extensions;
using ShapeCrawler.Factories;
using ShapeCrawler.Settings;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a group shape on a Slide.
    /// </summary>
    internal class SlideGroupShape : SlideShape, IGroupShape
    {
        private readonly ILocation _innerTransform;

        internal ShapeContext Context;

        #region Constructors

        internal SlideGroupShape(
            ILocation innerTransform,
            ShapeContext spContext,
            List<IShape> groupedShapes,
            OpenXmlCompositeElement pShapeTreeChild,
            SlideSc slide) : base(slide, pShapeTreeChild)
        {
            _innerTransform = innerTransform;
            Context = spContext;
            Shapes = groupedShapes;
        }

        #endregion Constructors

        #region Public Properties

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

        public IReadOnlyCollection<IShape> Shapes { get; }

        #endregion Properties
    }
}