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
    /// <inheritdoc cref="IGroupShape" />
    public class GroupShapeSc : Shape, IGroupShape
    {
        #region Constructors

        internal GroupShapeSc(
            ILocation innerTransform,
            ShapeContext spContext,
            List<IShape> groupedShapes,
            OpenXmlCompositeElement pShapeTreeChild) : base(pShapeTreeChild)
        {
            _innerTransform = innerTransform;
            Context = spContext;
            Shapes = groupedShapes;
        }

        #endregion Constructors

        #region Private Methods

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

        #endregion Private Methods

        #region Fields

        private bool? _hidden;
        private int _id;
        private string _name;
        private readonly ILocation _innerTransform;

        internal ShapeContext Context;
        internal SlideSc Slide { get; }

        #endregion Fields

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

        public GeometryType GeometryType => GeometryType.Rectangle;

        public IReadOnlyCollection<IShape> Shapes { get; }

        #endregion Properties
    }
}