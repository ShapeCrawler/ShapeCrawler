﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Placeholders;
using ShapeCrawler.SlideMaster;

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a shape on a Slide Layout.
    /// </summary>
    internal abstract class LayoutShape : Shape
    {
        protected LayoutShape(SlideLayoutSc slideLayout, OpenXmlCompositeElement pShapeTreeChild) : base(
            pShapeTreeChild)
        {
            SlideLayout = slideLayout;
        }

        internal SlideLayoutSc SlideLayout { get; }

        public override IPlaceholder Placeholder => LayoutPlaceholder.Create(PShapeTreeChild, this);

        internal override ThemePart ThemePart => SlideLayout.SlideLayoutPart.SlideMasterPart.ThemePart;

        public override SCPresentation Presentation => SlideLayout.SlideMaster.Presentation;

        public override SlideMasterSc SlideMaster => SlideLayout.SlideMaster;
    }
}