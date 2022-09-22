using System;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace
namespace ShapeCrawler
{
    /// <inheritdoc cref="IPortion"/>
    internal class SCPortion : IPortion
    {
        private readonly ResettableLazy<SCFont> font;

        /// <summary>
        ///     Initializes a new instance of the <see cref="SCPortion"/> class.
        /// </summary>
        internal SCPortion(A.Text sdkaText, SCParagraph paragraph)
        {
            this.SDKAText = sdkaText;
            this.ParentParagraph = paragraph;
            this.font = new ResettableLazy<SCFont>(() => new SCFont(this.SDKAText, this));
        }

        #region Public Properties

        /// <inheritdoc/>
        public string Text
        {
            get => this.GetText();

            set
            {
                this.ThrowIfRemoved();
                this.SetText(value);
            }
        }

        /// <inheritdoc/>
        public IFont Font => this.font.Value;

        public string Hyperlink
        {
            get => this.GetHyperlink();
            set => this.SetHyperlink(value);
        }
        
        public A.Text SDKAText { get; }

        #endregion Public Properties

        internal bool IsRemoved { get; set; }

        internal SCParagraph ParentParagraph { get; }

        private void ThrowIfRemoved()
        {
            if (this.IsRemoved)
            {
                throw new ElementIsRemovedException("Paragraph portion was removed.");
            }

            this.ParentParagraph.ThrowIfRemoved();
        }

        private string GetText()
        {
            string portionText = this.SDKAText.Text;
            if (this.SDKAText.Parent.NextSibling<A.Break>() != null)
            {
                portionText += Environment.NewLine;
            }

            return portionText;
        }

        private void SetText(string text)
        {
            this.SDKAText.Text = text;
        }

        private string? GetHyperlink()
        {
            var runProperties = this.SDKAText.PreviousSibling<A.RunProperties>();
            if (runProperties == null)
            {
                return null;
            }

            var hyperlink = runProperties.GetFirstChild<A.HyperlinkOnClick>();
            if (hyperlink == null)
            {
                return null;
            }

            var slideAutoShape = (SlideAutoShape)this.ParentParagraph.ParentTextBox.TextFrameContainer;
            var slidePart = slideAutoShape.Slide.SDKSlidePart;
            var hyperlinkRelationship = (HyperlinkRelationship) slidePart.GetReferenceRelationship(hyperlink.Id);

            return hyperlinkRelationship.Uri.AbsoluteUri;
        }

        private void SetHyperlink(string url)
        {
            if (!Uri.IsWellFormedUriString(url, UriKind.Absolute))
            {
                throw new ShapeCrawlerException("Hyperlink is invalid.");
            }

            var runProperties = this.SDKAText.PreviousSibling<A.RunProperties>();
            if (runProperties == null)
            {
                runProperties = new A.RunProperties();
            }

            var hyperlink = runProperties.GetFirstChild<A.HyperlinkOnClick>();
            if (hyperlink == null)
            {
                hyperlink = new A.HyperlinkOnClick();
                runProperties.Append(hyperlink);
            }

            var slideAutoShape = (SlideAutoShape)this.ParentParagraph.ParentTextBox.TextFrameContainer;
            var slidePart = slideAutoShape.Slide.SDKSlidePart;
            
            var uri = new Uri(url, UriKind.Absolute);
            var addedHyperlinkRelationship = slidePart.AddHyperlinkRelationship(uri, true);
            
            hyperlink.Id = addedHyperlinkRelationship.Id;
        }
    }
}