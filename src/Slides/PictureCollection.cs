using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ImageMagick;
using ShapeCrawler.Extensions;
using ShapeCrawler.Presentations;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable UseObjectOrCollectionInitializer
namespace ShapeCrawler.Slides;

internal sealed class PictureCollection(
    ISlideShapeCollection shapes,
    PresentationImageFiles imageFiles,
    SlidePart slidePart
) : ISlideShapeCollection
{
    public int Count => shapes.Count;

    public IShape this[int index] => shapes[index];

    public void AddPicture(Stream imageStream)
    {
        try
        {
            var imageContent = new Image(imageStream);

            P.Picture pPicture;
            if (imageContent.IsSvg)
            {
                var rasterStream = imageContent.GetRasterStream();
                var svgStream = imageContent.GetOriginalStream();

                var svgHash = imageContent.SvgHash;
                if (!this.TryGetImageRId(svgHash, out var svgPartRId))
                {
                    svgPartRId = slidePart.AddImagePart(svgStream, "image/svg+xml");
                }

                var imgHash = imageContent.Hash;
                if (!this.TryGetImageRId(imgHash, out var imgPartRId))
                {
                    imgPartRId = slidePart.AddImagePart(rasterStream, "image/png");
                }

                var xmlPicture = new XmlPicture(slidePart, (uint)this.GetNextShapeId(), "Picture");
                pPicture = xmlPicture.CreateSvgPPicture(imgPartRId, svgPartRId);
            }
            else
            {
                var imageForPart =

                    // Preserve original bytes for supported formats to ensure deterministic dedup across slides
                    imageContent.IsOriginalFormatPreserved ? imageContent.GetOriginalStream() :

                    // For formats that we convert (e.g., WebP/AVIF/BMP), write a deterministic raster representation
                    imageContent.GetRasterStream();

                var hash = imageContent.Hash;
                if (!this.TryGetImageRId(hash, out var imgPartRId))
                {
                    imgPartRId = slidePart.AddImagePart(imageForPart, imageContent.MimeType);
                }

                var xmlPicture = new XmlPicture(slidePart, (uint)this.GetNextShapeId(), "Picture");
                pPicture = xmlPicture.CreatePPicture(imgPartRId);
            }

            XmlPicture.SetTransform(pPicture, imageContent.Width, imageContent.Height);
        }
        catch (Exception ex) when (ex is MagickDelegateErrorException mex && mex.Message.Contains("ghostscript"))
        {
            throw new SCException(
                "The stream is an image format that requires GhostScript which is not installed on your system.", ex);
        }
        catch (MagickException)
        {
            throw new SCException(
                "The stream is not an image or a non-supported image format. Contact us for support: https://github.com/ShapeCrawler/ShapeCrawler/discussions");
        }
    }

    #region Shapes Public Methods

    public void AddAudio(int x, int y, Stream audio) => this.AddAudio(x, y, audio, AudioType.MP3);

    public void AddAudio(int x, int y, Stream audio, AudioType type) => shapes.AddAudio(x, y, audio, type);

    public void AddVideo(int x, int y, Stream stream) => shapes.AddVideo(x, y, stream);

    public void Add(IShape addingShape) => shapes.Add(addingShape);

    public void AddShape(
        int x,
        int y,
        int width,
        int height,
        Geometry geometry = Geometry.Rectangle
    ) => shapes.AddShape(x, y, width, height, geometry);

    public void AddShape(
        int x,
        int y,
        int width,
        int height,
        Geometry geometry,
        string text
    ) => shapes.AddShape(x, y, width, height, geometry, text);

    public void AddLine(string xml) => shapes.AddLine(xml);

    public void AddLine(
        int startPointX,
        int startPointY,
        int endPointX,
        int endPointY
    ) => shapes.AddLine(startPointX, startPointY, endPointX, endPointY);

    public void AddTable(
        int x,
        int y,
        int columnsCount,
        int rowsCount
    ) => shapes.AddTable(x, y, columnsCount, rowsCount);

    public void AddTable(
        int x,
        int y,
        int columnsCount,
        int rowsCount,
        ITableStyle style
    ) => shapes.AddTable(x, y, columnsCount, rowsCount, style);

    public void AddPieChart(
        int x,
        int y,
        int width,
        int height,
        Dictionary<string, double> categoryValues,
        string seriesName
    ) => shapes.AddPieChart(x, y, width, height, categoryValues, seriesName);

    public void AddPieChart(
        int x,
        int y,
        int width,
        int height,
        Dictionary<string, double> categoryValues,
        string seriesName,
        string chartName
    ) => shapes.AddPieChart(x, y, width, height, categoryValues, seriesName, chartName);

    public void AddBarChart(
        int x,
        int y,
        int width,
        int height,
        Dictionary<string, double> categoryValues,
        string seriesName
    ) => shapes.AddBarChart(x, y, width, height, categoryValues, seriesName);

    public void AddScatterChart(
        int x,
        int y,
        int width,
        int height,
        Dictionary<double, double> pointValues,
        string seriesName
    ) => shapes.AddScatterChart(x, y, width, height, pointValues, seriesName);

    public void AddStackedColumnChart(
        int x,
        int y,
        int width,
        int height,
        IDictionary<string, IList<double>> categoryValues,
        IList<string> seriesNames
    ) => shapes.AddStackedColumnChart(x, y, width, height, categoryValues, seriesNames);

    public void AddClusteredBarChart(
        int x,
        int y,
        int width,
        int height,
        IList<string> categories,
        IList<Presentations.DraftChart.SeriesData> seriesData,
        string chartName
    ) => shapes.AddClusteredBarChart(x, y, width, height, categories, seriesData, chartName);

    public IShape AddSmartArt(
        int x,
        int y,
        int width,
        int height,
        SmartArtType smartArtType
    ) => shapes.AddSmartArt(x, y, width, height, smartArtType);

    public IShape Group(IShape[] groupingShapes)
    {
        throw new NotImplementedException();
    }

    public IShape AddDateAndTime() => shapes.AddDateAndTime();

    public IShape AddFooter() => shapes.AddFooter();

    public IShape AddSlideNumber() => shapes.AddSlideNumber();

    public IEnumerator<IShape> GetEnumerator() => shapes.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => shapes.GetEnumerator();

    public IShape GetById(int id) => shapes.GetById<IShape>(id);

    public T GetById<T>(int id)
        where T : IShape => shapes.GetById<T>(id);

    public IShape Shape(string name) => shapes.Shape<IShape>(name);

    public T Shape<T>(string name)
        where T : IShape => shapes.Shape<T>(name);

    public T Last<T>()
        where T : IShape => shapes.Last<T>();

    #endregion Shapes Public Methods

    private int GetNextShapeId()
    {
        if (shapes.Any())
        {
            return shapes.Select(shape => shape.Id).Prepend(0).Max() + 1;
        }

        return 1;
    }

    private bool TryGetImageRId(string hash, out string imgPartRId)
    {
        var imagePart = imageFiles.ImagePartByImageHashOrNull(hash);
        if (imagePart is not null)
        {
            // Image already exists in the presentation so far.
            // Do we have a reference to it on this slide?
            var found = slidePart.ImageParts.Where(x => x.Uri == imagePart.Uri);
            if (found.Any())
            {
                // Yes, we already have a relationship with this part on this slide
                // So use that relationship ID
                imgPartRId = slidePart.GetIdOfPart(imagePart);
            }
            else
            {
                // No, so let's create a relationship to it
                imgPartRId = slidePart.CreateRelationshipToPart(imagePart);
            }

            return true;
        }

        // Sorry, you'll need to create a new image part
        imgPartRId = string.Empty;
        return false;
    }
}