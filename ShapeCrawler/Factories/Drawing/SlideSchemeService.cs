using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Charts;
using ShapeCrawler.Collections;
using ShapeCrawler.Models;
using ShapeCrawler.OLEObjects;
using ShapeCrawler.Pictures;
using ShapeCrawler.SlideMaster;
using ShapeCrawler.Tables;

namespace ShapeCrawler.Factories.Drawing
{
    /// <summary>
    /// Represents a slide scheme generator.
    /// </summary>
    public class SlideSchemeService
    {
        private const int Scale = 10000;
        private const int BitmapOffset = 50;
        private const int RectangleOffset = 10;

        #region Public Methods

        public static void SaveScheme(ShapesCollection shapes, int sldW, int sldH, string filePath)
        {
            var bitmap = GetBitmap(shapes, sldW, sldH);
            bitmap.Save(filePath);
            bitmap.Dispose();
        }

        /// <summary>
        /// Saves in PNG.
        /// </summary>
        /// <param name="shapes"></param>
        /// <param name="sldW"></param>
        /// <param name="sldH"></param>
        /// <param name="stream"></param>
        public static void SaveScheme(ShapesCollection shapes, int sldW, int sldH, Stream stream)
        {
            var bitmap = GetBitmap(shapes, sldW, sldH);
            bitmap.Save(stream, ImageFormat.Png);
            bitmap.Dispose();
        }

        #endregion Public Methods

        #region Private Methods

        private static Bitmap GetBitmap(ShapesCollection shapes, int sldW, int sldH)
        {
            var sldWidthPx = sldW / Scale;
            var sldHeightPx = sldH / Scale;

            // Prepare scheme bitmap
            var bitmap = new Bitmap(sldWidthPx + BitmapOffset, sldHeightPx + BitmapOffset);
            var graphics = Graphics.FromImage(bitmap);

            // Draw slide rectangle
            var sldRectangle = new Rectangle(RectangleOffset, RectangleOffset, sldWidthPx, sldHeightPx);
            using var blackPen = new Pen(Color.Black, 3);
            graphics.DrawRectangle(blackPen, sldRectangle);

            // Draw shape rectangles
            foreach (IShape shape in shapes)
            {
                var x = (int)(shape.X / Scale);
                var y = (int)(shape.Y / Scale);
                var w = (int)(shape.Width / Scale);
                var h = (int)(shape.Height / Scale);
                var shapeRectangle = new Rectangle(x, y, w, h);
                switch (shape)
                {
                    case IAutoShape:
                        graphics.DrawRectangle(Pens.Red, shapeRectangle);
                        break;
                    case IPicture:
                        graphics.DrawRectangle(Pens.Blue, shapeRectangle);
                        break;
                    case IChart:
                        graphics.DrawRectangle(Pens.Aqua, shapeRectangle);
                        break;
                    case IOLEObject:
                        graphics.DrawRectangle(Pens.Bisque, shapeRectangle);
                        break;
                    case ITable:
                        graphics.DrawRectangle(Pens.Chartreuse, shapeRectangle);
                        break;
                }
            }

            return bitmap;
        }

        #endregion Private Methods
    }
}