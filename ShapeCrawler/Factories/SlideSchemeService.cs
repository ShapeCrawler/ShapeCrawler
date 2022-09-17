using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using ShapeCrawler.Collections;
using ShapeCrawler.Drawing;
using ShapeCrawler.OLEObjects;
using ShapeCrawler.Shapes;

namespace ShapeCrawler.Factories
{
    /// <summary>
    ///     Represents a slide scheme generator.
    /// </summary>
    internal class SlideSchemeService
    {
        private const int Scale = 10000;
        private const int BitmapOffset = 50;
        private const int RectangleOffset = 10;
        
        public static void SaveScheme(ShapeCollection shapes, int sldW, int sldH, string filePath)
        {
            var bitmap = GetBitmap(shapes, sldW, sldH);
            bitmap.Save(filePath);
            bitmap.Dispose();
        }

        /// <summary>
        ///     Saves in PNG.
        /// </summary>
        public static void SaveScheme(ShapeCollection shapes, int sldW, int sldH, Stream stream)
        {
            var bitmap = GetBitmap(shapes, sldW, sldH);
            bitmap.Save(stream, ImageFormat.Png);
            bitmap.Dispose();
        }
        
        #region Private Methods

        private static Bitmap GetBitmap(ShapeCollection shapes, int sldW, int sldH)
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
                var x = (int) (shape.X / Scale);
                var y = (int) (shape.Y / Scale);
                var w = (int) (shape.Width / Scale);
                var h = (int) (shape.Height / Scale);
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