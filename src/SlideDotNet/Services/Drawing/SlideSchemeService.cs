using System.Drawing;
using SlideDotNet.Enums;
using SlideDotNet.Models;

namespace SlideDotNet.Services.Drawing
{
    public class SlideSchemeService : ISlideSchemeService
    {
        private const int Scale = 10000;
        private const int BitmapOffset = 50;
        private const int RectangleOffset = 10;

        public void SaveScheme(string filePath, ShapeCollection shapesValue, int sldW, int sldH)
        {
            var sldWidthPx = sldW / Scale;
            var sldHeightPx = sldH / Scale;

            // Prepare scheme bitmap
            using var bitmap = new Bitmap(sldWidthPx + BitmapOffset, sldHeightPx + BitmapOffset);
            var graphics = Graphics.FromImage(bitmap);

            // Draw slide rectangle
            var sldRectangle = new Rectangle(RectangleOffset, RectangleOffset, sldWidthPx, sldHeightPx);
            using var blackPen = new Pen(Color.Black, 3);
            graphics.DrawRectangle(blackPen, sldRectangle);

            // Draw shape rectangles
            foreach (var shape in shapesValue)
            {
                var x = shape.X / Scale;
                var y = shape.Y / Scale;
                var w = shape.Width / Scale;
                var h = shape.Height / Scale;
                var shapeRectangle = new Rectangle((int)x, (int)y, (int)w, (int)h);
                switch (shape.ContentType)
                {
                    case ShapeContentType.AutoShape:
                        graphics.DrawRectangle(Pens.Red, shapeRectangle);
                        break;
                    case ShapeContentType.Picture:
                        graphics.DrawRectangle(Pens.Blue, shapeRectangle);
                        break;
                    case ShapeContentType.Chart:
                        graphics.DrawRectangle(Pens.Aqua, shapeRectangle);
                        break;
                    case ShapeContentType.OLEObject:
                        graphics.DrawRectangle(Pens.Bisque, shapeRectangle);
                        break;
                    case ShapeContentType.Table:
                        graphics.DrawRectangle(Pens.Chartreuse, shapeRectangle);
                        break;
                }
            }

            bitmap.Save(filePath);
        }
    }
}