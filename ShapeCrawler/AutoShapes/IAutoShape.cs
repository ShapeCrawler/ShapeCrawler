using ShapeCrawler.Models;
using ShapeCrawler.Models.Styles;
using ShapeCrawler.Texts;

namespace ShapeCrawler.AutoShapes
{
    public interface IAutoShape : IShape
    {
        ShapeFill Fill { get; }
        TextBoxSc TextBox { get; }
    }
}