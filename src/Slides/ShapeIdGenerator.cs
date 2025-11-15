using System;
#pragma warning disable IDE0005 // Using directive is unnecessary
using System.Linq;
#pragma warning restore IDE0005
using System.Text.RegularExpressions;

namespace ShapeCrawler.Slides;

/// <summary>
///     Generates shape IDs and names.
/// </summary>
internal sealed class ShapeIdGenerator(ISlideShapeCollection shapes)
{
    /// <summary>
    ///     Gets the next available shape ID.
    /// </summary>
    internal int GetNextId()
    {
        if (shapes.Any())
        {
            return shapes.Select(shape => shape.Id).Prepend(0).Max() + 1;
        }

        return 1;
    }

    /// <summary>
    ///     Generates a shape ID and default name.
    /// </summary>
    internal (int, string) GenerateIdAndName()
    {
        var id = this.GetNextId();

        return (id, $"Shape {id}");
    }

    /// <summary>
    ///     Generates the next table name based on existing table names.
    /// </summary>
    internal string GenerateNextTableName()
    {
        var maxOrder = 0;
        foreach (var shape in shapes)
        {
            var matchOrder = Regex.Match(shape.Name, "(?!Table )\\d+", RegexOptions.None, TimeSpan.FromSeconds(100));
            if (!matchOrder.Success)
            {
                continue;
            }

            var order = int.Parse(matchOrder.Value);
            if (order > maxOrder)
            {
                maxOrder = order;
            }
        }

        return $"Table {maxOrder + 1}";
    }
}
