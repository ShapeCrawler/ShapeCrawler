using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;

namespace ShapeCrawler.Charts;

internal class CategoryChart(Categories categories) : Chart()
{
    public override IReadOnlyList<ICategory> Categories => categories;
}