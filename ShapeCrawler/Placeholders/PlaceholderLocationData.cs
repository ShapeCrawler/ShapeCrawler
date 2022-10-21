namespace ShapeCrawler.Placeholders;

internal class PlaceholderLocationData : PlaceholderData
{
    public PlaceholderLocationData(PlaceholderData phData)
    {
        this.PlaceholderType = phData.PlaceholderType;
        this.Index = phData.Index;
    }
}