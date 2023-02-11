// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a field.
/// </summary>
public interface IField
{
}

internal sealed class SCField : IField
{
    public SCField(DocumentFormat.OpenXml.Drawing.Field aField)
    {
    }
}