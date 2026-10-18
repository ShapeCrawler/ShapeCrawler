using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using C = DocumentFormat.OpenXml.CustomProperties;
using V = DocumentFormat.OpenXml.VariantTypes;

namespace ShapeCrawler.Presentations;

internal sealed class CustomProperties(PresentationDocument document)
{
    private const string FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";

    internal IReadOnlyList<string> Names => document.CustomFilePropertiesPart?.Properties?
        .Elements<C.CustomDocumentProperty>().Select(property => property.Name?.Value ?? string.Empty)
        .OrderBy(name => name, StringComparer.OrdinalIgnoreCase).ToArray() ?? [];

    internal object Get(string name)
    {
        var property = this.Find(name) ?? throw new KeyNotFoundException($"Custom property '{name}' was not found.");
        var value = property.FirstChild;
        return value switch
        {
            V.VTLPWSTR or V.VTLPSTR or V.VTBString => value.InnerText,
            V.VTInt32 => XmlConvert.ToInt32(value.InnerText),
            V.VTBool => XmlConvert.ToBoolean(value.InnerText),
            V.VTFloat => XmlConvert.ToSingle(value.InnerText),
            V.VTDouble => XmlConvert.ToDouble(value.InnerText),
            V.VTFileTime => XmlConvert.ToDateTime(value.InnerText, XmlDateTimeSerializationMode.Utc),
            _ => throw new NotSupportedException($"Custom property '{name}' has an unsupported value type '{value?.LocalName}'.")
        };
    }

    internal bool Contains(string name) => this.Find(name) is not null;

    internal void Set(string name, object value)
    {
        ArgumentNullException.ThrowIfNull(name);
        XmlConvert.VerifyXmlChars(name);
        var element = CreateValue(value);
        var property = this.Find(name);
        if (property is not null)
        {
            property.RemoveAllChildren();
            property.Append(element);
            return;
        }

        var root = document.CustomFilePropertiesPart?.Properties;
        var maximumId = root?.Elements<C.CustomDocumentProperty>()
            .Select(item => item.PropertyId?.Value ?? 1).DefaultIfEmpty(1).Max() ?? 1;
        var id = checked(Math.Max(maximumId, 1) + 1);
        if (root is null)
        {
            var part = document.CustomFilePropertiesPart ?? document.AddCustomFilePropertiesPart();
            root = new C.Properties();
            part.Properties = root;
        }

        root.Append(new C.CustomDocumentProperty(element) { Name = name, PropertyId = id, FormatId = FormatId });
    }

    internal bool Remove(string name)
    {
        var property = this.Find(name);
        if (property is null)
        {
            return false;
        }

        property.Remove();
        this.DeleteEmptyPart();
        return true;
    }

    internal void Clear()
    {
        document.CustomFilePropertiesPart?.Properties?.RemoveAllChildren<C.CustomDocumentProperty>();
        this.DeleteEmptyPart();
    }

    private static OpenXmlElement CreateValue(object value)
    {
        ArgumentNullException.ThrowIfNull(value);
        return value switch
        {
            string text => new V.VTLPWSTR(XmlConvert.VerifyXmlChars(text)),
            int number => new V.VTInt32(XmlConvert.ToString(number)),
            bool flag => new V.VTBool(XmlConvert.ToString(flag)),
            float number when float.IsFinite(number) => new V.VTFloat(XmlConvert.ToString(number)),
            double number when double.IsFinite(number) => new V.VTDouble(XmlConvert.ToString(number)),
            DateTime date => new V.VTFileTime(XmlConvert.ToString(
                date.Kind == DateTimeKind.Unspecified ? DateTime.SpecifyKind(date, DateTimeKind.Utc) : date.ToUniversalTime(),
                XmlDateTimeSerializationMode.Utc)),
            _ => throw new ArgumentException("Custom values must be string, int, bool, finite float/double, or DateTime.", nameof(value))
        };
    }

    private C.CustomDocumentProperty? Find(string name)
    {
        ArgumentNullException.ThrowIfNull(name);
        return document.CustomFilePropertiesPart?.Properties?.Elements<C.CustomDocumentProperty>()
            .FirstOrDefault(property => string.Equals(property.Name?.Value, name, StringComparison.OrdinalIgnoreCase));
    }

    private void DeleteEmptyPart()
    {
        var part = document.CustomFilePropertiesPart;
        if (part?.Properties is { HasChildren: false })
        {
            document.DeletePart(part);
        }
    }
}
