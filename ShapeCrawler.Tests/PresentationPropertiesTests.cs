using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Tests.Helpers;
using C = DocumentFormat.OpenXml.CustomProperties;
using V = DocumentFormat.OpenXml.VariantTypes;

namespace ShapeCrawler.Tests;

public class PresentationPropertiesTests : SCTest
{
    [TestCaseSource(nameof(CustomValues))]
    [SetCulture("de-DE")]
    public void Custom_properties_round_trip_supported_types(object value, string xmlType)
    {
        // Arrange
        using var pres = new Presentation(p => p.Slide());

        // Act
        pres.Properties["Review"] = value;

        // Assert
        pres.Properties["review"].Should().Be(value);
        using var saved = new MemoryStream();
        pres.Save(saved);
        using (var sdk = PresentationDocument.Open(saved, false))
        {
            var property = sdk.CustomFilePropertiesPart!.Properties!.Elements<C.CustomDocumentProperty>().Single();
            property.FirstChild!.LocalName.Should().Be(xmlType);
            property.FormatId!.Value.Should().Be("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}");
            property.PropertyId!.Value.Should().BeGreaterThanOrEqualTo(2);
        }

        using var reopened = new Presentation(saved);
        var actual = reopened.Properties["REVIEW"];
        actual.Should().BeOfType(value.GetType());
        actual.Should().Be(value);
        ValidatePresentation(reopened);
    }

    [TestCase("")]
    [TestCase(" ")]
    [TestCase("Owner & <team>")]
    [TestCase("Отдел")]
    public void Custom_property_names_round_trip(string name)
    {
        // Arrange
        using var pres = new Presentation(p => p.Slide());

        // Act
        pres.Properties[name] = "Value";

        // Assert
        using var saved = new MemoryStream();
        pres.Save(saved);
        using var reopened = new Presentation(saved);
        reopened.Properties.ContainsCustomProperty(name).Should().BeTrue();
        reopened.Properties.CustomPropertyNames.Should().Contain(name);
        reopened.Properties[name].Should().Be("Value");
        ValidatePresentation(reopened);
    }

    [Test]
    public void Custom_property_update_is_case_insensitive_and_preserves_ids_and_unrelated_values()
    {
        // Arrange
        using var pres = new Presentation(p => p.Slide());
        pres.Properties.Title = "Built in";
        pres.Properties["Title"] = "Custom title";
        pres.Properties["ReviewedBy"] = "Alice";
        pres.Properties["Version"] = 3;
        using var before = new MemoryStream();
        pres.Save(before);
        int originalId;
        using (var sdk = PresentationDocument.Open(before, false))
        {
            originalId = sdk.CustomFilePropertiesPart!.Properties!.Elements<C.CustomDocumentProperty>()
                .Single(property => property.Name == "ReviewedBy").PropertyId!.Value;
        }

        // Act
        pres.Properties["reviewedBY"] = true;

        // Assert
        pres.Properties.CustomPropertyNames.Should().Equal("ReviewedBy", "Title", "Version");
        pres.Properties["REVIEWEDBY"].Should().Be(true);
        pres.Properties["Version"].Should().Be(3);
        pres.Properties.Title.Should().Be("Built in");
        pres.Properties["Title"].Should().Be("Custom title");
        using var saved = new MemoryStream();
        pres.Save(saved);
        using var result = PresentationDocument.Open(saved, false);
        var updated = result.CustomFilePropertiesPart!.Properties!.Elements<C.CustomDocumentProperty>()
            .Single(property => property.Name == "ReviewedBy");
        updated.PropertyId!.Value.Should().Be(originalId);
        updated.FirstChild.Should().BeOfType<V.VTBool>();
        ValidatePresentation(pres);
    }

    [Test]
    public void Missing_custom_properties_are_read_without_creating_a_part()
    {
        // Arrange
        using var pres = new Presentation(p => p.Slide());
        pres.Properties.ClearCustomProperties();

        // Act
        var names = pres.Properties.CustomPropertyNames;
        var contains = pres.Properties.ContainsCustomProperty("Missing");
        var removed = pres.Properties.RemoveCustomProperty("Missing");
        var read = () => pres.Properties["Missing"];

        // Assert
        names.Should().BeEmpty();
        contains.Should().BeFalse();
        removed.Should().BeFalse();
        read.Should().Throw<KeyNotFoundException>();
        using var saved = new MemoryStream();
        pres.Save(saved);
        using var sdk = PresentationDocument.Open(saved, false);
        sdk.CustomFilePropertiesPart.Should().BeNull();
    }

    [Test]
    public void RemoveCustomProperty_removes_only_the_named_property_and_is_idempotent()
    {
        // Arrange
        using var pres = new Presentation(p => p.Slide());
        pres.Properties["Remove"] = 1;
        pres.Properties["Keep"] = "Remaining";

        // Act
        var removed = pres.Properties.RemoveCustomProperty("REMOVE");
        var repeated = pres.Properties.RemoveCustomProperty("Remove");

        // Assert
        removed.Should().BeTrue();
        repeated.Should().BeFalse();
        using var saved = new MemoryStream();
        pres.Save(saved);
        using var reopened = new Presentation(saved);
        reopened.Properties.CustomPropertyNames.Should().Equal("Keep");
        reopened.Properties["Keep"].Should().Be("Remaining");
        ValidatePresentation(reopened);
    }

    [Test]
    public void ClearCustomProperties_preserves_builtin_properties_and_allows_adding_again()
    {
        // Arrange
        using var pres = new Presentation(p => p.Slide());
        pres.Properties.Title = "Report";
        pres.Properties.Author = "Owner";
        pres.Properties["First"] = 1;
        pres.Properties["Second"] = "Two";
        var properties = pres.Properties;

        // Act
        properties.ClearCustomProperties();
        properties.ClearCustomProperties();

        // Assert
        properties.CustomPropertyNames.Should().BeEmpty();
        properties.Title.Should().Be("Report");
        properties.Author.Should().Be("Owner");
        using var empty = new MemoryStream();
        pres.Save(empty);
        using (var sdk = PresentationDocument.Open(empty, false))
        {
            sdk.CustomFilePropertiesPart.Should().BeNull();
        }

        // Act
        properties["Again"] = false;

        // Assert
        using var saved = new MemoryStream();
        pres.Save(saved);
        using var reopened = new Presentation(saved);
        reopened.Properties.CustomPropertyNames.Should().Equal("Again");
        reopened.Properties["Again"].Should().Be(false);
        ValidatePresentation(reopened);
    }

    [Test]
    public void Custom_property_update_preserves_other_variant_types_and_uses_unique_ids()
    {
        // Arrange
        using var source = new Presentation(p => p.Slide());
        using var input = new MemoryStream();
        source.Save(input);
        string unknownXml;
        using (var sdk = PresentationDocument.Open(input, true))
        {
            var part = sdk.CustomFilePropertiesPart ?? sdk.AddCustomFilePropertiesPart();
            var unsupported = new C.CustomDocumentProperty(new V.VTInt64("1234567890123"))
            {
                Name = "Unknown", PropertyId = 99, FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
            };
            part.Properties = new C.Properties(
                new C.CustomDocumentProperty(new V.VTLPSTR("Old"))
                {
                    Name = "Legacy", PropertyId = 10, FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
                }, unsupported);
            unknownXml = unsupported.OuterXml;
        }

        using var pres = new Presentation(input);
        var legacy = pres.Properties["Legacy"];

        // Act
        pres.Properties["legacy"] = 42;
        pres.Properties["New"] = "Added";

        // Assert
        legacy.Should().Be("Old");
        var readUnsupported = () => pres.Properties["Unknown"];
        readUnsupported.Should().Throw<NotSupportedException>();
        using var saved = new MemoryStream();
        pres.Save(saved);
        using var result = PresentationDocument.Open(saved, false);
        var elements = result.CustomFilePropertiesPart!.Properties!.Elements<C.CustomDocumentProperty>().ToArray();
        elements.Single(property => property.Name == "Unknown").OuterXml.Should().Be(unknownXml);
        elements.Single(property => property.Name == "Legacy").PropertyId!.Value.Should().Be(10);
        elements.Single(property => property.Name == "New").PropertyId!.Value.Should().Be(100);
        elements.Select(property => property.PropertyId!.Value).Should().OnlyHaveUniqueItems();
        ValidatePresentation(pres);
    }

    [TestCaseSource(nameof(InvalidCustomValues))]
    public void Invalid_custom_values_do_not_change_existing_properties(object value)
    {
        // Arrange
        using var pres = new Presentation(p => p.Slide());
        pres.Properties["Existing"] = 42;

        // Act
        var write = () => pres.Properties["existing"] = value;

        // Assert
        write.Should().Throw<ArgumentException>();
        pres.Properties["Existing"].Should().Be(42);
        pres.Properties.CustomPropertyNames.Should().Equal("Existing");
    }

    [Test]
    public void Invalid_custom_names_and_xml_values_do_not_create_a_part()
    {
        // Arrange
        using var pres = new Presentation(p => p.Slide());
        pres.Properties.ClearCustomProperties();

        // Act
        var nullName = () => pres.Properties[null!] = "Value";
        var invalidName = () => pres.Properties["Bad\0Name"] = "Value";
        var invalidText = () => pres.Properties["Name"] = "Bad\0Value";
        var invalidType = () => pres.Properties["Name"] = 1m;

        // Assert
        nullName.Should().Throw<ArgumentNullException>();
        invalidName.Should().Throw<XmlException>();
        invalidText.Should().Throw<XmlException>();
        invalidType.Should().Throw<ArgumentException>();
        using var saved = new MemoryStream();
        pres.Save(saved);
        using var sdk = PresentationDocument.Open(saved, false);
        sdk.CustomFilePropertiesPart.Should().BeNull();
    }

    [TestCase(DateTimeKind.Local)]
    [TestCase(DateTimeKind.Unspecified)]
    public void Custom_dates_are_normalized_to_UTC(DateTimeKind kind)
    {
        // Arrange
        using var pres = new Presentation(p => p.Slide());
        var date = new DateTime(2026, 10, 18, 12, 30, 45, kind);
        var expected = kind == DateTimeKind.Unspecified ? DateTime.SpecifyKind(date, DateTimeKind.Utc) : date.ToUniversalTime();

        // Act
        pres.Properties["Date"] = date;

        // Assert
        using var saved = new MemoryStream();
        pres.Save(saved);
        using var reopened = new Presentation(saved);
        var actual = (DateTime)reopened.Properties["Date"];
        actual.Should().Be(expected);
        actual.Kind.Should().Be(DateTimeKind.Utc);
        ValidatePresentation(reopened);
    }

    private static IEnumerable<TestCaseData> CustomValues()
    {
        yield return new TestCaseData("Text & <tag> Отдел", "lpwstr");
        yield return new TestCaseData(string.Empty, "lpwstr");
        yield return new TestCaseData(42, "i4");
        yield return new TestCaseData(int.MinValue, "i4");
        yield return new TestCaseData(true, "bool");
        yield return new TestCaseData(false, "bool");
        yield return new TestCaseData(1.25f, "r4");
        yield return new TestCaseData(1.23456789012345d, "r8");
        yield return new TestCaseData(new DateTime(2026, 10, 18, 12, 30, 45, DateTimeKind.Utc).AddTicks(1234567), "filetime");
    }

    private static IEnumerable<TestCaseData> InvalidCustomValues()
    {
        foreach (var value in new object?[] { null, 1L, 1m, new object(), float.NaN, float.PositiveInfinity, double.NaN, double.NegativeInfinity })
        {
            yield return new TestCaseData(value);
        }
    }
}
