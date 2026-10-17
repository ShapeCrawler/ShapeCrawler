using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using FluentAssertions;

namespace ShapeCrawler.Tests;

public class PublicApiTests
{
    [Test]
    public void Public_API_does_not_expose_OpenXml_objects()
    {
        // Arrange
        const BindingFlags flags = BindingFlags.Public | BindingFlags.Instance |
                                   BindingFlags.Static | BindingFlags.DeclaredOnly;
        var publicTypes = typeof(Presentation).Assembly.GetExportedTypes();

        // Act
        var exposedMembers = publicTypes.SelectMany(type => type.GetMembers(flags))
            .Where(member => member switch
            {
                MethodInfo method => IsSdkObject(method.ReturnType) ||
                                     method.GetParameters().Any(parameter => IsSdkObject(parameter.ParameterType)),
                ConstructorInfo constructor => constructor.GetParameters()
                    .Any(parameter => IsSdkObject(parameter.ParameterType)),
                PropertyInfo property => IsSdkObject(property.PropertyType),
                FieldInfo field => IsSdkObject(field.FieldType),
                _ => false
            })
            .Select(member => $"{member.DeclaringType!.FullName}.{member.Name}");

        // Assert
        exposedMembers.Should().BeEmpty("SDK elements and packages must remain implementation details");
    }

    private static bool IsSdkObject(Type type)
    {
        return typeof(OpenXmlElement).IsAssignableFrom(type) ||
               typeof(OpenXmlPartContainer).IsAssignableFrom(type) ||
               (type.HasElementType && IsSdkObject(type.GetElementType()!)) ||
               (type.IsGenericType && type.GetGenericArguments().Any(IsSdkObject));
    }
}
