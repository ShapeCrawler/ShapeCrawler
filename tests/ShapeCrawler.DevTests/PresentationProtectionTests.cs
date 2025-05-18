using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.DevTests.Helpers;
using ShapeCrawler.Protection;

namespace ShapeCrawler.DevTests;

public class PresentationProtectionTests : SCTest
{
    [Test]
    public void Encrypt_sets_IsEncrypted_true()
    {
        // Arrange
        var pres = new Presentation(TestAsset("001.pptx"));
        
        // Act
        pres.Protection.Encrypt("password");
        
        // Assert
        pres.Protection.IsEncrypted.Should().BeTrue();
    }
}
