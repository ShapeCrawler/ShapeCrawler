using FluentAssertions;

namespace ShapeCrawler.CITests;

public class PresentationTests
{
    [Test]
    public void Save_does_not_throw_exception_When_stream_is_a_File_stream()
    {
        // Arrange
        var pres = new Presentation();
        var file = Path.GetTempFileName();
        using var stream = File.OpenWrite(file);
        
        // Act
        var saving = () => pres.Save(stream);
        
        // Assert
        saving.Should().NotThrow();
        
        // Cleanup
        stream.Close();
        File.Delete(file);
    }
}