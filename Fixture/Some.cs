using System;
using System.IO;
using ImageMagick;

namespace Fixture;

public class Fixtures
{
    private readonly Random random = new Random();

    public int Int()
    {
        // Return a positive random integer within a sane range for slide coordinates/sizes
        return this.random.Next(1, 400);
    }

    public Stream Image()
    {
        var width = this.random.Next(32, 256);
        var height = this.random.Next(32, 256);

        var stream = new MemoryStream();

        var background = new MagickColor((byte)this.random.Next(256), (byte)this.random.Next(256), (byte)this.random.Next(256));
        using (var image = new MagickImage(background, (uint)width, (uint)height))
        {
            image.Format = MagickFormat.Png;
            image.Write(stream);
        }

        stream.Position = 0;
        return stream;
    }
}
