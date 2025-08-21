using ImageMagick;

namespace Fixture;

public class Fixtures
{
    private readonly Random random = new();
    private readonly List<string> files = new();

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

    public Stream Image(Action<ImageOptions> configure)
    {
        var options = new ImageOptions();
        configure(options);

        var width = this.random.Next(32, 256);
        var height = this.random.Next(32, 256);
        var stream = new MemoryStream();

        var background = new MagickColor((byte)this.random.Next(256), (byte)this.random.Next(256), (byte)this.random.Next(256));

        var formatUpper = (options.FormatName ?? "PNG").ToUpperInvariant();
        if (formatUpper == "GIF")
        {
            using var collection = new MagickImageCollection();

            // Create frames with clearly different colors for visible animation
            var colors = new[]
            {
                MagickColors.Red,
                MagickColors.Green,
                MagickColors.Blue,
                MagickColors.Yellow,
                MagickColors.Magenta,
                MagickColors.Cyan
            };

            for (int i = 0; i < colors.Length; i++)
            {
                using var frame = new MagickImage(colors[i], (uint)width, (uint)height);
                frame.Format = MagickFormat.Gif;
                
                // Set animation timing - 50 centiseconds = 0.5 seconds per frame
                frame.AnimationDelay = 50;
                frame.AnimationIterations = 0; // Infinite loop
                frame.GifDisposeMethod = GifDisposeMethod.Background;
                
                collection.Add(frame.Clone());
            }

            collection.Write(stream, MagickFormat.Gif);
        }
        else if (formatUpper is "PNG")
        {
            using var image = new MagickImage(background, (uint)width, (uint)height);
            image.Format = MagickFormat.Png;
            image.Write(stream);
        }
        else if (formatUpper is "JPG" or "JPEG")
        {
            using var image = new MagickImage(background, (uint)width, (uint)height);
            image.Format = MagickFormat.Jpeg;
            image.Write(stream);
        }
        else if (formatUpper is "TIFF" or "TIF")
        {
            using var image = new MagickImage(background, (uint)width, (uint)height);
            image.Format = MagickFormat.Tiff;
            image.Write(stream);
        }
        else
        {
            using var image = new MagickImage(background, (uint)width, (uint)height);
            image.Format = MagickFormat.Png;
            image.Write(stream);
        }

        stream.Position = 0;
        return stream;
    }

    public string File()
    {
        var file = Path.Combine(Directory.GetCurrentDirectory(), Guid.NewGuid().ToString());
        var stream = System.IO.File.Create(file);
        stream.Close();
        
        this.files.Add(file);

        return file;
    }

    public void Clean() => files.ForEach(System.IO.File.Delete);

    public string String() => Guid.NewGuid().ToString();
}

public sealed class ImageOptions
{
    public string? FormatName { get; private set; }

    public void Format(string format)
    {
        this.FormatName = format;
    }
}
