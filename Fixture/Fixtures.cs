using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using ImageMagick;

namespace Fixture;

public class Fixtures
{
    private readonly Random random = new();
    private readonly List<string> files = new();
    private readonly Assembly? assembly;

    public Fixtures()
    {
        
    }
    
    public Fixtures(Assembly assembly)
    {
        this.assembly = assembly;
    }

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

    /// <summary>
    ///     Gets a path to a temporary file.
    /// </summary>
    /// <returns></returns>
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

    public string String(Action<StringOptions> configure)
    {
        var options = new StringOptions();
        configure(options);

        var targetLength = options.LengthValue ?? 36;
        if (targetLength <= 0)
        {
            return string.Empty;
        }

        // Generate a readable random string with spaces that can wrap in text boxes
        // while guaranteeing exact requested length.
        const string letters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";

        var buffer = new char[targetLength];
        bool lastWasSpace = false;
        for (int i = 0; i < targetLength; i++)
        {
            bool canBeSpace = i != 0 && i != targetLength - 1 && !lastWasSpace;
            // Roughly 1/6 chance to place a space when allowed to encourage wrapping
            if (canBeSpace && this.random.Next(6) == 0)
            {
                buffer[i] = ' ';
                lastWasSpace = true;
                continue;
            }

            buffer[i] = letters[this.random.Next(letters.Length)];
            lastWasSpace = false;
        }

        return new string(buffer);
    }
    
    public Stream AssemblyFile(string file)
    {
        var stream = GetResourceStream(file);
        var mStream = new MemoryStream();
        stream.CopyTo(mStream);
        mStream.Position = 0;

        return mStream;
    }
    
    private MemoryStream GetResourceStream(string fileName)
    {
        var pattern = $@"\.{Regex.Escape(fileName)}";
        var path = this.assembly.GetManifestResourceNames().First(r =>
        {
            var matched = Regex.Match(r, pattern, RegexOptions.None, TimeSpan.FromSeconds(1));
            return matched.Success;
        });
        var stream = assembly.GetManifestResourceStream(path);
        var mStream = new MemoryStream();
        stream!.CopyTo(mStream);
        mStream.Position = 0;

        return mStream;
    }
}