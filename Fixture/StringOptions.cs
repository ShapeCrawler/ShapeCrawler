namespace Fixture;

public sealed class StringOptions
{
    public int? LengthValue { get; private set; }

    public void Length(int length)
    {
        this.LengthValue = length;
    }
}