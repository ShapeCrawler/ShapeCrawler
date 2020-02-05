namespace SlideXML.Models
{
    /// <summary>
    /// Represents a presentation component that has a tag.
    /// </summary>
    public abstract class TaggedComponent
    {
        /// <summary>
        /// Gets or sets a tag or custom property that can store a reference to a user object for a different business scenario.
        /// </summary>
        public object Tag { get; set; }
    }
}
