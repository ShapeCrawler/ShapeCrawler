namespace SlideXML.Exceptions
{
    /// <summary>
    /// Thrown if a type of element could not be defined.
    /// </summary>
    public class TypeException : SlideXMLException
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="TypeException"/> class with default error message.
        /// </summary>
        public TypeException(): base(101, "Element type was not defined.") { }

        #endregion Constructors
    }
}
