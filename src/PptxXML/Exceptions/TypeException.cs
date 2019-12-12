namespace PptxXML.Exceptions
{
    /// <summary>
    /// Thrown when a type of element could not be defined.
    /// </summary>
    public class TypeException : PptxXMLException
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="TypeException"/> class with default error message.
        /// </summary>
        public TypeException(): base(101, "Element type was not defined.") { }

        #endregion Constructors
    }
}
