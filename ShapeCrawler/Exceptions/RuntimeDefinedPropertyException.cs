namespace ShapeCrawler.Exceptions
{
    /// <summary>
    /// Thrown while attempting to access runtime defined property, but it does not exist for the current object.
    /// </summary>
    public class RuntimeDefinedPropertyException : ShapeCrawlerException
    {
        #region Constructors

        public RuntimeDefinedPropertyException(string message) 
            : base(message, ExceptionCode.RuntimeDefinedPropertyException) { }

        public RuntimeDefinedPropertyException()
        {
        }

        public RuntimeDefinedPropertyException(string message, System.Exception innerException) : base(message, innerException)
        {
        }

        #endregion Constructors
    }
}
