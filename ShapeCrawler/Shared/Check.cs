using System;
using System.Diagnostics.CodeAnalysis;
using System.Text.RegularExpressions;

namespace ShapeCrawler.Shared
{
    /// <summary>
    ///     Represents parameter checker.
    /// </summary>
    public static class Check
    {
        /// <summary>
        ///     Checks whether specified object is not null.
        /// </summary>
        /// <param name="param"></param>
        /// <param name="paramName"></param>
        [SuppressMessage("ReSharper", "InvertIf")]
        public static void NotNull(object param, string paramName)
        {
            if (param == null)
            {
                if (!string.IsNullOrWhiteSpace(paramName))
                {
                    throw new ArgumentNullException(paramName);
                }

                throw new ArgumentNullException(paramName);
            }
        }

        /// <summary>
        ///     Determines whether a string is a valid email address.
        /// </summary>
        /// <param name="paramEmail"></param>
        /// <returns></returns>
        /// <remarks>Regex pattern was taken from https://bit.ly/33dw7C3 </remarks>
        public static bool IsEmail(string paramEmail)
        {
            if (string.IsNullOrWhiteSpace(paramEmail))
            {
                return false;
            }

            const string validEmailPattern = @"^(?!\.)(""([^""\r\\]|\\[""\r\\])*""|"
                                             + @"([-a-z0-9!#$%&'*+/=?^_`{|}~]|(?<!\.)\.)*)(?<!\.)"
                                             + @"@[a-z0-9][\w\.-]*[a-z0-9]\.[a-z][a-z\.]*[a-z]$";
            var validEmailRegex = new Regex(validEmailPattern, RegexOptions.IgnoreCase);

            return validEmailRegex.IsMatch(paramEmail);
        }
    }
}