using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
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
        ///     Checks whether a specified string is not null, empty, or not consists only of white-space characters.
        /// </summary>
        /// <param name="param"></param>
        /// <param name="paramName"></param>
        public static void NotEmpty(string param, string paramName)
        {
            if (string.IsNullOrWhiteSpace(param))
            {
                if (!string.IsNullOrWhiteSpace(paramName))
                {
                    throw new ArgumentException($"{paramName} is empty.");
                }

                throw new ArgumentException("String is empty.");
            }
        }

        /// <summary>
        ///     Checks whether a specified collection is not empty.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="param"></param>
        /// <param name="paramName"></param>
        public static void NotEmpty<T>(IEnumerable<T> param, string paramName)
        {
            NotNull(param, paramName);

            if (!param.Any())
            {
                if (!string.IsNullOrWhiteSpace(paramName))
                {
                    throw new ArgumentException($"Collection {paramName} is empty.");
                }

                throw new ArgumentException("Collection is empty.");
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

        /// <summary>
        ///     Determines whether a number is positive.
        /// </summary>
        public static void IsPositive(int number, string paramName)
        {
            if (number < 1)
            {
                throw new ArgumentOutOfRangeException(paramName);
            }
        }
    }
}