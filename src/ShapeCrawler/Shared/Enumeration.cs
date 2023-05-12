using DocumentFormat.OpenXml;
using ShapeCrawler.Enums;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;

namespace ShapeCrawler.Shared;

/// <summary>
/// This is a generic enum.
/// </summary>
[DebuggerDisplay("{Name}")]
public abstract class Enumeration
{
    /// <summary>
    /// Initializes a new instance of the <see cref="Enumeration"/> class.
    /// </summary>
    /// <param name="value">Enum value.</param>
    /// <param name="name">Enum name.</param>
    internal protected Enumeration(string value, string name)
    {
        (Value, Name) = (value, name);
    }

    /// <summary>
    /// Gets the enum value.
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// Gets the enum value.
    /// </summary>
    public string Value { get; }

    /// <inheritdoc/>
    public override string ToString()
    {
        return this.Value;
    }
}

/// <summary>
/// This is a generic enum.
/// </summary>
public abstract class Enumeration<T> : Enumeration, IEnumeration where T : Enumeration
{
    /// <summary>
    /// Initializes a new instance of the <see cref="Enumeration{T}"/> class.
    /// </summary>
    /// <param name="value">Enum value.</param>
    /// <param name="name">Enum name.</param>
    internal protected Enumeration(string value, string name)
        : base(value, name)
    {
    }

    /// <summary>
    /// Returns an enum member of <typeparamref name="T"/>.
    /// </summary>
    /// <param name="value">Enum value.</param>
    /// <returns></returns>
    /// <exception cref="Exception">Throws when value doesn't exists.</exception>
    public static T Parse(string value)
    {
        if (TryParse(value, out T? result))
        {
            return result!;
        }

        throw new Exception();
    }

    /// <summary>
    /// Try to get a type from string value.
    /// </summary>
    /// <param name="value">Type value.</param>
    /// <param name="result">Enum member.</param>
    /// <returns>Returns <see langword="true"/> <paramref name="value"/> exists in <typeparamref name="T"/>.</returns>
    public static bool TryParse(string value, out T? result)
    {
        result = GetAll()
            .FirstOrDefault(item => item.Value == value);

        return result is not null;
    }

    /// <summary>
    /// Try to get a type from string value.
    /// </summary>
    /// <typeparam name="V">Value of the schema: type.</typeparam>
    /// <param name="value">Enum value.</param>
    /// <param name="result">Enum member.</param>
    /// <returns>Returns <see langword="true"/> <paramref name="value"/> exists in <typeparamref name="T"/>.</returns>
    public static bool TryParse<V>(EnumValue<V>? value, out T? result) where V : struct
    {
        if (value is null)
        {
            result = null;

            return false;
        }

        return TryParse(value?.InnerText ?? string.Empty, out result);
    }

    /// <summary>
    /// Gets all public and static members of type <typeparamref name="T"/>.
    /// </summary>
    /// <returns>A member list of type value.</returns>
    protected static IEnumerable<T> GetAll() 
    {
        return typeof(T).GetFields(BindingFlags.Public |
                            BindingFlags.Static |
                            BindingFlags.DeclaredOnly)
                 .Select(f => f.GetValue(null))
                 .Cast<T>();
    }
}
