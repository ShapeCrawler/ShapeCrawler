using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml;

namespace ShapeCrawler.Shared;

/// <summary>
/// This is a generic enum.
/// </summary>
[DebuggerDisplay("{Name}")]
public abstract class SCEnumeration
{
    /// <summary>
    /// Initializes a new instance of the <see cref="SCEnumeration"/> class.
    /// </summary>
    /// <param name="value">Enum value.</param>
    /// <param name="name">Enum name.</param>
    protected SCEnumeration(string value, string name)
    {
        (this.Value, this.Name) = (value, name);
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
/// <typeparam name="T">Enum type.</typeparam>
public abstract class SCEnumeration
    <T> : SCEnumeration
    where T : SCEnumeration
{
    /// <summary>
    /// Initializes a new instance of the <see cref="SCEnumeration{T}"/> class.
    /// </summary>
    /// <param name="value">Enum value.</param>
    /// <param name="name">Enum name.</param>
    protected SCEnumeration(string value, string name)
        : base(value, name)
    {
    }

    /// <summary>
    /// Returns an enum member of <typeparamref name="T"/>.
    /// </summary>
    /// <param name="value">Enum value.</param>
    /// <returns>An enum member.</returns>
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
    /// <typeparam name="TValue">Value of the schema: type.</typeparam>
    /// <param name="value">Enum value.</param>
    /// <param name="result">Enum member.</param>
    /// <returns>Returns <see langword="true"/> <paramref name="value"/> exists in <typeparamref name="T"/>.</returns>
    public static bool TryParse<TValue>(EnumValue<TValue>? value, out T? result)
        where TValue : struct
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
