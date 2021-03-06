﻿using System;

namespace ShapeCrawler.Exceptions
{
    public class PlaceholderCannotBeChangedException : ShapeCrawlerException
    {
        private static readonly string ExceptionMessage =
            "The property is part of slide layout placeholder and cannot be changed on a slide level. " +
            $"If you wanna change some placeholder format value, you can do it by using {nameof(SlideGroupShape.Placeholder)}.";

        public PlaceholderCannotBeChangedException(string message) : base(message)
        {
        }

        public PlaceholderCannotBeChangedException(string message, Exception innerException) : base(message,
            innerException)
        {
        }

        public PlaceholderCannotBeChangedException() : base(ExceptionMessage)
        {
        }
    }
}