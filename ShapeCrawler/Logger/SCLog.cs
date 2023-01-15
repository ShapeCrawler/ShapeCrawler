using System;
using System.Collections.Generic;

namespace ShapeCrawler.Logger;

internal class SCLog
{
    internal string UserId { get; set; } = "undefined";

    internal string? TargetFramework { get; set; }
    
    internal string? LibraryVersion { get; set; }

    internal DateTime SentDate { get; set; }

    internal List<string>? Errors { get; set; }

    internal void Reset()
    {
        this.SentDate = DateTime.Now;
        this.Errors = null;
    }
}