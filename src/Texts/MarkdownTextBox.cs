using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Texts;

internal sealed class MarkdownTextBox(P.TextBody textBody, SlidePart slidePart)
{
    internal void Write(MarkdownDocument document)
    {
        textBody.RemoveAllChildren<A.Paragraph>();
        this.WriteBlocks(document, 0);
        if (!textBody.Elements<A.Paragraph>().Any())
        {
            textBody.Append(new A.Paragraph(new A.ParagraphProperties(new A.NoBullet())));
        }
    }

    private void WriteBlocks(ContainerBlock blocks, int depth)
    {
        foreach (var block in blocks)
        {
            if (block is ThematicBreakBlock)
            {
                continue;
            }

            if (block is ListBlock list)
            {
                this.WriteList(list, depth);
            }
            else if (block is ContainerBlock container)
            {
                this.WriteBlocks(container, depth);
            }
            else if (block is LeafBlock leaf)
            {
                this.WriteLeaf(leaf);
            }
        }
    }

    private void WriteList(ListBlock list, int depth)
    {
        var number = int.TryParse(list.OrderedStart, out var start) ? start : 1;
        foreach (ListItemBlock item in list)
        {
            var first = true;
            foreach (var child in item)
            {
                if (child is LeafBlock leaf)
                {
                    var paragraph = this.WriteLeaf(leaf);
                    var properties = paragraph.ParagraphProperties!;
                    properties.Level = Math.Min(depth, 8);
                    properties.LeftMargin = (Math.Min(depth, 8) + 1) * 285750;
                    if (first)
                    {
                        properties.Indent = -285750;
                        properties.RemoveAllChildren<A.NoBullet>();
                        if (list.IsOrdered)
                        {
                            properties.Append(new A.AutoNumberedBullet
                            {
                                Type = A.TextAutoNumberSchemeValues.ArabicPeriod,
                                StartAt = Math.Clamp(number, 1, 32767)
                            });
                        }
                        else
                        {
                            properties.Append(new A.CharacterBullet { Char = "•" });
                        }
                    }

                    first = false;
                }
                else if (child is ContainerBlock container)
                {
                    if (container is ListBlock nestedList)
                    {
                        this.WriteList(nestedList, depth + 1);
                    }
                    else
                    {
                        this.WriteBlocks(container, depth + 1);
                    }
                }
            }

            number++;
        }
    }

    private A.Paragraph WriteLeaf(LeafBlock leaf)
    {
        var paragraph = new A.Paragraph(new A.ParagraphProperties(new A.NoBullet())
        {
            Alignment = A.TextAlignmentTypeValues.Left
        });
        textBody.Append(paragraph);
        if (leaf.Inline is { } inline)
        {
            this.WriteInlines(paragraph, inline, leaf is HeadingBlock, false, null);
        }
        else
        {
            var lines = leaf.Lines.ToString().Replace("\r\n", "\n").Split('\n');
            for (var i = 0; i < lines.Length; i++)
            {
                if (i > 0)
                {
                    paragraph.Append(new A.Break());
                }

                AddRun(paragraph, lines[i], false, false, leaf is CodeBlock, null);
            }
        }

        return paragraph;
    }

    private void WriteInlines(A.Paragraph paragraph, ContainerInline inlines, bool bold, bool italic, string? url)
    {
        foreach (var inline in inlines)
        {
            switch (inline)
            {
                case EmphasisInline emphasis:
                    this.WriteInlines(paragraph, emphasis, bold || emphasis.DelimiterCount == 2,
                        italic || emphasis.DelimiterCount == 1, url);
                    break;
                case LinkInline link:
                    this.WriteInlines(paragraph, link, bold, italic, link.IsImage ? url : link.Url);
                    break;
                case AutolinkInline link:
                    AddRun(paragraph, link.Url, bold, italic, false, link.IsEmail ? "mailto:" + link.Url : link.Url);
                    break;
                case LineBreakInline lineBreak:
                    if (lineBreak.IsHard)
                    {
                        paragraph.Append(new A.Break());
                    }
                    else
                    {
                        AddRun(paragraph, " ", bold, italic, false, url);
                    }

                    break;
                case CodeInline code:
                    AddRun(paragraph, code.Content, bold, italic, true, url);
                    break;
                case LiteralInline literal:
                    AddRun(paragraph, literal.Content.ToString(), bold, italic, false, url);
                    break;
                case HtmlInline html:
                    AddRun(paragraph, html.Tag, bold, italic, false, url);
                    break;
                case HtmlEntityInline entity:
                    AddRun(paragraph, entity.Transcoded.ToString(), bold, italic, false, url);
                    break;
                case ContainerInline container:
                    this.WriteInlines(paragraph, container, bold, italic, url);
                    break;
            }
        }
    }

    private void AddRun(A.Paragraph paragraph, string text, bool bold, bool italic, bool code, string? url)
    {
        var properties = new A.RunProperties { Bold = bold, Italic = italic };
        if (code)
        {
            properties.Append(new A.LatinFont { Typeface = "Consolas" });
        }

        if (url is not null && Uri.TryCreate(url, UriKind.RelativeOrAbsolute, out var uri))
        {
            var relationship = slidePart.AddHyperlinkRelationship(uri, true);
            properties.Append(new A.HyperlinkOnClick { Id = relationship.Id });
        }

        paragraph.Append(new A.Run(properties, new A.Text(text)));
    }
}
