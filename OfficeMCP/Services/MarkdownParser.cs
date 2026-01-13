using System.Text.RegularExpressions;

namespace OfficeMCP.Services;

/// <summary>
/// Represents a parsed Markdown element.
/// </summary>
public abstract record MarkdownElement;

public sealed record MarkdownHeading(int Level, string Text, List<MarkdownInline> Inlines) : MarkdownElement;
public sealed record MarkdownParagraph(List<MarkdownInline> Inlines) : MarkdownElement;
public sealed record MarkdownBulletList(List<MarkdownListItem> Items) : MarkdownElement;
public sealed record MarkdownNumberedList(List<MarkdownListItem> Items) : MarkdownElement;
public sealed record MarkdownListItem(List<MarkdownInline> Inlines);
public sealed record MarkdownCodeBlock(string Code, string? Language) : MarkdownElement;
public sealed record MarkdownBlockquote(List<MarkdownInline> Inlines) : MarkdownElement;
public sealed record MarkdownHorizontalRule : MarkdownElement;
public sealed record MarkdownTable(List<string> Headers, List<List<string>> Rows) : MarkdownElement;
public sealed record MarkdownImage(string AltText, string Url) : MarkdownElement;

/// <summary>
/// Represents inline formatting within text.
/// </summary>
public abstract record MarkdownInline;
public sealed record MarkdownText(string Text) : MarkdownInline;
public sealed record MarkdownBold(string Text) : MarkdownInline;
public sealed record MarkdownItalic(string Text) : MarkdownInline;
public sealed record MarkdownBoldItalic(string Text) : MarkdownInline;
public sealed record MarkdownCode(string Text) : MarkdownInline;
public sealed record MarkdownStrikethrough(string Text) : MarkdownInline;
public sealed record MarkdownLink(string Text, string Url) : MarkdownInline;
public sealed record MarkdownInlineImage(string AltText, string Url) : MarkdownInline;

/// <summary>
/// Parser for converting Markdown text into structured elements.
/// </summary>
public static partial class MarkdownParser
{
    public static List<MarkdownElement> Parse(string markdown)
    {
        var elements = new List<MarkdownElement>();
        var lines = markdown.Replace("\r\n", "\n").Split('\n');
        var i = 0;

        while (i < lines.Length)
        {
            var line = lines[i];
            var trimmedLine = line.TrimStart();

            // Empty line
            if (string.IsNullOrWhiteSpace(line))
            {
                i++;
                continue;
            }

            // Horizontal rule
            if (HorizontalRuleRegex().IsMatch(trimmedLine))
            {
                elements.Add(new MarkdownHorizontalRule());
                i++;
                continue;
            }

            // Heading
            var headingMatch = HeadingRegex().Match(trimmedLine);
            if (headingMatch.Success)
            {
                var level = headingMatch.Groups[1].Value.Length;
                var text = headingMatch.Groups[2].Value.Trim();
                elements.Add(new MarkdownHeading(level, text, ParseInlines(text)));
                i++;
                continue;
            }

            // Code block
            if (trimmedLine.StartsWith("```"))
            {
                var language = trimmedLine.Length > 3 ? trimmedLine[3..].Trim() : null;
                var codeLines = new List<string>();
                i++;
                while (i < lines.Length && !lines[i].TrimStart().StartsWith("```"))
                {
                    codeLines.Add(lines[i]);
                    i++;
                }
                i++; // Skip closing ```
                elements.Add(new MarkdownCodeBlock(string.Join(Environment.NewLine, codeLines), language));
                continue;
            }

            // Table
            if (trimmedLine.Contains('|') && i + 1 < lines.Length && TableSeparatorRegex().IsMatch(lines[i + 1]))
            {
                var (table, newIndex) = ParseTable(lines, i);
                if (table != null)
                {
                    elements.Add(table);
                    i = newIndex;
                    continue;
                }
            }

            // Blockquote
            if (trimmedLine.StartsWith('>'))
            {
                var quoteLines = new List<string>();
                while (i < lines.Length && lines[i].TrimStart().StartsWith('>'))
                {
                    quoteLines.Add(lines[i].TrimStart().TrimStart('>').TrimStart());
                    i++;
                }
                var quoteText = string.Join(" ", quoteLines);
                elements.Add(new MarkdownBlockquote(ParseInlines(quoteText)));
                continue;
            }

            // Bullet list
            var bulletMatch = BulletListItemRegex().Match(line);
            if (bulletMatch.Success)
            {
                var items = new List<MarkdownListItem>();
                while (i < lines.Length)
                {
                    var itemMatch = BulletListItemRegex().Match(lines[i]);
                    if (!itemMatch.Success) break;
                    items.Add(new MarkdownListItem(ParseInlines(itemMatch.Groups[1].Value)));
                    i++;
                }
                elements.Add(new MarkdownBulletList(items));
                continue;
            }

            // Numbered list
            var numberedMatch = NumberedListItemRegex().Match(line);
            if (numberedMatch.Success)
            {
                var items = new List<MarkdownListItem>();
                while (i < lines.Length)
                {
                    var itemMatch = NumberedListItemRegex().Match(lines[i]);
                    if (!itemMatch.Success) break;
                    items.Add(new MarkdownListItem(ParseInlines(itemMatch.Groups[1].Value)));
                    i++;
                }
                elements.Add(new MarkdownNumberedList(items));
                continue;
            }

            // Image on its own line
            var imageMatch = ImageRegex().Match(trimmedLine);
            if (imageMatch.Success && imageMatch.Value == trimmedLine)
            {
                elements.Add(new MarkdownImage(imageMatch.Groups[1].Value, imageMatch.Groups[2].Value));
                i++;
                continue;
            }

            // Regular paragraph - collect continuous lines
            var paragraphLines = new List<string>();
            while (i < lines.Length)
            {
                var currentLine = lines[i];
                if (string.IsNullOrWhiteSpace(currentLine) ||
                    HeadingRegex().IsMatch(currentLine.TrimStart()) ||
                    currentLine.TrimStart().StartsWith("```") ||
                    currentLine.TrimStart().StartsWith('>') ||
                    BulletListItemRegex().IsMatch(currentLine) ||
                    NumberedListItemRegex().IsMatch(currentLine) ||
                    HorizontalRuleRegex().IsMatch(currentLine.TrimStart()))
                {
                    break;
                }
                paragraphLines.Add(currentLine);
                i++;
            }

            if (paragraphLines.Count > 0)
            {
                var paragraphText = string.Join(" ", paragraphLines);
                elements.Add(new MarkdownParagraph(ParseInlines(paragraphText)));
            }
        }

        return elements;
    }

    public static List<MarkdownInline> ParseInlines(string text)
    {
        var inlines = new List<MarkdownInline>();
        var remaining = text;

        while (!string.IsNullOrEmpty(remaining))
        {
            // Try to match inline patterns in order of specificity
            
            // Inline image
            var imageMatch = ImageRegex().Match(remaining);
            if (imageMatch.Success && imageMatch.Index == 0)
            {
                inlines.Add(new MarkdownInlineImage(imageMatch.Groups[1].Value, imageMatch.Groups[2].Value));
                remaining = remaining[imageMatch.Length..];
                continue;
            }

            // Link
            var linkMatch = LinkRegex().Match(remaining);
            if (linkMatch.Success && linkMatch.Index == 0)
            {
                inlines.Add(new MarkdownLink(linkMatch.Groups[1].Value, linkMatch.Groups[2].Value));
                remaining = remaining[linkMatch.Length..];
                continue;
            }

            // Inline code
            var codeMatch = InlineCodeRegex().Match(remaining);
            if (codeMatch.Success && codeMatch.Index == 0)
            {
                inlines.Add(new MarkdownCode(codeMatch.Groups[1].Value));
                remaining = remaining[codeMatch.Length..];
                continue;
            }

            // Bold + Italic (***text*** or ___text___)
            var boldItalicMatch = BoldItalicRegex().Match(remaining);
            if (boldItalicMatch.Success && boldItalicMatch.Index == 0)
            {
                inlines.Add(new MarkdownBoldItalic(boldItalicMatch.Groups[1].Value));
                remaining = remaining[boldItalicMatch.Length..];
                continue;
            }

            // Bold (**text** or __text__)
            var boldMatch = BoldRegex().Match(remaining);
            if (boldMatch.Success && boldMatch.Index == 0)
            {
                inlines.Add(new MarkdownBold(boldMatch.Groups[1].Value));
                remaining = remaining[boldMatch.Length..];
                continue;
            }

            // Italic (*text* or _text_)
            var italicMatch = ItalicRegex().Match(remaining);
            if (italicMatch.Success && italicMatch.Index == 0)
            {
                inlines.Add(new MarkdownItalic(italicMatch.Groups[1].Value));
                remaining = remaining[italicMatch.Length..];
                continue;
            }

            // Strikethrough (~~text~~)
            var strikeMatch = StrikethroughRegex().Match(remaining);
            if (strikeMatch.Success && strikeMatch.Index == 0)
            {
                inlines.Add(new MarkdownStrikethrough(strikeMatch.Groups[1].Value));
                remaining = remaining[strikeMatch.Length..];
                continue;
            }

            // Find the next special character
            var nextSpecialIndex = FindNextSpecialCharacter(remaining);
            if (nextSpecialIndex > 0)
            {
                inlines.Add(new MarkdownText(remaining[..nextSpecialIndex]));
                remaining = remaining[nextSpecialIndex..];
            }
            else if (nextSpecialIndex == 0)
            {
                // Special character that didn't match a pattern, treat as text
                inlines.Add(new MarkdownText(remaining[..1]));
                remaining = remaining[1..];
            }
            else
            {
                // No more special characters, add remaining text
                inlines.Add(new MarkdownText(remaining));
                break;
            }
        }

        return inlines;
    }

    private static int FindNextSpecialCharacter(string text)
    {
        var specialChars = new[] { '*', '_', '`', '~', '[', '!' };
        var minIndex = -1;

        foreach (var c in specialChars)
        {
            var index = text.IndexOf(c);
            if (index >= 0 && (minIndex == -1 || index < minIndex))
            {
                minIndex = index;
            }
        }

        return minIndex;
    }

    private static (MarkdownTable?, int) ParseTable(string[] lines, int startIndex)
    {
        var headerLine = lines[startIndex].Trim();
        var headers = ParseTableRow(headerLine);

        if (headers.Count == 0) return (null, startIndex + 1);

        var i = startIndex + 2; // Skip header and separator
        var rows = new List<List<string>>();

        while (i < lines.Length)
        {
            var line = lines[i].Trim();
            if (!line.Contains('|')) break;
            
            var row = ParseTableRow(line);
            if (row.Count > 0)
            {
                // Pad row to match header count
                while (row.Count < headers.Count)
                    row.Add(string.Empty);
                rows.Add(row);
            }
            i++;
        }

        return (new MarkdownTable(headers, rows), i);
    }

    private static List<string> ParseTableRow(string line)
    {
        var cells = line.Split('|')
            .Select(c => c.Trim())
            .Where(c => !string.IsNullOrEmpty(c) || line.StartsWith("|") || line.EndsWith("|"))
            .ToList();

        // Remove empty cells from start/end caused by leading/trailing pipes
        if (cells.Count > 0 && string.IsNullOrEmpty(cells[0]) && line.StartsWith("|"))
            cells.RemoveAt(0);
        if (cells.Count > 0 && string.IsNullOrEmpty(cells[^1]) && line.EndsWith("|"))
            cells.RemoveAt(cells.Count - 1);

        return cells;
    }

    [GeneratedRegex(@"^(#{1,6})\s+(.+)$")]
    private static partial Regex HeadingRegex();

    [GeneratedRegex(@"^(\*{3,}|-{3,}|_{3,})$")]
    private static partial Regex HorizontalRuleRegex();

    [GeneratedRegex(@"^\s*[-*+]\s+(.+)$")]
    private static partial Regex BulletListItemRegex();

    [GeneratedRegex(@"^\s*\d+\.\s+(.+)$")]
    private static partial Regex NumberedListItemRegex();

    [GeneratedRegex(@"^\s*\|?\s*[-:]+[-|\s:]+\s*\|?\s*$")]
    private static partial Regex TableSeparatorRegex();

    [GeneratedRegex(@"\*\*\*(.+?)\*\*\*|___(.+?)___")]
    private static partial Regex BoldItalicRegex();

    [GeneratedRegex(@"\*\*(.+?)\*\*|__(.+?)__")]
    private static partial Regex BoldRegex();

    [GeneratedRegex(@"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)|(?<!_)_(?!_)(.+?)(?<!_)_(?!_)")]
    private static partial Regex ItalicRegex();

    [GeneratedRegex(@"~~(.+?)~~")]
    private static partial Regex StrikethroughRegex();

    [GeneratedRegex(@"`([^`]+)`")]
    private static partial Regex InlineCodeRegex();

    [GeneratedRegex(@"\[([^\]]+)\]\(([^)]+)\)")]
    private static partial Regex LinkRegex();

    [GeneratedRegex(@"!\[([^\]]*)\]\(([^)]+)\)")]
    private static partial Regex ImageRegex();
}
