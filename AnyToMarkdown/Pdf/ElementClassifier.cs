using UglyToad.PdfPig.Content;

namespace AnyToMarkdown.Pdf;

public enum ElementType
{
    Empty,
    Header,
    Paragraph,
    ListItem,
    TableRow,
    CodeBlock,
    QuoteBlock,
    HorizontalLine
}

public class DocumentElement
{
    public ElementType Type { get; set; }
    public string Content { get; set; } = "";
    public double FontSize { get; set; }
    public double LeftMargin { get; set; }
    public bool IsIndented { get; set; }
    public List<Word> Words { get; set; } = new();
}

public class DocumentStructure
{
    public List<DocumentElement> Elements { get; set; } = new();
    public FontAnalysis FontAnalysis { get; set; } = new();
}