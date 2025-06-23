using UglyToad.PdfPig.Content;

namespace AnyToMarkdown.Pdf;

internal class TableDetectionResult
{
    public bool IsTable { get; set; }
    public int RowCount { get; set; }
    public List<List<List<Word>>> TableLines { get; set; } = [];
}