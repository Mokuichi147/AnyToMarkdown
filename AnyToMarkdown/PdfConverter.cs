using System.Text;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace AnyToMarkdown;

public static class PdfConverter
{
    public static ConvertResult Convert(FileStream stream)
    {
        StringBuilder sb = new();
        using var document = PdfDocument.Open(stream);
        foreach (Page page in document.GetPages())
        {
            sb.AppendLine(page.Text);
            sb.AppendLine();
        }

        string markdown = sb.ToString();

        #if DEBUG
        Console.WriteLine("Conversion Result:");
        Console.WriteLine(markdown);
        #endif

        return new ConvertResult
        {
            Text = markdown,
            Warnings = [],
        };
    }
}