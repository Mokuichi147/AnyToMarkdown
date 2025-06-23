using System.Text;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using AnyToMarkdown.Pdf;

namespace AnyToMarkdown;

public class PdfConverter : ConverterBase
{
    private const double DefaultVerticalTolerance = 5.0;
    private const double DefaultHorizontalTolerance = 5.0;
    
    public override string[] SupportedExtensions => ["pdf"];

    protected override ConvertResult ConvertInternal(Stream stream)
    {
        var sb = new StringBuilder();
        var result = new ConvertResult();
        var tableDetector = new PdfTableDetector(DefaultHorizontalTolerance, DefaultVerticalTolerance);
        
        using var document = PdfDocument.Open(stream);
        
        foreach (Page page in document.GetPages())
        {
            ProcessPage(page, sb, result, tableDetector);
        }
        
        result.Text = sb.ToString().Trim();
        LogDebugInfo(result);
        
        return result;
    }
    
    private void ProcessPage(Page page, StringBuilder sb, ConvertResult result, PdfTableDetector tableDetector)
    {
        var words = page.GetWords()
            .OrderByDescending(x => x.BoundingBox.Bottom)
            .ThenBy(x => x.BoundingBox.Left)
            .ToList();

        var lines = PdfWordProcessor.GroupWordsIntoLines(words, DefaultVerticalTolerance);
        var images = page.GetImages()
            .OrderByDescending(x => x.Bounds.Top)
            .ThenBy(x => x.Bounds.Left)
            .ToList();

        ProcessLines(lines, images, sb, result, tableDetector);
        ProcessRemainingImages(images, sb, result);
    }
    
    private void ProcessLines(List<List<Word>> lines, List<UglyToad.PdfPig.Content.IPdfImage> images, 
        StringBuilder sb, ConvertResult result, PdfTableDetector tableDetector)
    {
        int currentLineIndex = 0;
        while (currentLineIndex < lines.Count)
        {
            var tableData = tableDetector.TryDetectTable(lines, currentLineIndex);
            
            if (tableData.IsTable)
            {
                sb.AppendLine(PdfTableFormatter.ConvertToMarkdownTable(tableData.TableLines));
                currentLineIndex += tableData.RowCount;
            }
            else
            {
                ProcessSingleLine(lines[currentLineIndex], images, sb, result);
                currentLineIndex++;
            }
        }
    }
    
    private void ProcessSingleLine(List<Word> line, List<UglyToad.PdfPig.Content.IPdfImage> images, 
        StringBuilder sb, ConvertResult result)
    {
        var mergedWords = PdfWordProcessor.MergeWordsInLine(line, DefaultHorizontalTolerance);
        
        var lineImages = images.Where(x => x.Bounds.Bottom > line[0].BoundingBox.Bottom).ToList();
        foreach (var img in lineImages)
        {
            string base64 = PdfImageProcessor.ConvertImageToBase64(img.RawBytes);
            if (!string.IsNullOrEmpty(base64))
            {
                sb.AppendLine($"![]({base64})\n");
            }
            else
            {
                result.Warnings.Add("Error processing image: Unable to convert image to base64");
            }
            
            images.Remove(img);
        }

        var lineText = string.Join(" ", mergedWords.Select(x => string.Join("", x.Select(w => w.Text))));
        sb.AppendLine(lineText + "\n");
    }
    
    private void ProcessRemainingImages(List<UglyToad.PdfPig.Content.IPdfImage> images, 
        StringBuilder sb, ConvertResult result)
    {
        foreach (var img in images)
        {
            string base64 = PdfImageProcessor.ConvertImageToBase64(img.RawBytes);
            if (!string.IsNullOrEmpty(base64))
            {
                sb.AppendLine($"![]({base64})\n");
            }
            else
            {
                result.Warnings.Add("Error processing image: Unable to convert image to base64");
            }
        }
    }

}