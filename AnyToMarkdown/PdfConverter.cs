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
        var result = new ConvertResult();
        var allMarkdown = new StringBuilder();
        
        using var document = PdfDocument.Open(stream);
        
        foreach (Page page in document.GetPages())
        {
            // 新しい構造解析アプローチを使用
            var structure = PdfStructureAnalyzer.AnalyzePageStructure(page, DefaultHorizontalTolerance, DefaultVerticalTolerance);
            var pageMarkdown = MarkdownGenerator.GenerateMarkdown(structure);
            
            if (!string.IsNullOrWhiteSpace(pageMarkdown))
            {
                allMarkdown.AppendLine(pageMarkdown);
                if (document.NumberOfPages > 1)
                {
                    allMarkdown.AppendLine(); // ページ間の区切り
                }
            }
        }
        
        result.Text = allMarkdown.ToString().Trim();
        LogDebugInfo(result);
        
        return result;
    }

}