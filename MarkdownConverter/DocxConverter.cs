namespace MarkdownConverter;

public static class DocxConverter
{
    public static ConvertResult Convert(FileStream stream)
    {
        ConvertResult result = new();
        
        // Mammothを使ってDOCXからHTMLへ変換
        var htmlConverter = new Mammoth.DocumentConverter();
        var conversionResult = htmlConverter.ConvertToHtml(stream);
        string htmlContent = conversionResult.Value;
        
        // 変換時の警告がある場合、必要に応じて表示する
        if (conversionResult.Warnings != null)
        {
            result.Warnings.AddRange(conversionResult.Warnings);
        }

        // ReverseMarkdownを使ってHTMLからMarkdownへ変換
        var converter = new ReverseMarkdown.Converter();
        result.Text = converter.Convert(htmlContent);

        #if DEBUG
        Console.WriteLine("Conversion Result:");
        Console.WriteLine(result.Text);

        Console.WriteLine("Conversion Warnings:");
        foreach (var warning in result.Warnings)
        {
            Console.WriteLine(warning);
        }
        #endif

        return result;
    }
}
