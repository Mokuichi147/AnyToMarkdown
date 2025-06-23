namespace AnyToMarkdown;

public class DocxConverter : ConverterBase
{
    public override string[] SupportedExtensions => ["docx"];

    protected override ConvertResult ConvertInternal(Stream stream)
    {
        var result = new ConvertResult();
        
        var htmlConverter = new Mammoth.DocumentConverter();
        var conversionResult = htmlConverter.ConvertToHtml(stream);
        string htmlContent = conversionResult.Value;
        
        if (conversionResult.Warnings != null)
        {
            result.Warnings.AddRange(conversionResult.Warnings);
        }

        var converter = new ReverseMarkdown.Converter();
        result.Text = converter.Convert(htmlContent);

        LogDebugInfo(result);
        
        return result;
    }
}
