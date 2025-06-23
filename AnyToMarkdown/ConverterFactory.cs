namespace AnyToMarkdown;

public static class ConverterFactory
{
    private static readonly Dictionary<string, Func<IConverter>> _converters = new()
    {
        ["docx"] = () => new DocxConverter(),
        ["pdf"] = () => new PdfConverter()
    };
    
    public static IConverter GetConverter(string fileExtension)
    {
        var extension = fileExtension.TrimStart('.').ToLower();
        
        if (_converters.TryGetValue(extension, out var factory))
        {
            return factory();
        }
        
        throw new NotSupportedException($"Unsupported file type: {extension}");
    }
    
    public static void RegisterConverter(string extension, Func<IConverter> factory)
    {
        _converters[extension.TrimStart('.').ToLower()] = factory;
    }
    
    public static string[] GetSupportedExtensions()
    {
        return [.. _converters.Keys];
    }
}