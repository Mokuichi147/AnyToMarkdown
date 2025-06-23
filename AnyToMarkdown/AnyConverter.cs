namespace AnyToMarkdown;

public static class AnyConverter
{
    public static ConvertResult Convert(string filePath)
    {
        string fileExtension = Path.GetExtension(filePath);
        var converter = ConverterFactory.GetConverter(fileExtension);
        
        using var stream = File.OpenRead(filePath);
        return converter.Convert(stream);
    }
    
    public static ConvertResult Convert(Stream stream, string fileExtension)
    {
        var converter = ConverterFactory.GetConverter(fileExtension);
        return converter.Convert(stream);
    }
    
    public static string[] GetSupportedExtensions()
    {
        return ConverterFactory.GetSupportedExtensions();
    }
}