namespace AnyToMarkdown;

public interface IConverter
{
    string[] SupportedExtensions { get; }
    ConvertResult Convert(Stream stream);
}