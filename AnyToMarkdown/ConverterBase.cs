namespace AnyToMarkdown;

public abstract class ConverterBase : IConverter
{
    public abstract string[] SupportedExtensions { get; }
    
    public ConvertResult Convert(Stream stream)
    {
        try
        {
            return ConvertInternal(stream);
        }
        catch (Exception ex)
        {
            return new ConvertResult
            {
                Text = string.Empty,
                Warnings = [$"Conversion failed: {ex.Message}"]
            };
        }
    }
    
    protected abstract ConvertResult ConvertInternal(Stream stream);
    
    protected void LogDebugInfo(ConvertResult result)
    {
        #if DEBUG
        Console.WriteLine($"{GetType().Name} Conversion Result:");
        Console.WriteLine(result.Text);
        
        if (result.Warnings.Count > 0)
        {
            Console.WriteLine($"{GetType().Name} Conversion Warnings:");
            foreach (var warning in result.Warnings)
            {
                Console.WriteLine(warning);
            }
        }
        #endif
    }
}