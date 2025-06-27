using System;
using System.IO;
using AnyToMarkdown;

class DebugConversion
{
    static void Main()
    {
        var pdfPath = "/Users/mokuichi147/github/AnyToMarkdown/AnyToMarkdown.Tests/Resources/test-comprehensive-markdown.pdf";
        
        Console.WriteLine("Converting PDF to Markdown...");
        var result = AnyConverter.Convert(pdfPath);
        
        Console.WriteLine("Conversion successful: " + result.IsSuccess);
        
        if (result.IsSuccess)
        {
            var lines = result.Content.Split('\n');
            Console.WriteLine($"Total lines: {lines.Length}");
            Console.WriteLine("\n=== First 100 lines of converted markdown ===");
            for (int i = 0; i < Math.Min(100, lines.Length); i++)
            {
                Console.WriteLine($"{i+1:D3}: {lines[i]}");
            }
        }
        else
        {
            Console.WriteLine("Conversion failed:");
            Console.WriteLine(result.ErrorMessage);
        }
    }
}