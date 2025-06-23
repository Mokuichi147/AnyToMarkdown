namespace AnyToMarkdown.Tests;

public class PdfTest
{
    static string filePath = "./Resources/sample.pdf";

    [Fact]
    public void ConvertTest()
    {
        using FileStream stream = File.OpenRead(filePath);
        var converter = new PdfConverter();
        ConvertResult result = converter.Convert(stream);
        Assert.NotNull(result.Text);
        //Assert.Empty(result.Warnings);

        // ファイルに変換結果を保存する
        //File.WriteAllText("sample-pdf.md", result.Text);
    }
}
