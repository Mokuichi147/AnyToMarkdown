namespace AnyToMarkdown.Tests;

public class PdfTest
{
    static string filePath = "./Resources/sample.pdf";

    [Fact]
    public void ConvertTest()
    {
        using FileStream stream = File.OpenRead(filePath);
        ConvertResult result = PdfConverter.Convert(stream);
        Assert.NotNull(result.Text);
        //Assert.Empty(result.Warnings);

        // ファイルに変換結果を保存する
        //File.WriteAllText("sample-docx.md", result.Text);
    }
}
