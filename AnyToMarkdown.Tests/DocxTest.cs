namespace AnyToMarkdown.Tests;

public class DocxTest
{
    static string filePath = "./Resources/sample.docx";

    [Fact]
    public void ConvertTest()
    {
        using FileStream stream = File.OpenRead(filePath);
        ConvertResult result = DocxConverter.Convert(stream);
        Assert.NotNull(result.Text);
        //Assert.Empty(result.Warnings);

        // ファイルに変換結果を保存する
        //File.WriteAllText("sample-docx.md", result.Text);
    }
}
