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
    
    [Fact]
    public void ConvertComplexTableTest()
    {
        string complexTablePath = "./Resources/test-complex-table.pdf";
        
        if (!File.Exists(complexTablePath))
        {
            Assert.Fail($"Complex table PDF not found: {complexTablePath}");
        }
        
        using FileStream stream = File.OpenRead(complexTablePath);
        var converter = new PdfConverter();
        ConvertResult result = converter.Convert(stream);
        
        Assert.NotNull(result.Text);
        
        // 結果をファイルに保存して検証
        File.WriteAllText("test-output.md", result.Text);
        
        // Basic & Standard $50.00 が正しく分離されているかチェック
        Assert.Contains("Basic & Standard", result.Text);
        Assert.Contains("$50.00", result.Text);
        
        // 期待される結果：Basic & Standard と $50.00 が隣接する別々のセルになっている
        var lines = result.Text.Split('\n');
        var tableRows = lines.Where(line => line.Contains("Basic & Standard")).ToList();
        
        Assert.Single(tableRows); // Basic & Standard を含む行が1つだけ
        var basicRow = tableRows[0];
        
        // 行を分割してセルを確認
        var cells = basicRow.Split('|').Select(c => c.Trim()).Where(c => !string.IsNullOrEmpty(c)).ToList();
        
        // デバッグ情報をアサーションに含める
        var debugInfo = $"Row: {basicRow}\nCells: [{string.Join(", ", cells.Select(c => $"'{c}'"))}]";
        
        // Basic & Standard が一つのセルに、$50.00 が別のセルにあることを確認
        bool hasBasicCell = cells.Any(c => c.Contains("Basic & Standard") && !c.Contains("$50.00"));
        bool hasPriceCell = cells.Any(c => c.Contains("$50.00") && !c.Contains("Basic & Standard"));
        
        Assert.True(hasBasicCell && hasPriceCell, 
            $"Basic & Standard and $50.00 should be in separate cells.\n{debugInfo}");
    }
}
