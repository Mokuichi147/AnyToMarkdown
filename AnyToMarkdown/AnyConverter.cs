namespace AnyToMarkdown;

public static class AnyConverter
{
    public static ConvertResult Convert(string filePath)
    {
        // ファイルの拡張子を取得
        string fileType = Path.GetExtension(filePath).TrimStart('.').ToLower();
        
        // ファイルストリームを開く
        using FileStream stream = File.OpenRead(filePath);
        
        // 変換処理を実行
        if (fileType == "docx")
        {
            return DocxConverter.Convert(stream);
        }
        else
        {
            throw new NotSupportedException($"Unsupported file type: {fileType}");
        }
    }
}