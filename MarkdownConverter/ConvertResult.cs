namespace MarkdownConverter;

/// <summary>
/// 変換結果を表すクラス
/// </summary>
public class ConvertResult
{
    /// <summary>
    /// 変換結果のテキスト
    /// </summary>
    public string? Text { get; set; }

    /// <summary>
    /// 変換時の警告メッセージのリスト
    /// </summary>
    public List<string> Warnings { get; set; } = [];
}