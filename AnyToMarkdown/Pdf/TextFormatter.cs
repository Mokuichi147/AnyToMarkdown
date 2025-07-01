namespace AnyToMarkdown.Pdf;

internal static class TextFormatter
{
    public static bool IsJapaneseText(string text)
    {
        return text.Any(c => (c >= 0x3040 && c <= 0x309F) || // ひらがな
                           (c >= 0x30A0 && c <= 0x30FF) || // カタカナ
                           (c >= 0x4E00 && c <= 0x9FAF));  // 漢字
    }

    public static string DetermineParagraphSeparator(string currentText, string nextText)
    {
        // 日本語テキストの場合はスペースなし、英語はスペース追加
        if (IsJapaneseText(currentText) && IsJapaneseText(nextText))
        {
            return "";
        }
        return " ";
    }
}