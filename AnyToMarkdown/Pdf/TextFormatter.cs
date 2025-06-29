using System.Linq;
using System.Text;
using UglyToad.PdfPig.Content;

namespace AnyToMarkdown.Pdf;

internal static class TextFormatter
{
    public static string RestoreFormattingWithFontInfo(string text, List<Word> words)
    {
        if (string.IsNullOrWhiteSpace(text) || words == null || words.Count == 0)
            return text;

        var result = new StringBuilder();
        var textParts = text.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        
        for (int i = 0; i < Math.Min(textParts.Length, words.Count); i++)
        {
            var word = textParts[i];
            var wordObj = words[i];
            
            // フォント名から太字・斜体を検出
            bool isBold = IsWordBold(wordObj);
            bool isItalic = IsWordItalic(wordObj);
            
            if (isBold && isItalic)
            {
                word = $"***{word}***";
            }
            else if (isBold)
            {
                word = $"**{word}**";
            }
            else if (isItalic)
            {
                word = $"*{word}*";
            }
            
            if (result.Length > 0) result.Append(" ");
            result.Append(word);
        }
        
        return result.ToString();
    }

    public static string RestoreFormatting(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return text;

        // より洗練されたフォーマット復元
        text = RestoreBoldFormatting(text);
        text = RestoreItalicFormatting(text);
        text = RestoreLinks(text);
        
        return text;
    }

    private static string RestoreBoldFormatting(string text)
    {
        // 一般的な太字パターンを検出して復元
        
        // BOLD: text → **text**
        text = System.Text.RegularExpressions.Regex.Replace(text, @"\bBOLD:\s*([^:\n]+)", "**$1**");
        
        // [BOLD] text [/BOLD] → **text**
        text = System.Text.RegularExpressions.Regex.Replace(text, @"\[BOLD\]\s*([^\[]+?)\s*\[/BOLD\]", "**$1**");
        
        // <strong> text </strong> → **text**
        text = System.Text.RegularExpressions.Regex.Replace(text, @"<strong>\s*([^<]+?)\s*</strong>", "**$1**", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        
        // <b> text </b> → **text**
        text = System.Text.RegularExpressions.Regex.Replace(text, @"<b>\s*([^<]+?)\s*</b>", "**$1**", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        
        return text;
    }

    private static string RestoreItalicFormatting(string text)
    {
        // 一般的な斜体パターンを検出して復元
        
        // ITALIC: text → *text*
        text = System.Text.RegularExpressions.Regex.Replace(text, @"\bITALIC:\s*([^:\n]+)", "*$1*");
        
        // [ITALIC] text [/ITALIC] → *text*
        text = System.Text.RegularExpressions.Regex.Replace(text, @"\[ITALIC\]\s*([^\[]+?)\s*\[/ITALIC\]", "*$1*");
        
        // <em> text </em> → *text*
        text = System.Text.RegularExpressions.Regex.Replace(text, @"<em>\s*([^<]+?)\s*</em>", "*$1*", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        
        // <i> text </i> → *text*
        text = System.Text.RegularExpressions.Regex.Replace(text, @"<i>\s*([^<]+?)\s*</i>", "*$1*", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        
        return text;
    }

    private static string RestoreLinks(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return text;

        // URL パターンの検出と Markdown リンク形式への変換
        
        // 既存の Markdown リンク形式は保持
        if (text.Contains("[") && text.Contains("]("))
        {
            return text;
        }

        // HTTP/HTTPS URL の検出
        var urlPattern = @"(https?://[^\s]+)";
        text = System.Text.RegularExpressions.Regex.Replace(text, urlPattern, "[$1]($1)");

        // Email アドレスの検出
        var emailPattern = @"\b([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})\b";
        text = System.Text.RegularExpressions.Regex.Replace(text, emailPattern, "[mailto:$1](mailto:$1)");

        // 既存のMarkdownリンク記法はそのまま保持
        return text;
    }

    private static bool IsWordBold(Word word)
    {
        var fontName = word.FontName?.ToLowerInvariant() ?? "";
        return fontName.Contains("bold") || fontName.Contains("heavy") || fontName.Contains("black");
    }
    
    private static bool IsWordItalic(Word word)
    {
        var fontName = word.FontName?.ToLowerInvariant() ?? "";
        return fontName.Contains("italic") || fontName.Contains("oblique") || fontName.Contains("slant");
    }

    public static string FixTextSpacing(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return text;

        // 複数の連続する空白を単一のスペースに統合
        text = System.Text.RegularExpressions.Regex.Replace(text, @"\s+", " ");

        // 句読点前のスペースを除去
        text = System.Text.RegularExpressions.Regex.Replace(text, @"\s+([.,:;!?。、，：；！？])", "$1");

        // 括弧前後のスペース調整
        text = System.Text.RegularExpressions.Regex.Replace(text, @"\s+([)\]}」』】])", "$1");
        text = System.Text.RegularExpressions.Regex.Replace(text, @"([(\[{「『【])\s+", "$1");

        // 行末・行頭の空白除去
        var lines = text.Split('\n');
        for (int i = 0; i < lines.Length; i++)
        {
            lines[i] = lines[i].Trim();
        }
        
        return string.Join("\n", lines);
    }

    public static string NormalizeEscapeCharacters(string text)
    {
        if (string.IsNullOrEmpty(text)) return text;
        
        // 複数のアスタリスクが連続している場合の正規化
        text = System.Text.RegularExpressions.Regex.Replace(text, @"\*{3,}", "***");
        
        // Markdownエスケープ文字の汎用検出と処理
        
        // 不正なエスケープシーケンスの修正
        text = text.Replace("\\\\*", "\\*");  // \\* → \*
        text = text.Replace("\\\\_", "\\_");  // \\_ → \_
        text = text.Replace("\\\\#", "\\#");  // \\# → \#
        
        // 重複したアンダースコアの正規化
        text = System.Text.RegularExpressions.Regex.Replace(text, @"_{3,}", "__");
        
        // 不要なバックスラッシュの除去（ただし、有効なエスケープは保持）
        text = System.Text.RegularExpressions.Regex.Replace(text, @"\\([^*_#\[\]()\\])", "$1");
        
        return text;
    }

    public static bool ContainsMarkdownSyntax(string text)
    {
        // Markdownの特殊記法を検出
        return text.Contains("**") || text.Contains("__") || 
               text.Contains("```") || text.Contains("`") ||
               text.Contains("[") && text.Contains("](") ||
               text.StartsWith("#") || text.StartsWith(">") ||
               text.StartsWith("-") || text.StartsWith("*") || text.StartsWith("+") ||
               text.Contains("---") || text.Contains("***") || text.Contains("___");
    }

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

    public static double CalculateWordDensity(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return 0.0;
        
        var words = text.Split(new[] { ' ', '\t', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
        var characters = text.Length;
        
        return characters > 0 ? (double)words.Length / characters : 0.0;
    }

    public static double CalculateVariance(IEnumerable<double> values)
    {
        var enumerable = values.ToList();
        if (!enumerable.Any()) return 0.0;
        
        var average = enumerable.Average();
        var variance = enumerable.Sum(x => Math.Pow(x - average, 2)) / enumerable.Count();
        
        return variance;
    }
}