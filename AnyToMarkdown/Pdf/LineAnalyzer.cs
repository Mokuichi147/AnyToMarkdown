using System.Linq;
using System.Text;
using UglyToad.PdfPig.Content;

namespace AnyToMarkdown.Pdf;

internal static class LineAnalyzer
{
    public static DocumentElement AnalyzeLine(List<Word> line, FontAnalysis fontAnalysis, double horizontalTolerance)
    {
        if (line.Count == 0)
            return new DocumentElement { Type = ElementType.Empty, Content = "" };

        var mergedWords = PdfWordProcessor.MergeWordsInLine(line, horizontalTolerance);
        
        // フォント検出機能付きのテキスト生成（完全な実装）
        var formattedText = BuildFormattedText(mergedWords);

        if (string.IsNullOrWhiteSpace(formattedText))
            return new DocumentElement { Type = ElementType.Empty, Content = "" };

        // フォントサイズ分析
        var avgFontSize = line.Average(w => w.BoundingBox.Height);
        var maxFontSize = line.Max(w => w.BoundingBox.Height);
        
        // 位置分析
        var leftMargin = line.Min(w => w.BoundingBox.Left);
        var isIndented = leftMargin > 50; // 基準マージンから50pt以上右

        // 要素タイプの判定
        var elementType = DetermineElementType(formattedText, avgFontSize, maxFontSize, fontAnalysis, isIndented, line);

        return new DocumentElement
        {
            Type = elementType,
            Content = formattedText,
            FontSize = avgFontSize,
            LeftMargin = leftMargin,
            IsIndented = isIndented,
            Words = line
        };
    }

    private static ElementType DetermineElementType(string text, double avgFontSize, double maxFontSize, 
        FontAnalysis fontAnalysis, bool isIndented, List<Word> words)
    {
        // フォーマット済みテキストから元のテキストを抽出
        var cleanText = ExtractCleanTextForAnalysis(text);
        
        // 明確なMarkdownパターン
        if (cleanText.StartsWith("#")) return ElementType.Header;
        if (cleanText.StartsWith(">")) return ElementType.QuoteBlock;
        if (cleanText.StartsWith("```")) return ElementType.CodeBlock;
        
        // コードコメントやプログラミング構文の検出（コードブロック候補）
        if (cleanText.StartsWith("//") || cleanText.StartsWith("#") && (cleanText.Contains("Python") || cleanText.Contains("コード")))
        {
            return ElementType.CodeBlock;
        }
        
        // 空行チェック（完全に空または空白のみ）
        if (string.IsNullOrWhiteSpace(cleanText))
        {
            return ElementType.Empty;
        }
        
        // 単一文字の場合は段落として扱う
        if (cleanText.Trim().Length == 1)
        {
            return ElementType.Paragraph;
        }
        
        // 水平線パターン
        if (IsHorizontalLinePattern(cleanText))
        {
            return ElementType.HorizontalLine;
        }
        
        // 検出順序を慎重に調整（より具体的なパターンから順に）
        
        // 1. ヘッダー構造の検出（より慎重に）
        bool hasHeaderContent = ElementDetector.IsHeaderLike(cleanText);
        
        // 段落的なパターンを除外
        bool isParagraphLike = cleanText.EndsWith("。") || cleanText.EndsWith(".") || 
                              cleanText.Contains("、") || cleanText.Contains(",") ||
                              cleanText.Contains("**") || (cleanText.Contains("*") && !cleanText.StartsWith("*")) ||
                              cleanText.StartsWith("これは") || cleanText.StartsWith("それは") ||
                              cleanText.StartsWith("•") || cleanText.StartsWith("・") || cleanText.StartsWith("-");
        
        // 明らかに大きなフォントで段落的でない場合のみヘッダー
        var fontSizeRatio = maxFontSize / fontAnalysis.BaseFontSize;
        if (fontSizeRatio >= 1.6 && !isParagraphLike && 
            !ElementDetector.IsLikelyTableContent(cleanText, words) && cleanText.Length <= 50)
        {
            return ElementType.Header;
        }
        
        // 2. コードブロック検出（独特なパターンを持つ）
        if (ElementDetector.IsCodeBlockLike(cleanText, words, fontAnalysis)) return ElementType.CodeBlock;
        
        // 3. 引用ブロック検出
        if (ElementDetector.IsQuoteBlockLike(cleanText)) return ElementType.QuoteBlock;
        
        // 4. リストアイテム検出
        if (ElementDetector.IsListItemLike(cleanText)) return ElementType.ListItem;
        
        // 5. テーブル行検出（座標とパターン両方を使用）
        if (ElementDetector.IsTableRowLike(cleanText, words)) return ElementType.TableRow;
        
        // 6. テーブルコンテンツの詳細検出
        bool likelyTableContent = ElementDetector.IsLikelyTableContent(cleanText, words);
        if (likelyTableContent) return ElementType.TableRow;
        
        // 7. フォントサイズによる二次ヘッダー検出（より厳格に）
        var secondaryFontRatio = avgFontSize / fontAnalysis.BaseFontSize;
        
        // フォーマット記号を含むテキストはヘッダーではない
        if (cleanText.Contains("**") || (cleanText.Contains("*") && !cleanText.StartsWith("*")))
        {
            // 段落として分類
        }
        // リスト記号で始まるテキストはヘッダーではない
        else if (cleanText.StartsWith("•") || cleanText.StartsWith("・") || cleanText.StartsWith("-"))
        {
            // 段落として分類
        }
        // 明らかに大きなフォントでヘッダーらしいパターンのみ
        else if (secondaryFontRatio >= 1.5 && ElementDetector.IsHeaderLike(cleanText) && 
                !cleanText.EndsWith("。") && !cleanText.EndsWith(".") && cleanText.Length <= 50)
        {
            return ElementType.Header;
        }
        
        // 8. ヘッダーコンテンツパターンによる検出（より厳格に）
        if (hasHeaderContent && cleanText.Length <= 50 && secondaryFontRatio >= 1.3 &&
            !cleanText.EndsWith("。") && !cleanText.EndsWith(".") && !cleanText.Contains("、"))
        {
            return ElementType.Header;
        }
        
        // 9. インデントベースのリスト検出
        if (isIndented && ElementDetector.IsListItemLike(cleanText))
        {
            return ElementType.ListItem;
        }
        
        // デフォルトは段落
        return ElementType.Paragraph;
    }

    private static string ExtractCleanTextForAnalysis(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return "";

        // Markdownフォーマットを一時的に除去してコンテンツを分析
        var cleanText = text;
        
        // 太字・斜体記法を除去
        cleanText = System.Text.RegularExpressions.Regex.Replace(cleanText, @"\*\*(.+?)\*\*", "$1");
        cleanText = System.Text.RegularExpressions.Regex.Replace(cleanText, @"\*(.+?)\*", "$1");
        cleanText = System.Text.RegularExpressions.Regex.Replace(cleanText, @"__(.+?)__", "$1");
        cleanText = System.Text.RegularExpressions.Regex.Replace(cleanText, @"_(.+?)_", "$1");
        
        // その他の記法を除去
        cleanText = System.Text.RegularExpressions.Regex.Replace(cleanText, @"`(.+?)`", "$1");
        cleanText = System.Text.RegularExpressions.Regex.Replace(cleanText, @"~~(.+?)~~", "$1");
        
        return cleanText.Trim();
    }

    private static bool IsHorizontalLinePattern(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return false;

        var cleanText = text.Trim();
        
        // 典型的な水平線パターン
        if (cleanText == "---" || cleanText == "***" || cleanText == "___")
            return true;

        // 長い線パターン
        if (cleanText.Length >= 3)
        {
            var chars = cleanText.ToCharArray();
            var firstChar = chars[0];
            
            if ((firstChar == '-' || firstChar == '*' || firstChar == '_') &&
                chars.All(c => c == firstChar))
                return true;
        }

        return false;
    }

    private static string BuildFormattedText(List<List<Word>> mergedWords)
    {
        if (mergedWords == null || mergedWords.Count == 0)
            return "";

        var result = new StringBuilder();
        
        for (int i = 0; i < mergedWords.Count; i++)
        {
            var wordGroup = mergedWords[i];
            var groupResult = BuildFormattedWordGroup(wordGroup);
            
            if (!string.IsNullOrEmpty(groupResult))
            {
                if (result.Length > 0)
                    result.Append(" ");
                result.Append(groupResult);
            }
        }
        
        return result.ToString();
    }

    private static string BuildFormattedWordGroup(List<Word> wordGroup)
    {
        if (wordGroup == null || wordGroup.Count == 0)
            return "";

        var result = new StringBuilder();
        FontFormatting? currentFormatting = null;
        
        foreach (var word in wordGroup)
        {
            var wordText = word.Text ?? "";
            var wordFormatting = FontAnalyzer.AnalyzeFontFormatting(new List<Word> { word });
            
            // フォーマットの変更を検出
            if (currentFormatting != null && !FormattingEqual(currentFormatting, wordFormatting))
            {
                // 前のフォーマットを終了
                result.Append(GetFormattingCloseTag(currentFormatting));
            }
            
            // 新しいフォーマットを開始
            if (currentFormatting == null || !FormattingEqual(currentFormatting, wordFormatting))
            {
                result.Append(GetFormattingOpenTag(wordFormatting));
                currentFormatting = wordFormatting;
            }
            
            result.Append(wordText);
        }
        
        // 最後のフォーマットを終了
        if (currentFormatting != null)
        {
            result.Append(GetFormattingCloseTag(currentFormatting));
        }
        
        return result.ToString();
    }

    private static bool FormattingEqual(FontFormatting a, FontFormatting b)
    {
        return a.IsBold == b.IsBold && a.IsItalic == b.IsItalic;
    }

    private static string BuildFormattedTextSimple(List<List<Word>> mergedWords)
    {
        if (mergedWords == null || mergedWords.Count == 0)
            return "";

        var textParts = mergedWords
            .Where(group => group != null && group.Count > 0)
            .Select(group => string.Join("", group.Select(w => w.Text ?? "")))
            .Where(text => !string.IsNullOrEmpty(text));

        return string.Join(" ", textParts);
    }

    private static string GetFormattingOpenTag(FontFormatting formatting)
    {
        var result = "";
        
        if (formatting.IsBold && formatting.IsItalic)
            result = "***";
        else if (formatting.IsBold)
            result = "**";
        else if (formatting.IsItalic)
            result = "*";
        
        return result;
    }

    private static string GetFormattingCloseTag(FontFormatting formatting)
    {
        var result = "";
        
        if (formatting.IsBold && formatting.IsItalic)
            result = "***";
        else if (formatting.IsBold)
            result = "**";
        else if (formatting.IsItalic)
            result = "*";
        
        return result;
    }
}