using System.Linq;
using UglyToad.PdfPig.Content;

namespace AnyToMarkdown.Pdf;

internal static class ElementDetector
{
    public static bool IsLikelyTableContent(string cleanText, List<Word> line)
    {
        if (string.IsNullOrWhiteSpace(cleanText) || line == null || line.Count == 0)
            return false;

        // より精密な表コンテンツの検出
        
        // 数値パターンの強い検出
        if (HasNumericTablePattern(line))
            return true;

        // 規則的なスペーシングパターン
        if (HasRegularSpacing(line))
            return true;

        // テーブル的なテキストパターン
        if (HasTableLikeTextPattern(cleanText))
            return true;

        // 短い単語の規則的配置
        var words = cleanText.Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
        if (words.Length >= 3)
        {
            var avgLength = words.Average(w => w.Length);
            var allShort = words.All(w => w.Length <= 15);
            
            if (allShort && avgLength <= 8)
                return true;
        }

        // 特定の区切り文字パターン
        if (cleanText.Contains("|") || cleanText.Contains("\t") || 
            System.Text.RegularExpressions.Regex.IsMatch(cleanText, @"\s{3,}"))
            return true;

        // 数字と単語の混在パターン
        var hasNumbers = System.Text.RegularExpressions.Regex.IsMatch(cleanText, @"\d");
        var hasText = System.Text.RegularExpressions.Regex.IsMatch(cleanText, @"[a-zA-Zあ-ん]");
        if (hasNumbers && hasText && words.Length >= 2)
            return true;

        return false;
    }

    public static bool IsHeaderLike(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return false;

        var cleanText = text.Trim();
        
        // 明示的なMarkdownヘッダー
        if (cleanText.StartsWith("#"))
            return true;

        // 番号付きヘッダーパターン（階層構造）
        if (System.Text.RegularExpressions.Regex.IsMatch(cleanText, @"^(\d+\.)+\s*[^\d]"))
            return true;

        // 短いテキスト（ヘッダーの可能性）
        if (cleanText.Length <= 50)
        {
            var words = cleanText.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (words.Length <= 8)
            {
                // 全て大文字のパターン
                if (cleanText.ToUpperInvariant() == cleanText && cleanText.Any(char.IsLetter))
                    return true;

                // 一般的なヘッダーパターンの検出（マークダウン記号ベース）
                if (cleanText.StartsWith("第") && cleanText.Contains("章"))
                    return true;
                if (cleanText.StartsWith("第") && cleanText.Contains("節"))
                    return true;

                // 章・節パターン
                if (System.Text.RegularExpressions.Regex.IsMatch(cleanText, @"第\d+章|第\d+節|\d+\.\d+"))
                    return true;
            }
        }

        return false;
    }

    public static bool IsStrongHeaderPattern(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return false;

        var cleanText = text.Trim();

        // 明示的なMarkdownヘッダー
        if (cleanText.StartsWith("#"))
            return true;

        // 数字パターンヘッダー（1.1.1 形式）
        if (System.Text.RegularExpressions.Regex.IsMatch(cleanText, @"^(\d+\.){1,3}\s+\w"))
            return true;

        // 太字のパターンヘッダー
        if (IsPotentialBoldHeader(cleanText))
            return true;

        // 章・節の明示的パターン
        if (System.Text.RegularExpressions.Regex.IsMatch(cleanText, @"^(第\s*\d+\s*(章|節)|Chapter\s+\d+|Section\s+\d+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
            return true;

        return false;
    }

    public static bool IsListItemLike(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return false;

        var cleanText = text.Trim();

        // 明示的なMarkdownリスト記号
        if (cleanText.StartsWith("- ") || cleanText.StartsWith("* ") || cleanText.StartsWith("+ "))
            return true;

        // 日本語の箇条書き記号
        if (cleanText.StartsWith("・") || cleanText.StartsWith("•") || cleanText.StartsWith("◦"))
            return true;

        // ダッシュ系記号
        if (cleanText.StartsWith("‒") || cleanText.StartsWith("–") || cleanText.StartsWith("—"))
            return true;

        // 数字付きリスト
        if (System.Text.RegularExpressions.Regex.IsMatch(cleanText, @"^\d{1,3}[\.\)]\s+"))
            return true;

        // 括弧付きリスト
        if (System.Text.RegularExpressions.Regex.IsMatch(cleanText, @"^\(\d{1,3}\)\s+"))
            return true;

        // アルファベットリスト
        if (System.Text.RegularExpressions.Regex.IsMatch(cleanText, @"^[a-zA-Z][\.\)]\s+"))
            return true;

        // 太字リストマーカー（**-** など）
        if (System.Text.RegularExpressions.Regex.IsMatch(cleanText, @"^\*\*[‒–—\-\*\+•・]\*\*"))
            return true;

        return false;
    }

    public static bool IsTableRowLike(string text, List<Word> words)
    {
        if (string.IsNullOrWhiteSpace(text) || words == null || words.Count == 0)
            return false;

        var cleanText = text.Trim();

        // 明示的なテーブル記号
        if (cleanText.Contains("|"))
            return true;

        // タブ区切り
        if (cleanText.Contains("\t"))
            return true;

        // 複数の数値パターン
        if (HasNumericTablePattern(words))
            return true;

        // 規則的なスペーシング
        if (HasRegularSpacing(words))
            return true;

        // テーブル的テキストパターン
        if (HasTableLikeTextPattern(cleanText))
            return true;

        // 短い単語の配列（3つ以上）
        var wordList = cleanText.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        if (wordList.Length >= 3)
        {
            var allShort = wordList.All(w => w.Length <= 20);
            var avgLength = wordList.Average(w => w.Length);
            if (allShort && avgLength <= 10)
                return true;
        }

        // 座標ベースの規則的配置
        if (words.Count >= 3)
        {
            var sortedWords = words.OrderBy(w => w.BoundingBox.Left).ToList();
            var gaps = new List<double>();
            
            for (int i = 0; i < sortedWords.Count - 1; i++)
            {
                var gap = sortedWords[i + 1].BoundingBox.Left - sortedWords[i].BoundingBox.Right;
                gaps.Add(gap);
            }
            
            if (gaps.Any(g => g > 20)) // 大きなギャップがある
                return true;
        }

        return false;
    }

    public static bool IsHeaderStructure(string text, List<Word> words, double fontSize, FontAnalysis fontAnalysis)
    {
        if (string.IsNullOrWhiteSpace(text) || words == null || words.Count == 0)
            return false;

        var cleanText = text.Trim();

        // 明示的なマークダウンヘッダーパターン（最優先）
        if (cleanText.StartsWith("#"))
            return true;

        // フォントサイズベースの判定
        var fontSizeRatio = fontSize / fontAnalysis.BaseFontSize;
        
        // 明らかに段落的なテキストパターンを除外
        if (cleanText.EndsWith("。") || cleanText.EndsWith("です。") || cleanText.EndsWith("ます。") ||
            cleanText.EndsWith(".") || cleanText.Contains("、") || cleanText.Contains(",") ||
            cleanText.StartsWith("これは") || cleanText.StartsWith("それは") || 
            cleanText.StartsWith("この") || cleanText.StartsWith("その"))
        {
            return false;
        }
        
        // 長い文章（30文字以上）は段落の可能性が高い（閾値を緩和）
        if (cleanText.Length > 30)
        {
            return false;
        }
        
        // フォーマット記号を含むテキストは段落の可能性が高い
        if (cleanText.Contains("**") || (cleanText.Contains("*") && !cleanText.StartsWith("*")))
        {
            return false;
        }
        
        // リスト記号で始まるテキストはヘッダーではない
        if (cleanText.StartsWith("•") || cleanText.StartsWith("・") || cleanText.StartsWith("-") ||
            cleanText.StartsWith("◦") || cleanText.StartsWith("○"))
        {
            return false;
        }
        
        // 統計的フォントサイズ分析によるヘッダー判定（閾値を緩和）
        var largeFontRatio = fontSize / fontAnalysis.LargeFontThreshold;
        if (largeFontRatio >= 0.95) // LargeFontThreshold以上のサイズ（閾値緩和）
        {
            if (cleanText.Length <= 80) // 文字数制限を大幅緩和
                return true;
        }
        
        // より小さいフォントでも座標とテキスト長の組み合わせで判定（閾値緩和）
        if (fontSizeRatio >= 1.02) // 基底サイズより2%以上大きい（閾値緩和）
        {
            if (cleanText.Length <= 25) // 短いテキストなら許可（緩和）
                return true;
        }

        // 座標ベース判定の強化（統計的位置分析）
        var leftPosition = words.Min(w => w.BoundingBox.Left);
        var rightPosition = words.Max(w => w.BoundingBox.Right);
        var textWidth = rightPosition - leftPosition;
        
        // 左端配置で短いテキストの判定（閾値緩和）
        if (leftPosition <= 80.0 && cleanText.Length <= 50) // より寛容に
        {
            // 文の終端記号で終わるテキストは段落
            if (cleanText.EndsWith("。") || cleanText.EndsWith(".") || 
                cleanText.EndsWith("！") || cleanText.EndsWith("?"))
                return false;
            
            // フォントサイズが基底サイズ以上ならヘッダー候補（閾値緩和）
            if (fontSizeRatio >= 0.98) // より寛容に
                return true;
        }

        // 強いヘッダーパターン
        if (IsStrongHeaderPattern(cleanText))
            return true;

        // フォント分析による判定
        var hasBoldFont = words.Any(w => w.FontName?.ToLowerInvariant().Contains("bold") == true);
        if (hasBoldFont && cleanText.Length <= 120) // より寛容に
            return true;

        // 単語数による判定（保守的に）
        var wordCount = cleanText.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).Length;
        if (wordCount <= 8 && fontSizeRatio >= 1.1) // より寛容に
            return true;
            
        // 非常に短いテキストで相対的に大きなフォント
        if (cleanText.Length <= 12 && fontSizeRatio >= 1.05) // より寛容に
            return true;
            
        // テキスト幅が狭く独立している場合（レイアウト分析）
        if (textWidth <= 200.0 && cleanText.Length <= 30 && fontSizeRatio >= 0.95) // より寛容に
            return true;

        return false;
    }

    public static bool IsCodeBlockLike(string text, List<Word> words, FontAnalysis fontAnalysis)
    {
        if (string.IsNullOrWhiteSpace(text) || words == null || words.Count == 0)
            return false;

        var cleanText = text.Trim();

        // 明示的なコードブロック記号
        if (cleanText.StartsWith("```") || cleanText.EndsWith("```"))
            return true;

        if (cleanText.StartsWith("`") && cleanText.EndsWith("`"))
            return true;

        // インデントされたコード
        var leftPosition = words.Min(w => w.BoundingBox.Left);
        if (leftPosition > 80.0) // 大きくインデント
        {
            // プログラミングキーワードの検出
            var codeKeywords = new[] { "function", "class", "public", "private", "if", "else", "for", "while", "return" };
            if (codeKeywords.Any(keyword => cleanText.Contains(keyword)))
                return true;

            // 記号パターン（コード的）
            if (cleanText.Contains("{") || cleanText.Contains("}") || 
                cleanText.Contains("(") && cleanText.Contains(")") ||
                cleanText.Contains("=") || cleanText.Contains(";"))
                return true;
        }

        // 等幅フォントの検出
        var hasMonospaceFont = words.Any(w => 
            w.FontName?.ToLowerInvariant().Contains("mono") == true ||
            w.FontName?.ToLowerInvariant().Contains("courier") == true ||
            w.FontName?.ToLowerInvariant().Contains("consolas") == true);

        if (hasMonospaceFont)
            return true;

        // コードパターンの検出
        if (System.Text.RegularExpressions.Regex.IsMatch(cleanText, @"^[a-zA-Z_][a-zA-Z0-9_]*\s*[\(\)\{\}=;]"))
            return true;

        return false;
    }

    public static bool IsQuoteBlockLike(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return false;

        var cleanText = text.Trim();

        // 明示的な引用記号
        if (cleanText.StartsWith("> "))
            return true;

        // 引用符で囲まれている
        if ((cleanText.StartsWith("\"") && cleanText.EndsWith("\"")) ||
            (cleanText.StartsWith("'") && cleanText.EndsWith("'")) ||
            (cleanText.StartsWith("「") && cleanText.EndsWith("」")) ||
            (cleanText.StartsWith("『") && cleanText.EndsWith("』")))
            return true;

        // 引用を示すキーワード
        var quoteKeywords = new[] { "引用:", "Quote:", "出典:", "Source:" };
        if (quoteKeywords.Any(keyword => cleanText.StartsWith(keyword)))
            return true;

        return false;
    }

    public static bool HasBoldNumberPattern(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return false;

        // 太字数字パターンの検出
        return System.Text.RegularExpressions.Regex.IsMatch(text, @"\*\*\d+(\.\d+)?\*\*") ||
               System.Text.RegularExpressions.Regex.IsMatch(text, @"__\d+(\.\d+)?__");
    }

    public static bool IsPotentialBoldHeader(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return false;

        var cleanText = text.Trim();

        // 太字マークダウン記法
        if ((cleanText.StartsWith("**") && cleanText.EndsWith("**")) ||
            (cleanText.StartsWith("__") && cleanText.EndsWith("__")))
        {
            var innerText = cleanText.Substring(2, cleanText.Length - 4);
            return innerText.Length <= 100 && !innerText.Contains("\n");
        }

        // 太字パターンを含む短いテキスト
        if (System.Text.RegularExpressions.Regex.IsMatch(cleanText, @"\*\*[^*]+\*\*") && cleanText.Length <= 100)
            return true;

        return false;
    }

    private static bool HasNumericTablePattern(List<Word> words)
    {
        if (words == null || words.Count < 2)
            return false;

        // 数値パターンの統計的分析
        var numericWords = words.Where(w => 
            System.Text.RegularExpressions.Regex.IsMatch(w.Text?.Trim() ?? "", @"^\d+(\.\d+)?$")).ToList();
        var mixedWords = words.Where(w => 
            System.Text.RegularExpressions.Regex.IsMatch(w.Text?.Trim() ?? "", @"\d")).ToList();
        
        var numericRatio = (double)numericWords.Count / words.Count;
        var mixedRatio = (double)mixedWords.Count / words.Count;
        
        // 純粋な数値が多い場合
        if (numericWords.Count >= 2 && numericRatio >= 0.4)
            return true;
            
        // 数値を含む単語が多く、ギャップがある場合
        if (mixedWords.Count >= 3 && mixedRatio >= 0.5 && HasSignificantGaps(words))
            return true;

        return false;
    }
    
    private static bool HasSignificantGaps(List<Word> words)
    {
        if (words == null || words.Count < 2)
            return false;
            
        var sortedWords = words.OrderBy(w => w.BoundingBox.Left).ToList();
        var avgWordWidth = sortedWords.Average(w => w.BoundingBox.Width);
        
        for (int i = 0; i < sortedWords.Count - 1; i++)
        {
            var gap = sortedWords[i + 1].BoundingBox.Left - sortedWords[i].BoundingBox.Right;
            if (gap > avgWordWidth * 0.5) // 単語幅の50%以上のギャップ
                return true;
        }
        
        return false;
    }

    private static bool HasRegularSpacing(List<Word> words)
    {
        if (words == null || words.Count < 3)
            return false;

        var sortedWords = words.OrderBy(w => w.BoundingBox.Left).ToList();
        var gaps = new List<double>();
        var wordWidths = new List<double>();

        for (int i = 0; i < sortedWords.Count - 1; i++)
        {
            var gap = sortedWords[i + 1].BoundingBox.Left - sortedWords[i].BoundingBox.Right;
            gaps.Add(gap);
        }
        
        for (int i = 0; i < sortedWords.Count; i++)
        {
            var width = sortedWords[i].BoundingBox.Width;
            wordWidths.Add(width);
        }

        if (!gaps.Any())
            return false;

        // 統計的ギャップ分析
        var avgGap = gaps.Average();
        var avgWordWidth = wordWidths.Average();
        
        // ギャップの変動係数を計算
        var variance = gaps.Sum(g => Math.Pow(g - avgGap, 2)) / gaps.Count;
        var coefficient = avgGap > 0 ? Math.Sqrt(variance) / avgGap : double.MaxValue;
        
        // 大きなギャップの存在をチェック（テーブルセルの特徴）
        var largeGaps = gaps.Where(g => g > avgWordWidth * 0.5).ToList();
        var hasSignificantGaps = largeGaps.Count >= gaps.Count * 0.3; // 30%以上が大きなギャップ
        
        // ギャップパターンの判定
        var isRegularSpacing = coefficient < 0.6 && avgGap > 8;
        var hasTableLikeGaps = avgGap > 15 || hasSignificantGaps;
        
        return isRegularSpacing || hasTableLikeGaps;
    }

    private static bool HasTableLikeTextPattern(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return false;

        // 複数の空白が連続
        if (System.Text.RegularExpressions.Regex.IsMatch(text, @"\s{3,}"))
            return true;

        // パイプ記号
        if (text.Contains("|"))
            return true;

        // タブ文字
        if (text.Contains("\t"))
            return true;

        // 数字と文字の混在パターン
        var hasNumbers = System.Text.RegularExpressions.Regex.IsMatch(text, @"\d");
        var hasLetters = System.Text.RegularExpressions.Regex.IsMatch(text, @"[a-zA-Zあ-ん]");
        
        if (hasNumbers && hasLetters)
        {
            var words = text.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (words.Length >= 3)
                return true;
        }

        return false;
    }
}