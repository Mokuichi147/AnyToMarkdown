using System.Linq;
using UglyToad.PdfPig.Content;

namespace AnyToMarkdown.Pdf;

internal static class ElementClassifier
{
    public static ElementType ClassifyElement(List<Word> wordGroup, FontAnalysis fontAnalysis)
    {
        if (wordGroup.Count == 0) return ElementType.Paragraph;

        var text = CombineWordsWithFormatting(wordGroup);
        var cleanText = text.Replace("*", "").Replace("_", "").Replace("#", "").Trim();

        // 空行チェック
        if (string.IsNullOrWhiteSpace(cleanText)) return ElementType.Paragraph;

        // より精密な要素分類
        var avgFontSize = wordGroup.Average(w => w.BoundingBox.Height);
        var leftPosition = wordGroup.Min(w => w.BoundingBox.Left);

        // 水平線の検出（強化版）
        if (IsHorizontalLine(text)) return ElementType.HorizontalLine;

        // ヘッダーの詳細検出
        if (IsHeader(cleanText, avgFontSize, fontAnalysis, leftPosition)) return ElementType.Header;

        // リスト項目の検出（最も包括的）
        if (IsListItem(text, cleanText)) return ElementType.ListItem;

        // コードブロックの検出
        if (IsCodeBlock(text, cleanText)) return ElementType.CodeBlock;

        // 引用ブロックの検出
        if (IsQuoteBlock(text, cleanText)) return ElementType.QuoteBlock;

        // テーブル行の予備検出（後でより詳細に分析される）
        if (IsTableRow(wordGroup, text)) return ElementType.TableRow;

        return ElementType.Paragraph;
    }

    private static string CombineWordsWithFormatting(List<Word> wordGroup)
    {
        var result = new System.Text.StringBuilder();
        FontFormatting? currentFormatting = null;
        var currentSegment = new System.Text.StringBuilder();
        
        foreach (var word in wordGroup)
        {
            var wordFormatting = FontAnalyzer.AnalyzeFontFormatting([word]);
            
            // 書式が変わった場合、現在のセグメントを確定
            if (currentFormatting != null && !FontAnalyzer.FormattingEqual(currentFormatting, wordFormatting))
            {
                if (currentSegment.Length > 0)
                {
                    result.Append(FontAnalyzer.ApplyFormatting(currentSegment.ToString(), currentFormatting));
                    currentSegment.Clear();
                }
            }
            
            currentFormatting = wordFormatting;
            // null文字と置換文字を除去してからテキストを追加
            var cleanText = word.Text?.Replace("\0", "").Replace("￿", "").Replace("\uFFFD", "") ?? "";
            currentSegment.Append(cleanText);
        }
        
        // 最後のセグメントを処理
        if (currentSegment.Length > 0 && currentFormatting != null)
        {
            result.Append(FontAnalyzer.ApplyFormatting(currentSegment.ToString(), currentFormatting));
        }
        
        // 連続するスペースを統合し、前後の空白を除去
        var finalText = result.ToString().Trim();
        
        // 重複を避けるため、最初のアプローチの結果をそのまま使用
        return finalText;
    }

    private static bool IsHorizontalLine(string text)
    {
        // より包括的な水平線パターンの検出
        var cleanText = text.Replace(" ", "").Replace("\t", "");
        
        // 基本的なマークダウン水平線
        if (cleanText == "---" || cleanText == "***" || cleanText == "___") return true;
        
        // より長い水平線パターン
        if (cleanText.Length >= 3)
        {
            if (cleanText.All(c => c == '-') || 
                cleanText.All(c => c == '=') || 
                cleanText.All(c => c == '_') ||
                cleanText.All(c => c == '*'))
            {
                return true;
            }
        }
        
        // Unicode 文字による水平線
        if (text.Contains("─") || text.Contains("━") || text.Contains("═") || text.Contains("—"))
        {
            return true;
        }
        
        return false;
    }

    private static bool IsHeader(string text, double fontSize, FontAnalysis fontAnalysis, double leftPosition)
    {
        // 既にMarkdownヘッダーの場合
        if (text.StartsWith("#")) return true;
        
        // Markdownの書式記号で囲まれたテキストはヘッダーではない
        if ((text.StartsWith("**") && text.EndsWith("**")) ||
            (text.StartsWith("*") && text.EndsWith("*") && !text.StartsWith("**")) ||
            (text.StartsWith("_") && text.EndsWith("_")) ||
            (text.StartsWith("`") && text.EndsWith("`")))
        {
            return false;
        }
        
        // 明らかに段落テキストでない場合の判定を厳しくする
        if (text.Length > 80) return false;
        
        // 句読点で終わる文章は通常ヘッダーではない
        if (text.EndsWith(".") || text.EndsWith("。") || text.EndsWith("、") || text.EndsWith(",")) return false;
        
        // 接続詞や助詞で始まる文は通常ヘッダーではない（日本語対応）
        if (text.StartsWith("と") || text.StartsWith("の") || text.StartsWith("が") || text.StartsWith("を") ||
            text.StartsWith("and ") || text.StartsWith("or ") || text.StartsWith("but ") || text.StartsWith("the ")) return false;
        
        // フォントサイズベースの判定（より厳しく）
        var fontRatio = fontSize / fontAnalysis.BaseFontSize;
        if (fontRatio >= 1.4) return true; // 閾値を上げる
        
        // 左端配置かつ大きなフォントの場合のみ
        if (leftPosition < 50 && fontRatio >= 1.25) return true; // 閾値を上げる
        
        // 全て大文字の短いテキスト（より厳しく）
        if (text.Length <= 30 && text.Length >= 3 && 
            text.All(c => char.IsUpper(c) || char.IsWhiteSpace(c) || char.IsPunctuation(c) || char.IsDigit(c)))
        {
            return true;
        }
        
        return false;
    }

    private static bool IsListItem(string text, string cleanText)
    {
        // Markdownスタイルのリスト
        if (text.StartsWith("- ") || text.StartsWith("* ") || text.StartsWith("+ ")) return true;
        
        // 番号付きリスト
        if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^\d+\.\s+")) return true;
        
        // 日本語および国際的な箇条書き記号
        if (cleanText.StartsWith("・") || cleanText.StartsWith("•")) return true;
        if (cleanText.StartsWith("‒") || cleanText.StartsWith("–") || cleanText.StartsWith("—")) return true;
        
        // 太字でフォーマットされたリスト記号も検出
        if (text.StartsWith("**‒**") || text.StartsWith("**–**") || text.StartsWith("**—**")) return true;
        if (text.StartsWith("**-**") || text.StartsWith("**+**") || text.StartsWith("***") && text.Contains("***")) return true;
        if (text.StartsWith("**•**") || text.StartsWith("**・**")) return true;
        
        // より複雑な太字リストパターンの検出
        var boldListPattern = System.Text.RegularExpressions.Regex.Match(text, @"^\*\*([‒–—\-\*\+•・])\*\*");
        if (boldListPattern.Success) return true;
        
        // アルファベット順リスト (a), (b), (c) など
        if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^\([a-zA-Z]\)\s+")) return true;
        
        // ローマ数字リスト
        if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^[ivxlcdm]+\.\s+", System.Text.RegularExpressions.RegexOptions.IgnoreCase)) return true;
        
        return false;
    }

    private static bool IsCodeBlock(string text, string cleanText)
    {
        // バッククォートで囲まれたコード
        if (text.StartsWith("```") || (text.StartsWith("`") && text.EndsWith("`"))) return true;
        
        // インデントされたコード（4スペース以上）
        if (text.StartsWith("    ") || text.StartsWith("\t")) return true;
        
        // プログラミング言語のキーワードや構文を含む
        var codePatterns = new[]
        {
            @"^(function|var|let|const|class|def|import|export)\s+",
            @"^(public|private|protected|static)\s+",
            @"^\s*(if|else|for|while|switch|try|catch)\s*\(",
            @"^#include\s+",
            @"^\s*<\w+.*>.*<\/\w+>\s*$", // HTML tags
            @"console\.(log|error|warn|info)\s*\(",
            @"print\s*\(",
            @"System\.(out|err)\.print",
            @"^[a-zA-Z_]\w*\s*\(.*\)\s*\{?$" // Function calls
        };
        
        return codePatterns.Any(pattern => 
            System.Text.RegularExpressions.Regex.IsMatch(text, pattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase));
    }

    private static bool IsQuoteBlock(string text, string cleanText)
    {
        // Markdownスタイルの引用
        if (text.StartsWith("> ")) return true;
        
        // 他の引用記号
        if (text.StartsWith("\" ") && text.EndsWith("\"")) return true;
        if (text.StartsWith("' ") && text.EndsWith("'")) return true;
        
        // 日本語の引用符
        if (text.StartsWith("「") && text.EndsWith("」")) return true;
        if (text.StartsWith("『") && text.EndsWith("』")) return true;
        
        return false;
    }

    private static bool IsTableRow(List<Word> wordGroup, string text)
    {
        // Markdownテーブル記法
        if (text.StartsWith("|") && text.EndsWith("|")) return true;
        
        // パイプで区切られたテキスト（最低2つのパイプ）
        if (text.Count(c => c == '|') >= 2) return true;
        
        // タブで区切られた複数の項目
        if (text.Count(c => c == '\t') >= 2) return true;
        
        // 数値データパターン（表の可能性）
        var words = text.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
        if (words.Length >= 3)
        {
            var numericCount = words.Count(w => IsNumericValue(w));
            if (numericCount >= 2) return true;
        }
        
        // 座標ベースの分析：単語間の均等な間隔
        if (wordGroup.Count >= 3)
        {
            var positions = wordGroup.Select(w => w.BoundingBox.Left).OrderBy(p => p).ToList();
            var gaps = new List<double>();
            
            for (int i = 1; i < positions.Count; i++)
            {
                gaps.Add(positions[i] - positions[i-1]);
            }
            
            // 間隔が比較的均等な場合（標準偏差が小さい）
            if (gaps.Count >= 2)
            {
                var avgGap = gaps.Average();
                var variance = gaps.Sum(g => Math.Pow(g - avgGap, 2)) / gaps.Count;
                var stdDev = Math.Sqrt(variance);
                
                if (stdDev < avgGap * 0.5) // 変動係数が50%未満
                {
                    return true;
                }
            }
        }
        
        return false;
    }

    private static bool IsNumericValue(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return false;
        
        // 純粋な数値
        if (double.TryParse(text, out _)) return true;
        
        // 通貨記号付き
        if (text.StartsWith("$") || text.StartsWith("¥") || text.StartsWith("€") || text.StartsWith("£"))
        {
            return double.TryParse(text.Substring(1), out _);
        }
        
        // パーセンテージ
        if (text.EndsWith("%"))
        {
            return double.TryParse(text.Substring(0, text.Length - 1), out _);
        }
        
        // カンマ区切りの数値
        var withoutCommas = text.Replace(",", "");
        return double.TryParse(withoutCommas, out _);
    }
}

public enum ElementType
{
    Empty,
    Header,
    Paragraph,
    ListItem,
    TableRow,
    CodeBlock,
    QuoteBlock,
    HorizontalLine
}

public class DocumentElement
{
    public ElementType Type { get; set; }
    public string Content { get; set; } = "";
    public double FontSize { get; set; }
    public double LeftMargin { get; set; }
    public bool IsIndented { get; set; }
    public List<Word> Words { get; set; } = new();
}

public class DocumentStructure
{
    public List<DocumentElement> Elements { get; set; } = new();
    public FontAnalysis FontAnalysis { get; set; } = new();
}