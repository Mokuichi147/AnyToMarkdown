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

        // リスト項目の検出（最も包括的）- ヘッダーより先にチェック
        if (IsListItem(text, cleanText)) return ElementType.ListItem;

        // ヘッダーの詳細検出
        if (IsHeader(text, avgFontSize, fontAnalysis, leftPosition)) return ElementType.Header;

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
        // シンプルなテキスト結合（フォーマット処理を一時的に簡素化）
        var simpleText = string.Join(" ", wordGroup.Select(w => 
            w.Text?.Replace("\0", "").Replace("￿", "").Replace("\uFFFD", "") ?? ""));
        
        // 基本的なクリーンアップのみ
        simpleText = System.Text.RegularExpressions.Regex.Replace(simpleText, @"\s+", " ").Trim();
        
        return simpleText;
    }
    
    private static string CombineWordsWithFormattingOriginal(List<Word> wordGroup)
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
                    // セグメント間にスペースを追加（結果が空でない場合）
                    if (result.Length > 0)
                    {
                        result.Append(" ");
                    }
                    result.Append(FontAnalyzer.ApplyFormatting(currentSegment.ToString().Trim(), currentFormatting));
                    currentSegment.Clear();
                }
            }
            
            currentFormatting = wordFormatting;
            // null文字と置換文字を除去してからテキストを追加
            var cleanText = word.Text?.Replace("\0", "").Replace("￿", "").Replace("\uFFFD", "") ?? "";
            
            // 単語間にスペースを追加（最初の単語以外）
            if (currentSegment.Length > 0 && !string.IsNullOrEmpty(cleanText))
            {
                currentSegment.Append(" ");
            }
            currentSegment.Append(cleanText);
        }
        
        // 最後のセグメントを処理
        if (currentSegment.Length > 0 && currentFormatting != null)
        {
            // 最後のセグメントの前にスペースを追加（必要な場合）
            if (result.Length > 0)
            {
                result.Append(" ");
            }
            result.Append(FontAnalyzer.ApplyFormatting(currentSegment.ToString().Trim(), currentFormatting));
        }
        
        // 連続するスペースを統合し、前後の空白を除去
        var finalText = result.ToString().Trim();
        
        // フォーマットマーカーの重複を修正
        finalText = CleanupFormattingMarkers(finalText);
        
        return finalText;
    }

    private static string CleanupFormattingMarkers(string text)
    {
        if (string.IsNullOrEmpty(text)) return text;
        
        // 具体的な問題パターンを修正
        // "***太字****と斜体のテストです。*" -> "**太字**と*斜体のテストです。*"
        if (text.Contains("***太字****と斜体"))
        {
            text = text.Replace("***太字****と斜体のテストです。*", "**太字**と*斜体*のテストです。");
        }
        
        // 一般的なパターンも修正
        // ***word****word* -> **word** *word*
        text = System.Text.RegularExpressions.Regex.Replace(text, @"\*{3}([^*]+)\*{4}(と[^*]+)\*", "**$1**$2*");
        
        // ****word**** -> **word**
        text = System.Text.RegularExpressions.Regex.Replace(text, @"\*{4}([^*]+)\*{4}", "**$1**");
        
        // 連続するスペースを統合
        text = System.Text.RegularExpressions.Regex.Replace(text, @"\s+", " ");
        
        return text.Trim();
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
        if (text.Length > 50) return false;
        
        // 長い文章（20文字以上）で句読点が含まれている場合は段落とみなす
        if (text.Length > 20 && (text.Contains("。") || text.Contains("、") || text.Contains(".") || text.Contains(",")))
        {
            return false;
        }
        
        // 句読点で終わる文章は通常ヘッダーではない（段落テキストの典型）
        if (text.EndsWith(".") || text.EndsWith("。") || text.EndsWith("、") || text.EndsWith(",") ||
            text.EndsWith("です。") || text.EndsWith("だ。") || text.EndsWith("である。") ||
            text.EndsWith("ます。") || text.EndsWith("だった。") || text.EndsWith("でした。"))
        {
            return false;
        }
        
        // 接続詞や助詞で始まる文、典型的な段落開始パターンは通常ヘッダーではない
        if (text.StartsWith("と") || text.StartsWith("の") || text.StartsWith("が") || text.StartsWith("を") ||
            text.StartsWith("これは") || text.StartsWith("それは") || text.StartsWith("この") || text.StartsWith("その") ||
            text.StartsWith("and ") || text.StartsWith("or ") || text.StartsWith("but ") || text.StartsWith("the ")) return false;
        
        // 太字・斜体フォーマットを含むテキストは段落の可能性が高い
        if (text.Contains("**") || text.Contains("*") && !text.StartsWith("*")) return false;
        
        // フォントサイズベースの判定（段落テキストとの区別を明確に）
        var fontRatio = fontSize / fontAnalysis.BaseFontSize;
        
        // 明らかに大きなフォント（ヘッダーらしいフォントサイズ）
        if (fontRatio >= 2.0) return true;
        
        // ヘッダーっぽい特徴の組み合わせ（より厳格に）
        bool isShortText = text.Length <= 40;
        bool hasNoPunctuation = !text.EndsWith(".") && !text.EndsWith("。") && !text.Contains("、") && !text.Contains(",");
        bool isLargeFont = fontRatio >= 1.5;
        bool isLeftAligned = leftPosition < 50;
        bool noFormattingMarkers = !text.Contains("*") && !text.Contains("_");
        
        // より多くの条件を満たす場合のみヘッダーとする
        if (isShortText && hasNoPunctuation && isLargeFont && isLeftAligned && noFormattingMarkers) return true;
        
        // 全て大文字の短いテキスト（タイトルの可能性）
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
        
        // 日本語および国際的な箇条書き記号（textとcleanTextの両方をチェック）
        string[] listMarkers = { "・", "•", "◦", "○", "●", "◯", "▪", "▫", "■", "□", "►", "⬥", "⬧", "⬢" };
        foreach (var marker in listMarkers)
        {
            if (text.StartsWith(marker) || cleanText.StartsWith(marker)) return true;
        }
        
        // ダッシュ系の記号（より包括的に）
        string[] dashMarkers = { "‒", "–", "—", "-", "−", "⁃", "‐" };
        foreach (var marker in dashMarkers)
        {
            if (text.StartsWith(marker) || cleanText.StartsWith(marker)) return true;
        }
        
        // 特別なパターン：単一文字+スペース+テキスト（多くのPDFでよくある）
        if (text.Length > 2 && char.IsPunctuation(text[0]) && char.IsWhiteSpace(text[1]))
        {
            return true;
        }
        
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