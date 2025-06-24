using System.Text;
using UglyToad.PdfPig.Content;

namespace AnyToMarkdown.Pdf;

internal class PdfStructureAnalyzer
{
    public static DocumentStructure AnalyzePageStructure(Page page, double horizontalTolerance = 5.0, double verticalTolerance = 5.0)
    {
        var words = page.GetWords()
            .OrderByDescending(x => x.BoundingBox.Bottom)
            .ThenBy(x => x.BoundingBox.Left)
            .ToList();

        var lines = PdfWordProcessor.GroupWordsIntoLines(words, verticalTolerance);
        var documentStructure = new DocumentStructure();

        // より詳細なフォント分析
        var fontAnalysis = AnalyzeFontDistribution(words);
        
        foreach (var line in lines)
        {
            var element = AnalyzeLine(line, fontAnalysis, horizontalTolerance);
            documentStructure.Elements.Add(element);
        }

        return documentStructure;
    }

    private static FontAnalysis AnalyzeFontDistribution(List<Word> words)
    {
        var fontSizes = words.GroupBy(w => Math.Round(w.BoundingBox.Height, 1))
                            .OrderByDescending(g => g.Count())
                            .ToList();

        var fontNames = words.GroupBy(w => w.FontName ?? "unknown")
                            .OrderByDescending(g => g.Count())
                            .ToList();

        double baseFontSize = fontSizes.First().Key;
        string dominantFont = fontNames.First().Key;

        return new FontAnalysis
        {
            BaseFontSize = baseFontSize,
            LargeFontThreshold = baseFontSize * 1.3,
            SmallFontThreshold = baseFontSize * 0.8,
            DominantFont = dominantFont,
            AllFontSizes = [.. fontSizes.Select(g => g.Key)]
        };
    }

    private static DocumentElement AnalyzeLine(List<Word> line, FontAnalysis fontAnalysis, double horizontalTolerance)
    {
        if (line.Count == 0)
            return new DocumentElement { Type = ElementType.Empty, Content = "" };

        var mergedWords = PdfWordProcessor.MergeWordsInLine(line, horizontalTolerance);
        
        // フォント情報を活用してMarkdown書式付きテキストを生成
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
        // 明確なMarkdownパターン
        if (text.StartsWith("#")) return ElementType.Header;
        if (text.StartsWith("-") || text.StartsWith("*") || text.StartsWith("+")) return ElementType.ListItem;
        if (text.StartsWith("・") || text.StartsWith("•")) return ElementType.ListItem;

        // フォントサイズベースの判定（最優先）
        if (maxFontSize > fontAnalysis.LargeFontThreshold)
        {
            if (IsHeaderLike(text)) return ElementType.Header;
        }

        // テーブル行の判定（座標とギャップベース）
        if (IsTableRowLike(text, words)) return ElementType.TableRow;

        // 位置ベースの判定
        if (isIndented)
        {
            if (IsListItemLike(text)) return ElementType.ListItem;
        }

        // パターンベースの判定
        if (IsHeaderLike(text)) return ElementType.Header;
        if (IsListItemLike(text)) return ElementType.ListItem;

        return ElementType.Paragraph;
    }

    private static bool IsHeaderLike(string text)
    {
        if (text.Length > 80) return false;
        if (text.EndsWith("。") || text.EndsWith(".") || text.EndsWith(",") || text.EndsWith("、")) return false;
        
        // 明確なMarkdownヘッダーパターン
        if (text.StartsWith("#")) return true;

        // 階層的な数字パターン (1., 1.1, 1.1.1, 1.1.1.1)
        var hierarchicalPattern = @"^\d+(\.\d+)*\.?\s";
        if (System.Text.RegularExpressions.Regex.IsMatch(text, hierarchicalPattern)) return true;

        // 数字で始まるパターン
        if (text.Length > 2 && char.IsDigit(text[0]) && (text.Contains(".") || text.Contains(")"))) return true;

        // 短くて大文字/カタカナ/漢字を含む（汎用的なタイトル判定）
        if (text.Length <= 30 && (text.Any(char.IsUpper) || 
            text.Any(c => c >= 0x30A0 && c <= 0x30FF) ||  // カタカナ
            text.Any(c => c >= 0x4E00 && c <= 0x9FAF)))   // 漢字
        {
            return true;
        }

        return false;
    }

    private static bool IsListItemLike(string text)
    {
        text = text.Trim();
        
        // 明確なリストマーカー
        if (text.StartsWith("・") || text.StartsWith("•") || text.StartsWith("◦")) return true;
        if (text.StartsWith("-") || text.StartsWith("*") || text.StartsWith("+")) return true;
        
        // 数字付きリスト
        if (text.Length > 2 && char.IsDigit(text[0]) && (text[1] == '.' || text[1] == ')')) return true;
        
        // 括弧付き数字
        if (text.Length > 3 && text[0] == '(' && char.IsDigit(text[1]) && text[2] == ')') return true;
        
        // アルファベット付きリスト
        if (text.Length > 2 && char.IsLetter(text[0]) && (text[1] == '.' || text[1] == ')')) return true;
        
        // Unicode数字記号（丸数字、四角数字など）
        if (text.Length > 0)
        {
            var firstChar = text[0];
            // 丸数字 (U+2460-U+2473)
            if (firstChar >= '\u2460' && firstChar <= '\u2473') return true;
            // 括弧付き数字 (U+2474-U+2487)
            if (firstChar >= '\u2474' && firstChar <= '\u2487') return true;
            // 丸括弧付き数字 (U+2488-U+249B)
            if (firstChar >= '\u2488' && firstChar <= '\u249B') return true;
        }

        return false;
    }

    private static bool IsTableRowLike(string text, List<Word> words)
    {
        var parts = text.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length < 2) return false; // 最低2列

        // パイプ文字がある（既にMarkdownテーブル）
        if (text.Contains("|")) return true;

        // アルファベット+数字の組み合わせ（A1, B1, C1など）- セル座標パターン
        bool hasAlphaNumPattern = parts.Count(p => p.Length == 2 && 
            char.IsLetter(p[0]) && char.IsDigit(p[1])) >= 2;
        if (hasAlphaNumPattern && parts.Length >= 3) return true;

        // 数値の比率が高い場合（統計データ、財務データなど）
        int numericParts = parts.Count(p => double.TryParse(p, out _) || p.All(char.IsDigit) || 
            p.Contains("%") || p.Contains(",") || p.StartsWith("+") || p.StartsWith("-"));
        if (parts.Length >= 3 && (double)numericParts / parts.Length > 0.4) return true;

        // 単語間の距離が大きく、均等に配置されている（表形式の最も重要な指標）
        if (words.Count >= 3)
        {
            var gaps = new List<double>();
            for (int i = 1; i < words.Count; i++)
            {
                gaps.Add(words[i].BoundingBox.Left - words[i-1].BoundingBox.Right);
            }
            
            if (gaps.Count >= 2)
            {
                var avgGap = gaps.Average();
                var largeGaps = gaps.Count(g => g > Math.Max(avgGap * 1.5, 15));
                var consistentGaps = gaps.Count(g => Math.Abs(g - avgGap) < avgGap * 0.5);
                
                // 大きなギャップが複数あり、かつ比較的一貫したギャップがある
                if (largeGaps >= 2 && consistentGaps >= gaps.Count * 0.6) return true;
            }
        }

        return false;
    }
    
    private static string BuildFormattedText(List<List<Word>> mergedWords)
    {
        var result = new System.Text.StringBuilder();
        
        foreach (var wordGroup in mergedWords)
        {
            if (wordGroup.Count == 0) continue;
            
            var groupText = string.Join("", wordGroup.Select(w => w.Text));
            if (string.IsNullOrWhiteSpace(groupText)) continue;
            
            // フォント情報から書式を判定
            var formatting = AnalyzeFontFormatting(wordGroup);
            
            // 書式を適用
            var formattedText = ApplyFormatting(groupText, formatting);
            
            if (result.Length > 0) result.Append(" ");
            result.Append(formattedText);
        }
        
        return result.ToString().Trim();
    }
    
    private static FontFormatting AnalyzeFontFormatting(List<Word> words)
    {
        var formatting = new FontFormatting();
        
        foreach (var word in words)
        {
            var fontName = word.FontName?.ToLowerInvariant() ?? "";
            
            // 太字の判定
            if (fontName.Contains("bold") || fontName.Contains("black") || fontName.Contains("heavy") || 
                fontName.Contains("medium") || fontName.Contains("semibold"))
            {
                formatting.IsBold = true;
            }
            
            // 斜体の判定
            if (fontName.Contains("italic") || fontName.Contains("oblique") || fontName.Contains("slanted"))
            {
                formatting.IsItalic = true;
            }
        }
        
        return formatting;
    }
    
    private static string ApplyFormatting(string text, FontFormatting formatting)
    {
        if (formatting.IsBold && formatting.IsItalic)
        {
            return $"***{text}***";
        }
        else if (formatting.IsBold)
        {
            return $"**{text}**";
        }
        else if (formatting.IsItalic)
        {
            return $"*{text}*";
        }
        
        return text;
    }
}

public class FontFormatting
{
    public bool IsBold { get; set; }
    public bool IsItalic { get; set; }
}

public class FontAnalysis
{
    public double BaseFontSize { get; set; }
    public double LargeFontThreshold { get; set; }
    public double SmallFontThreshold { get; set; }
    public string DominantFont { get; set; } = "";
    public List<double> AllFontSizes { get; set; } = [];
}

public class DocumentStructure
{
    public List<DocumentElement> Elements { get; set; } = [];
}

public class DocumentElement
{
    public ElementType Type { get; set; }
    public string Content { get; set; } = "";
    public double FontSize { get; set; }
    public double LeftMargin { get; set; }
    public bool IsIndented { get; set; }
    public List<Word> Words { get; set; } = [];
}

public enum ElementType
{
    Empty,
    Header,
    Paragraph,
    ListItem,
    TableRow
}