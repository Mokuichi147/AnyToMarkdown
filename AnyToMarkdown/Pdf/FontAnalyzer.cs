using UglyToad.PdfPig.Content;

namespace AnyToMarkdown.Pdf;

internal static class FontAnalyzer
{
    public static FontAnalysis AnalyzeFontDistribution(List<Word> allWords)
    {
        var allSizes = allWords
            .Where(w => w.BoundingBox.Height > 0)
            .Select(w => w.BoundingBox.Height)
            .Distinct()
            .OrderBy(s => s)
            .ToList();

        if (allSizes.Count == 0)
        {
            return new FontAnalysis { BaseFontSize = 12.0, LargeFontThreshold = 14.0, AllFontSizes = new List<double> { 12.0 } };
        }

        var baseFontSize = CalculateBaseFontSize(allSizes);
        var largeFontThreshold = CalculateLargeFontThreshold(allSizes, baseFontSize);

        return new FontAnalysis
        {
            BaseFontSize = baseFontSize,
            LargeFontThreshold = largeFontThreshold,
            AllFontSizes = allSizes
        };
    }

    private static double CalculateBaseFontSize(List<double> allSizes)
    {
        if (allSizes.Count == 1) return allSizes[0];
        if (allSizes.Count == 0) return 12.0;

        // 統計的アプローチ：四分位数と頻度分析を組み合わせ
        var sortedSizes = allSizes.OrderBy(s => s).ToList();
        
        // 第1四分位数と第3四分位数を計算
        var q1Index = (int)(sortedSizes.Count * 0.25);
        var q3Index = (int)(sortedSizes.Count * 0.75);
        var q1 = sortedSizes[q1Index];
        var q3 = sortedSizes[q3Index];
        
        // IQR範囲内のサイズのみを使って基底サイズを決定（外れ値除去）
        var iqrSizes = sortedSizes.Where(s => s >= q1 && s <= q3).ToList();
        
        if (iqrSizes.Count == 0)
        {
            // フォールバック：中央値
            var middle = sortedSizes.Count / 2;
            return sortedSizes.Count % 2 == 0
                ? (sortedSizes[middle - 1] + sortedSizes[middle]) / 2.0
                : sortedSizes[middle];
        }
        
        // IQR範囲内で最頻値を計算
        var sizeGroups = iqrSizes
            .GroupBy(s => Math.Round(s, 1))
            .OrderByDescending(g => g.Count())
            .ThenBy(g => g.Key) // 同じ頻度なら小さいサイズを優先
            .ToList();

        return sizeGroups.First().Key;
    }
    
    private static double CalculateLargeFontThreshold(List<double> allSizes, double baseFontSize)
    {
        if (allSizes.Count <= 1) return baseFontSize * 1.2;
        
        // 統計的アプローチ：大きなフォントサイズの関係を分析
        var distinctSizes = allSizes.Distinct().OrderBy(s => s).ToList();
        
        // 基底サイズより大きいサイズを抽出
        var largerSizes = distinctSizes.Where(s => s > baseFontSize).ToList();
        
        if (largerSizes.Count == 0)
        {
            // 基底サイズしかない場合、僅かに大きい闾値を設定
            return baseFontSize * 1.1;
        }
        
        // 最小の大きいサイズを闾値として使用
        var threshold = largerSizes.First();
        
        // 闾値が基底サイズに近すぎる場合の調整
        var ratio = threshold / baseFontSize;
        if (ratio < 1.05) // 5%未満の差の場合
        {
            // 次に大きいサイズがあればそれを使用
            if (largerSizes.Count > 1)
            {
                threshold = largerSizes[1];
            }
            else
            {
                threshold = baseFontSize * 1.15; // フォールバック
            }
        }
        
        return threshold;
    }

    public static FontFormatting AnalyzeFontFormatting(List<Word> words)
    {
        var formatting = new FontFormatting();
        
        foreach (var word in words)
        {
            var fontName = word.FontName?.ToLowerInvariant() ?? "";
            
            // 日本語PDFではイタリックがフォント名に反映されない場合が多い
            // そのため、太字検出に重点を置く
            
            // 改良されたフォント検出パターン
            
            // 太字判定：より包括的で厳密なパターン
            var boldPattern = @"(bold|black|heavy|semibold|demibold|extrabold|ultrabold|medium|[6789]00|w[5-9]|thick|dark|strength)";
            if (System.Text.RegularExpressions.Regex.IsMatch(fontName, boldPattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase))
            {
                formatting.IsBold = true;
            }
            
            // フォントウェイト番号による太字判定の強化
            var weightMatch = System.Text.RegularExpressions.Regex.Match(fontName, @"(\d{3})");
            if (weightMatch.Success && int.TryParse(weightMatch.Groups[1].Value, out int weight))
            {
                if (weight >= 600) // 600以上は太字とみなす
                {
                    formatting.IsBold = true;
                }
            }
            
            // 斜体判定：より包括的で柔軟なパターン（大文字小文字を無視）
            var italicPattern = @"(italic|oblique|slanted|cursive|emphasis|stress|kursiv|inclined|tilted|skewed|angled)";
            if (System.Text.RegularExpressions.Regex.IsMatch(fontName, italicPattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase))
            {
                formatting.IsItalic = true;
            }
            
            // 追加の斜体パターン検出（より厳格に）
            if (fontName.Contains("-italic") || fontName.Contains("_italic") || 
                fontName.Contains("-oblique") || fontName.Contains("-slant"))
            {
                formatting.IsItalic = true;
            }
            
            // 明確な斜体パターンのみ
            if (fontName.EndsWith("-italic") || fontName.EndsWith("_italic") ||
                fontName.EndsWith("-oblique") || fontName.Contains("italicmt"))
            {
                formatting.IsItalic = true;
            }
            
            
            // PostScriptフォント名のサブセットタグを除去して再判定
            // 例: "EOODIA+Poetica-Bold" -> "Poetica-Bold"
            var cleanedFontName = System.Text.RegularExpressions.Regex.Replace(fontName, @"^[A-Z]{6}\+", "");
            if (cleanedFontName != fontName)
            {
                if (System.Text.RegularExpressions.Regex.IsMatch(cleanedFontName, boldPattern))
                {
                    formatting.IsBold = true;
                }
                if (System.Text.RegularExpressions.Regex.IsMatch(cleanedFontName, italicPattern))
                {
                    formatting.IsItalic = true;
                }
            }
        }
        
        return formatting;
    }

    public static string ApplyFormatting(string text, FontFormatting formatting)
    {
        // null文字と置換文字を除去
        text = text.Replace("\0", "").Replace("￿", "").Replace("\uFFFD", "");
        
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

    public static bool FormattingEqual(FontFormatting a, FontFormatting b)
    {
        return a.IsBold == b.IsBold && a.IsItalic == b.IsItalic;
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
    public List<double> AllFontSizes { get; set; } = new();
}