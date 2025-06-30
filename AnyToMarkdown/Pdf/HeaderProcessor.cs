using System.Linq;

namespace AnyToMarkdown.Pdf;

internal static class HeaderProcessor
{
    public static int DetermineHeaderLevel(DocumentElement element, FontAnalysis fontAnalysis)
    {
        var content = element.Content?.Trim() ?? "";
        var leftPosition = element.LeftMargin;
        
        // 空のコンテンツは通常のテキストとして扱う
        if (string.IsNullOrWhiteSpace(content)) return 6;
        
        // フォントサイズ分析
        var allSizes = fontAnalysis.AllFontSizes.Distinct().OrderByDescending(s => s).ToList();
        var fontSizeScore = CalculateFontSizeScore(element.FontSize, allSizes);
        
        // 座標位置分析（左端に近いほど上位レベル）
        var coordinateScore = CalculateCoordinateScore(leftPosition);
        
        // テキスト長分析（短いほど上位レベルの可能性）
        var lengthScore = CalculateTextLengthScore(element.Content);
        
        // フォントサイズを主軸にした重み付け統合スコア
        var combinedScore = fontSizeScore * 0.7 + coordinateScore * 0.2 + lengthScore * 0.1;
        
        // フォントサイズの絶対値も考慮したより精密な閾値
        var baseFontRatio = element.FontSize / fontAnalysis.BaseFontSize;
        
        // 座標ベースの階層検出を強化
        var hierarchyLevel = CalculateHierarchyLevel(leftPosition);
        
        // コンテンツベースのレベル検出を優先
        var explicitLevel = GetExplicitHeaderLevel(content);
        if (explicitLevel > 0)
        {
            return explicitLevel;
        }
        
        // より精密なレベル決定（相対的なフォントサイズに基づく）
        
        // 最大フォントサイズとの比較でレベルを決定
        var maxFontSize = allSizes.Count > 0 ? allSizes[0] : element.FontSize;
        var relativeRatio = element.FontSize / maxFontSize;
        
        // LargeFontThresholdを基準とした統計的レベル判定
        var largeFontRatio = element.FontSize / fontAnalysis.LargeFontThreshold;
        
        if (largeFontRatio >= 1.8) return 1;   // 非常に大きい
        if (largeFontRatio >= 1.4) return 2;   // 大きい
        if (largeFontRatio >= 1.1) return 3;   // 中程度
        if (largeFontRatio >= 1.0) return 4;   // LargeFontThreshold以上
        if (baseFontRatio >= 1.05) return 5;   // 基底サイズより少し大きい
        
        // フォールバック：相対的サイズ判定
        if (relativeRatio >= 0.95) return 1;
        if (relativeRatio >= 0.85) return 2;
        if (relativeRatio >= 0.75) return 3;
        
        // デフォルトで3レベルを返す（6は使わない）
        return 3;
    }
    
    private static double CalculateFontSizeScore(double fontSize, List<double> allSizes)
    {
        if (allSizes.Count == 0) return 0.5;
        
        // より詳細なフォントサイズ分析
        var sortedSizes = allSizes.OrderByDescending(s => s).ToList();
        var currentSizeRank = sortedSizes.IndexOf(fontSize);
        if (currentSizeRank < 0) return 0.3;
        
        // フォントサイズの相対的位置を計算
        var totalSizes = sortedSizes.Count;
        var normalizedRank = 1.0 - ((double)currentSizeRank / Math.Max(totalSizes - 1, 1));
        
        // 上位のサイズに対してボーナススコア
        if (currentSizeRank == 0) return 1.0;  // 最大サイズ
        if (currentSizeRank == 1 && totalSizes > 2) return 0.9;  // 第2位
        if (currentSizeRank == 2 && totalSizes > 3) return 0.8;  // 第3位
        
        return normalizedRank;
    }
    
    private static double CalculateCoordinateScore(double leftPosition)
    {
        // 左端からの距離を正規化（より精密な階層認識）
        var basePosition = 30.0;
        var normalizedPosition = Math.Max(0, leftPosition - basePosition);
        
        // 左端に近いほど高いスコア
        var maxDistance = 200.0;
        var score = 1.0 - Math.Min(normalizedPosition / maxDistance, 1.0);
        
        // 完全に左端（0-30px）は最高スコア
        if (leftPosition <= basePosition) return 1.0;
        
        return Math.Max(score, 0.1);
    }
    
    private static double CalculateTextLengthScore(string? content)
    {
        if (string.IsNullOrWhiteSpace(content)) return 0.0;
        
        var length = content?.Trim().Length ?? 0;
        
        // ヘッダーに適した長さの評価
        if (length <= 10) return 1.0;
        if (length <= 20) return 0.9;
        if (length <= 30) return 0.8;
        if (length <= 50) return 0.6;
        if (length <= 80) return 0.4;
        
        return 0.2; // 長すぎる場合はヘッダーとしての可能性を下げる
    }
    
    private static int CalculateHierarchyLevel(double leftPosition)
    {
        // インデントレベルに基づく階層の推定
        var indentUnit = 20.0; // 1レベルあたりのインデント幅（ピクセル）
        var basePosition = 30.0;
        
        var normalizedPosition = Math.Max(0, leftPosition - basePosition);
        var level = (int)(normalizedPosition / indentUnit) + 1;
        
        return Math.Min(Math.Max(level, 1), 6);
    }
    
    private static int GetExplicitHeaderLevel(string content)
    {
        if (string.IsNullOrWhiteSpace(content)) return 0;
        
        // 既存のMarkdownヘッダー記法を検出
        var trimmed = content.Trim();
        if (!trimmed.StartsWith("#")) return 0;
        
        var hashCount = 0;
        foreach (var c in trimmed)
        {
            if (c == '#') hashCount++;
            else if (c == ' ') break;
            else return 0; // 無効なヘッダー記法
        }
        
        return Math.Min(Math.Max(hashCount, 1), 6);
    }

    public static string ConvertHeader(DocumentElement element, FontAnalysis fontAnalysis)
    {
        var text = element.Content.Trim();
        
        // 既にMarkdownヘッダーの場合はそのまま
        if (text.StartsWith("#")) return text;
        
        // ヘッダーのフォーマットを除去してクリーンなテキストを取得
        var cleanText = StripMarkdownFormatting(text);
        
        // 空のヘッダーは無視
        if (string.IsNullOrWhiteSpace(cleanText)) return "";
        
        // 既存のマークダウンヘッダーの場合はそのまま処理
        if (cleanText.StartsWith("#"))
        {
            return cleanText;
        }
        
        var level = DetermineHeaderLevel(element, fontAnalysis);
        
        // レベルが無効でも最低3レベルヘッダーにする
        if (level > 5) level = 3;
        
        var prefix = new string('#', level);
        
        return $"{prefix} {cleanText}";
    }

    private static string StripMarkdownFormatting(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return "";
        
        // より安全なフォーマット除去（対称的なマークを正確に除去）
        
        // ***太字斜体*** パターンを最初に処理
        text = System.Text.RegularExpressions.Regex.Replace(text, @"\*\*\*(.+?)\*\*\*", "$1");
        
        // **太字** パターンを処理
        text = System.Text.RegularExpressions.Regex.Replace(text, @"\*\*(.+?)\*\*", "$1");
        
        // *斜体* パターンを処理
        text = System.Text.RegularExpressions.Regex.Replace(text, @"\*(.+?)\*", "$1");
        
        // __太字__ パターンを処理
        text = System.Text.RegularExpressions.Regex.Replace(text, @"__(.+?)__", "$1");
        
        // _斜体_ パターンを処理
        text = System.Text.RegularExpressions.Regex.Replace(text, @"_(.+?)_", "$1");
        
        // ~~取り消し線~~ パターンを処理
        text = System.Text.RegularExpressions.Regex.Replace(text, @"~~(.+?)~~", "$1");
        
        // `インラインコード` パターンを処理
        text = System.Text.RegularExpressions.Regex.Replace(text, @"`(.+?)`", "$1");
        
        return text.Trim();
    }
}