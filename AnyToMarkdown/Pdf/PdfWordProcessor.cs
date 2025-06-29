using UglyToad.PdfPig.Content;

namespace AnyToMarkdown.Pdf;

internal static class PdfWordProcessor
{
    public static List<List<Word>> GroupWordsIntoLines(IEnumerable<Word> words, double yThreshold)
    {
        var lines = new List<List<Word>>();
        var sortedWords = words.OrderByDescending(w => w.BoundingBox.Bottom).ThenBy(w => w.BoundingBox.Left);

        foreach (var word in sortedWords)
        {
            bool added = false;
            double bestMatch = double.MaxValue;
            int bestLineIndex = -1;
            
            // より良いマッチング：最も近い行を見つける
            for (int i = 0; i < lines.Count; i++)
            {
                var line = lines[i];
                var avgBottom = line.Select(x => x.BoundingBox.Bottom).Average();
                var avgTop = line.Select(x => x.BoundingBox.Top).Average();
                var lineHeight = avgTop - avgBottom;
                
                // 単語のベースラインと行のベースラインの距離
                var distance = Math.Abs(word.BoundingBox.Bottom - avgBottom);
                
                // より精密な境界検出：上下両方向の重複をチェック
                var wordTop = word.BoundingBox.Top;
                var wordBottom = word.BoundingBox.Bottom;
                var lineTop = avgTop;
                var lineBottom = avgBottom;
                
                // 垂直重複を計算
                var overlapTop = Math.Min(wordTop, lineTop);
                var overlapBottom = Math.Max(wordBottom, lineBottom);
                var overlap = Math.Max(0, overlapTop - overlapBottom);
                var wordHeight = wordTop - wordBottom;
                var overlapRatio = wordHeight > 0 ? overlap / wordHeight : 0;
                
                // 重複率が高い、または距離が閾値内で最も近い行を選択（改善版）
                var dynamicThreshold = Math.Max(yThreshold, Math.Min(lineHeight * 0.8, wordHeight * 0.8));
                if ((overlapRatio > 0.25 || distance <= dynamicThreshold) && 
                    distance < bestMatch)
                {
                    bestMatch = distance;
                    bestLineIndex = i;
                }
            }
            
            if (bestLineIndex >= 0)
            {
                lines[bestLineIndex].Add(word);
                added = true;
            }

            if (!added)
            {
                lines.Add([word]);
            }
        }

        // 各行内で左から右へソート
        for (int i = 0; i < lines.Count; i++)
        {
            lines[i] = [.. lines[i].OrderBy(w => w.BoundingBox.Left)];
        }
        
        // 行を上から下へソート
        lines = [.. lines.OrderByDescending(line => line.Select(w => w.BoundingBox.Bottom).Average())];

        return lines;
    }

    public static List<List<Word>> MergeWordsInLine(List<Word> words, double xThreshold)
    {
        if (words.Count == 0)
        {
            return [];
        }

        var mergedWords = new List<List<Word>>();
        
        for (int i = 0; i < words.Count; i++)
        {
            if (mergedWords.Count == 0)
            {
                mergedWords.Add([words[i]]);
                continue;
            }
            
            // 前のグループの右端と現在の単語の左端の距離を計算
            var lastGroup = mergedWords.Last();
            var distance = words[i].BoundingBox.Left - GetMaxRight(lastGroup);
            
            // より精密な単語境界検出
            var currentWordHeight = words[i].BoundingBox.Height;
            var lastGroupAvgHeight = lastGroup.Count > 0 ? lastGroup.Average(w => w.BoundingBox.Height) : currentWordHeight;
            
            // 文字サイズに基づく動的な閾値調整（改善版）
            var fontBasedThreshold = Math.Min(currentWordHeight, lastGroupAvgHeight) * 0.4;
            var adaptiveThreshold = Math.Max(xThreshold, fontBasedThreshold);
            
            // 文字の重複やマイナス距離の場合は強制的に統合
            // また、同じ単語の文字間スペースが異常に大きい場合も考慮
            if (distance < adaptiveThreshold || distance < 0 || 
                (distance <= fontBasedThreshold * 2 && ShouldMergeWords(lastGroup.Last(), words[i])))
            {
                // 近接している場合は同じグループに追加
                lastGroup.Add(words[i]);
            }
            else
            {
                // 離れている場合は新しいグループを作成
                mergedWords.Add([words[i]]);
            }
        }
        return mergedWords;
    }

    private static double GetMaxRight(IEnumerable<Word> words)
    {
        return words.Count() == 0 ? 0 : words.Select(w => w.BoundingBox.Right).Max();
    }
    
    private static bool ShouldMergeWords(Word word1, Word word2)
    {
        // フォントサイズと名前が似ている場合は同じ文字列の一部の可能性
        var fontSizeDiff = Math.Abs(word1.BoundingBox.Height - word2.BoundingBox.Height);
        var fontSizeThreshold = Math.Min(word1.BoundingBox.Height, word2.BoundingBox.Height) * 0.1;
        
        var font1 = word1.FontName ?? "";
        var font2 = word2.FontName ?? "";
        bool sameFontFamily = font1.Equals(font2, StringComparison.OrdinalIgnoreCase) ||
                             (font1.Length > 0 && font2.Length > 0 && 
                              font1.Substring(0, Math.Min(font1.Length, 6)).Equals(
                                  font2.Substring(0, Math.Min(font2.Length, 6)), StringComparison.OrdinalIgnoreCase));
        
        // 垂直位置も考慮
        var verticalDistance = Math.Abs(word1.BoundingBox.Bottom - word2.BoundingBox.Bottom);
        var heightThreshold = Math.Max(word1.BoundingBox.Height, word2.BoundingBox.Height) * 0.3;
        
        return fontSizeDiff <= fontSizeThreshold && sameFontFamily && verticalDistance <= heightThreshold;
    }
}