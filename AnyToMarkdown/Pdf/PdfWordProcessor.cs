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
                
                // テーブル行検出のためより厳格な重複基準（CLAUDE.md準拠）
                var dynamicThreshold = Math.Max(yThreshold, Math.Min(lineHeight * 0.5, wordHeight * 0.5));
                if ((overlapRatio > 0.4 || distance <= dynamicThreshold) && 
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
            
            // テーブルセル分離のため厳格な閾値調整（CLAUDE.md準拠）
            var fontBasedThreshold = Math.Min(currentWordHeight, lastGroupAvgHeight) * 0.3;
            var adaptiveThreshold = Math.Max(xThreshold, fontBasedThreshold);
            
            // 文字の重複時のみ統合（テーブルセル境界を保持）
            if (distance < 0 || 
                (distance <= fontBasedThreshold && ShouldMergeWords(lastGroup.Last(), words[i])))
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
        var fontSizeThreshold = Math.Min(word1.BoundingBox.Height, word2.BoundingBox.Height) * 0.05;
        
        var font1 = word1.FontName ?? "";
        var font2 = word2.FontName ?? "";
        bool sameFontFamily = font1.Equals(font2, StringComparison.OrdinalIgnoreCase) ||
                             (font1.Length > 0 && font2.Length > 0 && 
                              font1.Substring(0, Math.Min(font1.Length, 6)).Equals(
                                  font2.Substring(0, Math.Min(font2.Length, 6)), StringComparison.OrdinalIgnoreCase));
        
        // 垂直位置も考慮（テーブルセル分離のため厳格化）
        var verticalDistance = Math.Abs(word1.BoundingBox.Bottom - word2.BoundingBox.Bottom);
        var heightThreshold = Math.Max(word1.BoundingBox.Height, word2.BoundingBox.Height) * 0.15;
        
        // 水平距離も考慮（テーブルセルの境界検出）
        var horizontalGap = word2.BoundingBox.Left - word1.BoundingBox.Right;
        var avgWidth = (word1.BoundingBox.Width + word2.BoundingBox.Width) / 2;
        
        // テーブルセル間のギャップが大きい場合は統合しない
        if (horizontalGap > avgWidth * 0.5)
        {
            return false;
        }
        
        return fontSizeDiff <= fontSizeThreshold && sameFontFamily && verticalDistance <= heightThreshold;
    }
}