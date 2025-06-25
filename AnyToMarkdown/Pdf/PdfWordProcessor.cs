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
                
                // 行の高さの半分以内で、最も近い行を選択
                if (distance <= Math.Max(yThreshold, lineHeight / 2) && distance < bestMatch)
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
            
            if (distance < xThreshold)
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
}