using UglyToad.PdfPig.Content;

namespace AnyToMarkdown.Pdf;

internal static class PdfWordProcessor
{
    public static List<List<Word>> GroupWordsIntoLines(IEnumerable<Word> words, double yThreshold)
    {
        var lines = new List<List<Word>>();

        foreach (var word in words)
        {
            bool added = false;
            foreach (var line in lines)
            {
                if (Math.Abs(word.BoundingBox.Bottom - line.Select(x => x.BoundingBox.Bottom).Average()) < yThreshold)
                {
                    line.Add(word);
                    added = true;
                    break;
                }
            }

            if (!added)
            {
                lines.Add([word]);
            }
        }

        for (int i = 0; i < lines.Count; i++)
        {
            lines[i] = [.. lines[i].OrderBy(w => w.BoundingBox.Left)];
        }
        
        lines = [.. lines.OrderByDescending(line => line[0].BoundingBox.Bottom)];

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
                mergedWords.Add([words[0]]);
                continue;
            }
            
            if (Math.Abs(GetMaxRight(mergedWords.Last()) - words[i].BoundingBox.Left) < xThreshold)
            {
                int index = mergedWords.Count - 1;
                mergedWords[index].Add(words[i]);
            }
            else
            {
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