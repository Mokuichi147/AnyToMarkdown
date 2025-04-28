using System.Text;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using SkiaSharp;

namespace AnyToMarkdown;

public static class PdfConverter
{
    public static ConvertResult Convert(FileStream stream)
    {
        StringBuilder sb = new();
        List<string> warnings = [];
        using var document = PdfDocument.Open(stream);
        foreach (Page page in document.GetPages())
        {
            List<Word> words = page.GetWords()
                .OrderByDescending(x => x.BoundingBox.Bottom)
                .ThenBy(x => x.BoundingBox.Left)
                .ToList();

            // PDF座標系での許容値
            double verticalTolerance = 5.0;
            double horizontalTolerance = 5;

            /*
            #if DEBUG
            foreach (var word in words)
            {
                Console.WriteLine($"{word.Text}: ({word.BoundingBox.Left}, {word.BoundingBox.Top}, {word.BoundingBox.Right}, {word.BoundingBox.Bottom})");
            }
            #endif
            */

            // 単語を行ごとにグループ化
            var lines = GroupWordsIntoLines(words, verticalTolerance);
            var images = page.GetImages().ToList();
            foreach (var line in lines)
            {
                // 行内の単語を結合して、Markdown形式に変換
                var mergedWords = MergeWordsInLine(line, horizontalTolerance);

                // 行内の画像を取得し、Markdown形式に変換
                var imgs = images.Where(x => x.Bounds.Bottom > line[0].BoundingBox.Bottom).OrderByDescending(x => x.Bounds.Top).ThenBy(x => x.Bounds.Left);
                foreach (var img in imgs)
                {
                    try
                    {
                        string base64 = ConvertImageToBase64(img.RawBytes);
                        sb.AppendLine($"![]({base64})\n");
                    }
                    catch (Exception ex)
                    {
                        warnings.Add($"Error processing image: {ex.Message}");
                    }
                    
                    images.Remove(img);
                }

                var lineText = string.Join(" ", mergedWords.Select(x => string.Join("", x.Select(w => w.Text))));
                sb.AppendLine(lineText + "\n");
            }

            foreach (var img in images)
            {
                try
                {
                    string base64 = ConvertImageToBase64(img.RawBytes);
                    sb.AppendLine($"![]({base64})\n");
                }
                catch (Exception ex)
                {
                    warnings.Add($"Error processing image: {ex.Message}");
                }
            }
        }
        var markdown = sb.ToString().Trim();

        #if DEBUG
        Console.WriteLine("PDF Conversion Result:");
        Console.WriteLine(markdown);

        Console.WriteLine("PDF Conversion Warnings:");
        foreach (var warning in warnings)
        {
            Console.WriteLine(warning);
        }
        #endif

        return new ConvertResult
        {
            Text = markdown,
            Warnings = warnings,
        };
    }

    /// <summary>
    /// 単語リストをY座標に基づいて行ごとにグループ化するメソッド
    /// </summary>
    /// <param name="words">ページ内の単語リスト</param>
    /// <returns>各行ごとの単語のリスト</returns>
    private static List<List<Word>> GroupWordsIntoLines(IEnumerable<Word> words, double yThreshold)
    {
        List<List<Word>> lines = [];

        foreach (var word in words)
        {
            bool added = false;
            // 既存の行に対して、Y座標の差で同じ行に属するかをチェックする
            foreach (var line in lines)
            {
                if (Math.Abs(word.BoundingBox.Bottom - line.Select(x => x.BoundingBox.Bottom).Average()) < yThreshold)
                {
                    line.Add(word);
                    added = true;
                    break;
                }
            }

            // どの行にも属さなければ、新しい行を追加
            if (!added)
            {
                lines.Add([word]);
            }
        }

        // 各行内の単語を、左側の座標順（X座標）で昇順ソート
        for (int i = 0; i < lines.Count; i++)
        {
            lines[i] = [.. lines[i].OrderBy(w => w.BoundingBox.Left)];
        }
        // 行全体を、ページ上部から下部へ（Y座標の高い順）並べ替え
        lines = [.. lines.OrderByDescending(line => line[0].BoundingBox.Bottom)];

        return lines;
    }

    /// <summary>
    /// 同一行内で、単語間のX方向のギャップが閾値より小さい場合、隣接単語を一塊にマージします。
    /// </summary>
    /// <param name="words">X座標で昇順に並んだ単語リスト</param>
    /// <param name="xThreshold">単語同士の間隔（前の単語の右端から次の単語の左端の距離）がこの値未満なら一塊として結合します。</param>
    /// <returns>行内で位置の近い単語をまとめる</returns>
    private static List<List<Word>> MergeWordsInLine(List<Word> words, double xThreshold)
    {
        if (words.Count == 0)
        {
            return [];
        }

        List<List<Word>> mergedWords = [];
        for (int i = 0; i < words.Count; i++)
        {
            if (mergedWords.Count == 0)
            {
                mergedWords.Add([words[0]]);
                continue;
            }
            
            if (Math.Abs(GetMax(mergedWords.Last(), x => x.BoundingBox.Right) - words[i].BoundingBox.Left) < xThreshold)
            {
                // 前の単語と結合
                int index = mergedWords.Count - 1;
                mergedWords[index].Add(words[i]);
            }
            else
            {
                // 新しい単語として追加
                mergedWords.Add([words[i]]);
            }
        }
        return mergedWords;
    }


    /// <summary>
    /// 指定された単語リストから、指定されたプロパティの最大値を取得します。
    /// </summary>
    /// <param name="words">単語リスト</param>
    /// <param name="selector">単語リストからプロパティを指定するセレクタ</param>
    /// <returns>指定されたプロパティの最大値</returns>
    private static double GetMax(IEnumerable<Word> words, Func<Word, double> selector)
    {
        if (words.Count() == 0)
        {
            return 0;
        }
        else
        {
            return words.Select(selector).Max();
        }
    }


    /// <summary>
    /// PDFPigで取得した画像の RawBytes をSkiaSharpでデコードし、BASE64文字列として返すメソッド
    /// </summary>
    /// <param name="imageBytes">PDFから抽出した画像のバイト配列</param>
    /// <param name="quality">エンコードの品質 (0-100)</param>
    /// <returns>PNG形式の画像を埋め込むためのData URL形式文字列 (例: "data:image/png;base64,......")</returns>
    static string ConvertImageToBase64(ReadOnlySpan<byte> imageBytes, int quality = 75)
    {
        using var bitmap = SKBitmap.Decode(imageBytes);
        using var encodedData = bitmap.Encode(SKEncodedImageFormat.Png, quality);
        if (encodedData == null)
        {
            return string.Empty;
        }

        byte[] webpBytes = encodedData.ToArray();
        string base64 = System.Convert.ToBase64String(webpBytes);
        return $"data:image/png;base64,{base64}";
    }
}