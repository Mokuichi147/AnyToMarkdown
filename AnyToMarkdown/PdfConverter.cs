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
        ConvertResult result = new();
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
            var images = page.GetImages()
                .OrderByDescending(x => x.Bounds.Top)
                .ThenBy(x => x.Bounds.Left)
                .ToList();

            // 行ごとの処理
            int currentLineIndex = 0;
            while (currentLineIndex < lines.Count)
            {
                var mergedWords = MergeWordsInLine(lines[currentLineIndex], horizontalTolerance);
                
                // テーブル構造の検出を試みる
                var tableData = TryDetectTable(lines, currentLineIndex, horizontalTolerance, verticalTolerance);
                
                if (tableData.IsTable)
                {
                    // テーブルが検出されたら、マークダウン形式のテーブルを作成
                    sb.AppendLine(ConvertToMarkdownTable(tableData.TableLines));
                    sb.AppendLine(); // テーブル後の空行
                    
                    // テーブルの行数分、インデックスを進める
                    currentLineIndex += tableData.RowCount;
                }
                else
                {
                    // 行内の画像を取得し、Markdown形式に変換
                    foreach (var img in images.Where(x => x.Bounds.Bottom > lines[currentLineIndex][0].BoundingBox.Bottom).ToList())
                    {
                        try
                        {
                            string base64 = ConvertImageToBase64(img.RawBytes);
                            sb.AppendLine($"![]({base64})\n");
                        }
                        catch (Exception ex)
                        {
                            result.Warnings.Add($"Error processing image: {ex.Message}");
                        }
                        
                        images.Remove(img);
                    }

                    var lineText = string.Join(" ", mergedWords.Select(x => string.Join("", x.Select(w => w.Text))));
                    sb.AppendLine(lineText + "\n");
                    
                    currentLineIndex++;
                }
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
                    result.Warnings.Add($"Error processing image: {ex.Message}");
                }
            }
        }
        result.Text = sb.ToString().Trim();

        #if DEBUG
        Console.WriteLine("PDF Conversion Result:");
        Console.WriteLine(result.Text);

        Console.WriteLine("PDF Conversion Warnings:");
        foreach (var warning in result.Warnings)
        {
            Console.WriteLine(warning);
        }
        #endif

        return result;
    }

    /// <summary>
    /// テーブル構造を検出する関数
    /// </summary>
    /// <param name="lines">すべての行データ</param>
    /// <param name="startLineIndex">検査を開始する行のインデックス</param>
    /// <param name="horizontalTolerance">水平方向の許容値</param>
    /// <param name="verticalTolerance">垂直方向の許容値</param>
    /// <returns>テーブル情報を含む（TableDetectionResult）オブジェクト</returns>
    private static TableDetectionResult TryDetectTable(List<List<Word>> lines, int startLineIndex, double horizontalTolerance, double verticalTolerance)
    {
        // 最低2行2列以上のデータが必要
        if (startLineIndex >= lines.Count - 1)
        {
            return new TableDetectionResult { IsTable = false };
        }

        // 最初の行をベースにカラム位置を特定
        var firstLineWords = MergeWordsInLine(lines[startLineIndex], horizontalTolerance);
        if (firstLineWords.Count < 2)
        {
            // 2列未満ならテーブルではない
            return new TableDetectionResult { IsTable = false };
        }

        // カラム位置（X座標の境界値）を記録
        List<double> columnPositions = [];
        foreach (var wordGroup in firstLineWords)
        {
            double midPoint = (wordGroup.First().BoundingBox.Left + wordGroup.Last().BoundingBox.Right) / 2;
            columnPositions.Add(midPoint);
        }

        // テーブル候補の行を収集
        List<List<List<Word>>> tableLines = new();
        tableLines.Add(firstLineWords);
        
        int rowCount = 1;
        for (int i = startLineIndex + 1; i < lines.Count; i++)
        {
            var nextLineWords = MergeWordsInLine(lines[i], horizontalTolerance);
            
            // 以下の条件のいずれかを満たす場合、テーブル構造でない可能性が高い
            // 1. 次の行のワード数が著しく異なる（列数が変化）
            // 2. 行間の距離が通常より大きい（テーブル行間は一定）
            if (Math.Abs(nextLineWords.Count - firstLineWords.Count) > 1)
            {
                break;
            }
            
            // 行間距離をチェック (2行目以降)
            if (tableLines.Count > 1)
            {
                var prevLineBottom = lines[i-1][0].BoundingBox.Bottom;
                var currLineBottom = lines[i][0].BoundingBox.Bottom;
                var distanceBetweenLines = Math.Abs(prevLineBottom - currLineBottom);
                
                // 最初の行間距離の基準値
                var firstLineDistance = Math.Abs(lines[startLineIndex][0].BoundingBox.Bottom - 
                                               lines[startLineIndex + 1][0].BoundingBox.Bottom);
                
                // 行間距離が基準値から大きく外れている場合は、テーブルの終了と判断
                if (Math.Abs(distanceBetweenLines - firstLineDistance) > verticalTolerance * 2)
                {
                    break;
                }
            }
            
            // 列位置の一致度を検証
            bool columnsAligned = true;
            if (nextLineWords.Count >= 2)
            {
                for (int col = 0; col < Math.Min(nextLineWords.Count, columnPositions.Count); col++)
                {
                    // 各列の中心位置がある程度一致しているか確認
                    if (col < nextLineWords.Count)
                    {
                        double midPoint = (nextLineWords[col].First().BoundingBox.Left + 
                                          nextLineWords[col].Last().BoundingBox.Right) / 2;
                        
                        if (Math.Abs(midPoint - columnPositions[col]) > horizontalTolerance * 3)
                        {
                            columnsAligned = false;
                            break;
                        }
                    }
                }
            }
            else
            {
                columnsAligned = false;
            }
            
            if (!columnsAligned)
            {
                break;
            }
            
            tableLines.Add(nextLineWords);
            rowCount++;
        }
        
        // 最低2行2列以上のデータがあればテーブルとして認識
        if (rowCount >= 2 && firstLineWords.Count >= 2)
        {
            return new TableDetectionResult
            {
                IsTable = true,
                TableLines = tableLines,
                RowCount = rowCount
            };
        }
        
        return new TableDetectionResult { IsTable = false };
    }
    
    /// <summary>
    /// テーブルデータをマークダウン形式に変換する
    /// </summary>
    /// <param name="tableData">テーブルの行と列のデータ</param>
    /// <returns>マークダウン形式のテーブル文字列</returns>
    private static string ConvertToMarkdownTable(List<List<List<Word>>> tableData)
    {
        if (tableData == null || tableData.Count == 0)
        {
            return string.Empty;
        }
        
        StringBuilder sb = new();
        int maxColumns = tableData.Max(row => row.Count);
        
        // テーブルの各行を処理
        for (int rowIndex = 0; rowIndex < tableData.Count; rowIndex++)
        {
            var row = tableData[rowIndex];
            sb.Append('|');
            
            // 各列のデータを追加
            for (int colIndex = 0; colIndex < maxColumns; colIndex++)
            {
                string cellText = colIndex < row.Count 
                    ? string.Join("", row[colIndex].Select(w => w.Text)) 
                    : "";
                    
                sb.Append($" {cellText.Trim()} |");
            }
            sb.AppendLine();
            
            // 先頭行の後にヘッダー区切り行を追加
            if (rowIndex == 0)
            {
                sb.Append('|');
                for (int i = 0; i < maxColumns; i++)
                {
                    sb.Append(" --- |");
                }
                sb.AppendLine();
            }
        }
        
        return sb.ToString();
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
    
    /// <summary>
    /// テーブル検出の結果を保持するクラス
    /// </summary>
    private class TableDetectionResult
    {
        /// <summary>
        /// テーブルが検出されたかどうか
        /// </summary>
        public bool IsTable { get; set; }
        
        /// <summary>
        /// テーブルの行数
        /// </summary>
        public int RowCount { get; set; }
        
        /// <summary>
        /// テーブルのデータ（行と列の単語グループ）
        /// </summary>
        public List<List<List<Word>>> TableLines { get; set; } = [];
    }
}