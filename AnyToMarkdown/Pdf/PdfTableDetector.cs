using UglyToad.PdfPig.Content;

namespace AnyToMarkdown.Pdf;

internal class PdfTableDetector
{
    private readonly double _horizontalTolerance;
    private readonly double _verticalTolerance;
    
    public PdfTableDetector(double horizontalTolerance = 5.0, double verticalTolerance = 5.0)
    {
        _horizontalTolerance = horizontalTolerance;
        _verticalTolerance = verticalTolerance;
    }
    
    public TableDetectionResult TryDetectTable(List<List<Word>> lines, int startLineIndex)
    {
        if (startLineIndex >= lines.Count - 1)
        {
            return new TableDetectionResult { IsTable = false };
        }

        var firstLineWords = PdfWordProcessor.MergeWordsInLine(lines[startLineIndex], _horizontalTolerance);
        if (firstLineWords.Count < 2)
        {
            return new TableDetectionResult { IsTable = false };
        }

        var columnPositions = GetColumnPositions(firstLineWords);
        var tableLines = new List<List<List<Word>>> { firstLineWords };
        
        int rowCount = 1;
        int consecutiveFailures = 0;
        
        for (int i = startLineIndex + 1; i < lines.Count; i++)
        {
            var nextLineWords = PdfWordProcessor.MergeWordsInLine(lines[i], _horizontalTolerance);
            
            // 空行をスキップ
            if (nextLineWords.Count == 0 || 
                nextLineWords.All(group => group.All(w => string.IsNullOrWhiteSpace(w.Text))))
            {
                consecutiveFailures++;
                if (consecutiveFailures >= 2) break;
                continue;
            }
            
            if (!IsValidTableRow(nextLineWords, firstLineWords.Count, lines, i, startLineIndex, columnPositions))
            {
                consecutiveFailures++;
                if (consecutiveFailures >= 2) break;
                continue;
            }
            
            consecutiveFailures = 0;
            tableLines.Add(nextLineWords);
            rowCount++;
        }
        
        // テーブル検出基準を緩和
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
    
    private List<double> GetColumnPositions(List<List<Word>> firstLineWords)
    {
        var columnPositions = new List<double>();
        foreach (var wordGroup in firstLineWords)
        {
            double midPoint = (wordGroup.First().BoundingBox.Left + wordGroup.Last().BoundingBox.Right) / 2;
            columnPositions.Add(midPoint);
        }
        return columnPositions;
    }
    
    private bool IsValidTableRow(List<List<Word>> nextLineWords, int expectedColumnCount, 
        List<List<Word>> lines, int currentIndex, int startLineIndex, List<double> columnPositions)
    {
        // 列数の検証を緩和（空のセルを許可）
        if (nextLineWords.Count == 0 || nextLineWords.Count > expectedColumnCount * 2)
        {
            return false;
        }
        
        // 行間隔チェックを緩和
        if (!IsLineSpacingConsistent(lines, currentIndex, startLineIndex))
        {
            // 行間隔が一致しなくても、列の配置が合っていればテーブル行と認識
            if (!AreColumnsAligned(nextLineWords, columnPositions))
            {
                return false;
            }
        }
        
        return true;
    }
    
    private bool IsLineSpacingConsistent(List<List<Word>> lines, int currentIndex, int startLineIndex)
    {
        if (currentIndex <= startLineIndex + 1) return true;
        
        var prevLineBottom = lines[currentIndex - 1][0].BoundingBox.Bottom;
        var currLineBottom = lines[currentIndex][0].BoundingBox.Bottom;
        var distanceBetweenLines = Math.Abs(prevLineBottom - currLineBottom);
        
        var firstLineDistance = Math.Abs(lines[startLineIndex][0].BoundingBox.Bottom - 
                                       lines[startLineIndex + 1][0].BoundingBox.Bottom);
        
        return Math.Abs(distanceBetweenLines - firstLineDistance) <= _verticalTolerance * 2;
    }
    
    private bool AreColumnsAligned(List<List<Word>> nextLineWords, List<double> columnPositions)
    {
        if (nextLineWords.Count < 2) return false;
        
        for (int col = 0; col < Math.Min(nextLineWords.Count, columnPositions.Count); col++)
        {
            if (col < nextLineWords.Count)
            {
                double midPoint = (nextLineWords[col].First().BoundingBox.Left + 
                                  nextLineWords[col].Last().BoundingBox.Right) / 2;
                
                if (Math.Abs(midPoint - columnPositions[col]) > _horizontalTolerance * 3)
                {
                    return false;
                }
            }
        }
        
        return true;
    }
}