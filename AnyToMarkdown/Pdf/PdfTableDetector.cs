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
        
        // テーブル検出基準を厳格化（品質向上）
        if (rowCount >= 2 && firstLineWords.Count >= 2)
        {
            // テーブルの一貫性をチェック
            var consistency = CalculateTableConsistency(tableLines);
            if (consistency >= 0.6) // 60%以上の一貫性が必要
            {
                return new TableDetectionResult
                {
                    IsTable = true,
                    TableLines = tableLines,
                    RowCount = rowCount
                };
            }
        }
        
        return new TableDetectionResult { IsTable = false };
    }
    
    private double CalculateTableConsistency(List<List<List<Word>>> tableLines)
    {
        if (tableLines.Count < 2) return 0.0;
        
        // 各行の列数を比較
        var columnCounts = tableLines.Select(row => row.Count).ToList();
        var mostFrequentCount = columnCounts.GroupBy(x => x).OrderByDescending(g => g.Count()).First().Key;
        var consistentRows = columnCounts.Count(c => Math.Abs(c - mostFrequentCount) <= 1);
        
        return (double)consistentRows / tableLines.Count;
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
        if (nextLineWords.Count == 0)
        {
            return false;
        }
        
        // 列数が大幅に異なる場合は無効
        if (nextLineWords.Count > expectedColumnCount * 1.5 && expectedColumnCount > 1)
        {
            return false;
        }
        
        // 列配置の確認（より柔軟に）
        if (!AreColumnsAligned(nextLineWords, columnPositions))
        {
            // 列配置が合わない場合、最低限の要件をチェック
            if (nextLineWords.Count < Math.Max(1, expectedColumnCount / 2))
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
        if (nextLineWords.Count == 0) return false;
        if (columnPositions.Count == 0) return false;
        
        // 単一列の場合は柔軟に許可
        if (nextLineWords.Count == 1 || columnPositions.Count == 1) return true;
        
        int alignedColumns = 0;
        int columnsToCheck = Math.Min(nextLineWords.Count, columnPositions.Count);
        
        for (int col = 0; col < columnsToCheck; col++)
        {
            double midPoint = (nextLineWords[col].First().BoundingBox.Left + 
                              nextLineWords[col].Last().BoundingBox.Right) / 2;
            
            // より寛容な許容範囲
            if (Math.Abs(midPoint - columnPositions[col]) <= _horizontalTolerance * 4)
            {
                alignedColumns++;
            }
        }
        
        // 半分以上の列が配置されていれば有効
        return alignedColumns >= Math.Max(1, columnsToCheck / 2);
    }
}