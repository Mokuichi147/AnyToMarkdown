using System.Text;
using UglyToad.PdfPig.Content;

namespace AnyToMarkdown.Pdf;

internal static class PdfTableFormatter
{
    public static string ConvertToMarkdownTable(List<List<List<Word>>> tableData)
    {
        if (tableData == null || tableData.Count == 0)
        {
            return string.Empty;
        }
        
        var sb = new StringBuilder();
        int maxColumns = tableData.Max(row => row.Count);
        
        var columnAlignments = DetermineColumnAlignments(tableData, maxColumns);
        
        for (int rowIndex = 0; rowIndex < tableData.Count; rowIndex++)
        {
            var row = tableData[rowIndex];
            sb.Append('|');
            
            for (int colIndex = 0; colIndex < maxColumns; colIndex++)
            {
                string cellText = colIndex < row.Count 
                    ? string.Join("", row[colIndex].Select(w => w.Text)) 
                    : "";
                    
                sb.Append($" {cellText.Trim()} |");
            }
            sb.AppendLine();
            
            if (rowIndex == 0)
            {
                sb.Append('|');
                for (int i = 0; i < maxColumns; i++)
                {
                    string alignmentMark = columnAlignments[i] switch
                    {
                        1 => " :---: |",
                        2 => " ---: |", 
                        _ => " --- |"   
                    };
                    sb.Append(alignmentMark);
                }
                sb.AppendLine();
            }
        }
        
        return sb.ToString();
    }
    
    private static int[] DetermineColumnAlignments(List<List<List<Word>>> tableData, int columnCount)
    {
        int[] alignments = new int[columnCount];
        
        for (int colIndex = 0; colIndex < columnCount; colIndex++)
        {
            var cellPositions = new List<(double left, double right, double width, double center)>();
            
            foreach (var row in tableData)
            {
                if (colIndex < row.Count && row[colIndex].Count > 0)
                {
                    var wordGroup = row[colIndex];
                    double left = wordGroup.First().BoundingBox.Left;
                    double right = wordGroup.Last().BoundingBox.Right;
                    double width = right - left;
                    double center = left + (width / 2);
                    
                    cellPositions.Add((left, right, width, center));
                }
            }
            
            if (cellPositions.Count < 2)
            {
                alignments[colIndex] = 0;
                continue;
            }
            
            double avgLeft = cellPositions.Average(p => p.left);
            double avgRight = cellPositions.Average(p => p.right);
            double avgCenter = cellPositions.Average(p => p.center);
            
            double leftVariance = cellPositions.Average(p => Math.Pow(p.left - avgLeft, 2));
            double rightVariance = cellPositions.Average(p => Math.Pow(p.right - avgRight, 2));
            double centerVariance = cellPositions.Average(p => Math.Pow(p.center - avgCenter, 2));
            
            if (rightVariance < leftVariance && rightVariance < centerVariance)
            {
                alignments[colIndex] = 2;
            }
            else if (centerVariance < leftVariance && centerVariance < rightVariance)
            {
                alignments[colIndex] = 1;
            }
            else
            {
                alignments[colIndex] = 0;
            }
        }
        
        return alignments;
    }
}