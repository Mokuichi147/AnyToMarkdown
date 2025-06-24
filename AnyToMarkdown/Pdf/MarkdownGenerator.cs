using System.Text;
using UglyToad.PdfPig.Content;

namespace AnyToMarkdown.Pdf;

internal static class MarkdownGenerator
{
    public static string GenerateMarkdown(DocumentStructure structure)
    {
        var sb = new StringBuilder();
        var elements = structure.Elements.Where(e => e.Type != ElementType.Empty).ToList();
        
        for (int i = 0; i < elements.Count; i++)
        {
            var element = elements[i];
            var markdown = ConvertElementToMarkdown(element, elements, i);
            
            if (!string.IsNullOrWhiteSpace(markdown))
            {
                sb.AppendLine(markdown);
                
                // ヘッダーの後に空行を追加
                if (element.Type == ElementType.Header && i < elements.Count - 1)
                {
                    sb.AppendLine();
                }
            }
        }

        return PostProcessMarkdown(sb.ToString());
    }

    private static string ConvertElementToMarkdown(DocumentElement element, List<DocumentElement> allElements, int currentIndex)
    {
        return element.Type switch
        {
            ElementType.Header => ConvertHeader(element),
            ElementType.ListItem => ConvertListItem(element),
            ElementType.TableRow => ConvertTableRow(element, allElements, currentIndex),
            ElementType.Paragraph => ConvertParagraph(element),
            _ => element.Content
        };
    }

    private static string ConvertHeader(DocumentElement element)
    {
        var text = element.Content.Trim();
        
        // 既にMarkdownヘッダーの場合はそのまま
        if (text.StartsWith("#")) return text;
        
        var level = DetermineHeaderLevel(element);
        var prefix = new string('#', level);
        
        return $"{prefix} {text}";
    }

    private static int DetermineHeaderLevel(DocumentElement element)
    {
        var text = element.Content.Trim();
        
        // 明示的なMarkdownヘッダーレベル
        if (text.StartsWith("####")) return 4;
        if (text.StartsWith("###")) return 3;
        if (text.StartsWith("##")) return 2;
        if (text.StartsWith("#")) return 1;
        
        // 階層的数字パターンベース
        var hierarchicalMatch = System.Text.RegularExpressions.Regex.Match(text, @"^(\d+(\.\d+)*)");
        if (hierarchicalMatch.Success)
        {
            var parts = hierarchicalMatch.Groups[1].Value.Split('.');
            return Math.Min(parts.Length, 6); // 最大6レベル
        }
        
        // フォントサイズベースの判定（相対的評価に変更）
        if (element.FontSize > 18) return 1;
        if (element.FontSize > 16) return 2;
        if (element.FontSize > 14) return 3;
        if (element.FontSize > 12) return 4;
        
        // 短いテキストは上位レベルのヘッダーになりやすい
        if (text.Length <= 30) return 2;
        if (text.Length <= 50) return 3;
        
        // デフォルトは4レベル
        return 4;
    }

    private static string ConvertListItem(DocumentElement element)
    {
        var text = element.Content.Trim();
        
        // 既にMarkdown形式の場合はそのまま
        if (text.StartsWith("-") || text.StartsWith("*") || text.StartsWith("+"))
            return text;
            
        // 日本語の箇条書き記号を変換
        if (text.StartsWith("・"))
            return "- " + text.Substring(1).Trim();
        if (text.StartsWith("•") || text.StartsWith("◦"))
            return "- " + text.Substring(1).Trim();
            
        // 数字付きリストの修正 - スペースを確保
        if (text.Length >= 2 && char.IsDigit(text[0]) && text[1] == '.')
        {
            if (text.Length > 2 && text[2] != ' ')
                return text[0] + ". " + text.Substring(2);
            return text;
        }
            
        // 括弧付き数字を変換
        if (text.Length > 3 && text[0] == '(' && char.IsDigit(text[1]) && text[2] == ')')
            return $"{text[1]}. {text.Substring(3).Trim()}";
            
        // その他はダッシュを付ける
        return $"- {text}";
    }

    private static string ConvertTableRow(DocumentElement element, List<DocumentElement> allElements, int currentIndex)
    {
        // 連続するテーブル行を検出してMarkdownテーブルを生成
        var tableRows = new List<DocumentElement>();
        
        // 現在の行から後方のテーブル行を収集
        for (int i = currentIndex; i < allElements.Count; i++)
        {
            if (allElements[i].Type == ElementType.TableRow)
            {
                tableRows.Add(allElements[i]);
            }
            else
            {
                break;
            }
        }

        // 最初の行の場合のみテーブルを生成
        if (currentIndex == 0 || allElements[currentIndex - 1].Type != ElementType.TableRow)
        {
            return GenerateMarkdownTable(tableRows);
        }

        // 後続の行は空文字を返す（既にテーブルに含まれている）
        return "";
    }

    private static string GenerateMarkdownTable(List<DocumentElement> tableRows)
    {
        if (tableRows.Count == 0) return "";

        var sb = new StringBuilder();
        var maxColumns = 0;
        var allCells = new List<List<string>>();

        // 各行をセルに分割
        foreach (var row in tableRows)
        {
            var cells = ParseTableCells(row);
            // 空のセルを除外してより正確なテーブルを作成
            if (cells.Count > 0 && cells.Any(c => !string.IsNullOrWhiteSpace(c)))
            {
                allCells.Add(cells);
                maxColumns = Math.Max(maxColumns, cells.Count);
            }
        }

        // テーブル行が不足している場合は空文字を返す
        if (allCells.Count < 2) return "";

        // 列数の正規化：最も一般的な列数に合わせる
        var columnCounts = allCells.GroupBy(row => row.Count).OrderByDescending(g => g.Count());
        var normalizedColumnCount = columnCounts.First().Key;
        maxColumns = Math.Min(maxColumns, normalizedColumnCount + 1); // 最大1列の差を許容

        // 列数を統一
        foreach (var cells in allCells)
        {
            while (cells.Count < maxColumns)
            {
                cells.Add("");
            }
            // 余分な列は削除
            if (cells.Count > maxColumns)
            {
                cells.RemoveRange(maxColumns, cells.Count - maxColumns);
            }
        }

        // Markdownテーブルを生成
        for (int rowIndex = 0; rowIndex < allCells.Count; rowIndex++)
        {
            var cells = allCells[rowIndex];
            sb.Append("|");
            foreach (var cell in cells)
            {
                var cleanCell = cell.Replace("|", "\\|").Trim();
                if (string.IsNullOrWhiteSpace(cleanCell)) cleanCell = " ";
                sb.Append($" {cleanCell} |");
            }
            sb.AppendLine();

            // ヘッダー行の後に区切り行を追加
            if (rowIndex == 0)
            {
                sb.Append("|");
                for (int i = 0; i < maxColumns; i++)
                {
                    sb.Append(" --- |");
                }
                sb.AppendLine();
            }
        }

        return sb.ToString();
    }

    private static List<string> ParseTableCells(DocumentElement row)
    {
        var text = row.Content;
        var words = row.Words;
        
        // パイプ文字があれば既にMarkdownテーブル形式（改良版）
        if (text.Contains("|"))
        {
            var tableCells = text.Split('|')
                .Select(s => s.Trim())
                .Where(s => !string.IsNullOrEmpty(s))
                .ToList();
            
            // 改行を含むセルの処理
            for (int i = 0; i < tableCells.Count; i++)
            {
                tableCells[i] = ProcessMultiLineCell(tableCells[i]);
            }
            
            return tableCells;
        }
        
        // 単語間の大きなギャップでセルを分割
        var cells = new List<string>();
        var currentCell = new List<Word>();
        
        if (words.Count == 0)
        {
            // テキストベースのフォールバック - より賢い分割
            return SplitTextIntoCells(text);
        }

        // 単語間のギャップを計算
        var gaps = new List<double>();
        for (int i = 1; i < words.Count; i++)
        {
            gaps.Add(words[i].BoundingBox.Left - words[i-1].BoundingBox.Right);
        }

        if (gaps.Count == 0)
        {
            return SplitTextIntoCells(text);
        }

        // ダイナミックな閾値設定 - より複雑なテーブルに対応
        var avgGap = gaps.Average();
        var medianGap = gaps.OrderBy(g => g).ElementAt(gaps.Count / 2);
        var threshold = Math.Max(Math.Max(avgGap * 1.5, medianGap * 2), 15);

        currentCell.Add(words[0]);
        
        for (int i = 1; i < words.Count; i++)
        {
            var gap = words[i].BoundingBox.Left - words[i-1].BoundingBox.Right;
            
            if (gap > threshold)
            {
                // セル境界
                var cellText = string.Join("", currentCell.Select(w => w.Text)).Trim();
                if (!string.IsNullOrEmpty(cellText))
                {
                    cells.Add(cellText);
                }
                currentCell.Clear();
            }
            
            currentCell.Add(words[i]);
        }

        // 最後のセル
        if (currentCell.Count > 0)
        {
            var cellText = string.Join("", currentCell.Select(w => w.Text)).Trim();
            if (!string.IsNullOrEmpty(cellText))
            {
                cells.Add(cellText);
            }
        }

        // フォールバック: セル数が少なすぎる場合
        if (cells.Count < 2)
        {
            return SplitTextIntoCells(text);
        }

        return cells;
    }

    private static List<string> SplitTextIntoCells(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return new List<string>();
        
        // より洗練された分割ロジック
        var parts = new List<string>();
        
        // スペース、タブ、全角スペースで分割
        var candidates = text.Split(new[] { ' ', '\t', '　' }, StringSplitOptions.RemoveEmptyEntries);
        
        // 日本語のテーブルヘッダーパターンを考慮
        if (candidates.Length >= 4)
        {
            // 「項目 2024年Q2 2023年Q2 増減額 増減率」のようなパターン
            var result = new List<string>();
            var currentGroup = new List<string>();
            
            foreach (var candidate in candidates)
            {
                // 年度・数値パターンの場合は独立したセル
                if (System.Text.RegularExpressions.Regex.IsMatch(candidate, @"^\d{4}年|^\+?\-?\d+\.?\d*%?$|^[A-Z]\d+$"))
                {
                    if (currentGroup.Count > 0)
                    {
                        result.Add(string.Join(" ", currentGroup));
                        currentGroup.Clear();
                    }
                    result.Add(candidate);
                }
                else
                {
                    currentGroup.Add(candidate);
                }
            }
            
            if (currentGroup.Count > 0)
            {
                result.Add(string.Join(" ", currentGroup));
            }
            
            if (result.Count >= 2)
            {
                return result;
            }
        }
        
        // 単純な分割をフォールバック
        return candidates.ToList();
    }

    private static string ConvertParagraph(DocumentElement element)
    {
        var text = element.Content.Trim();
        
        // 強調表現の復元
        text = RestoreFormatting(text);
        
        return text;
    }
    
    private static string RestoreFormatting(string text)
    {
        // この関数は不要 - フォントベースの書式設定が既に適用されている
        return text;
    }
    
    private static string ProcessMultiLineCell(string cellContent)
    {
        if (string.IsNullOrWhiteSpace(cellContent)) return cellContent;
        
        // 改行文字を <br> タグに変換（Markdownテーブル内での改行表現）
        cellContent = cellContent.Replace("\r\n", "<br>")
                                .Replace("\n", "<br>")
                                .Replace("\r", "<br>");
        
        // 連続する <br> タグを単一に統合
        cellContent = System.Text.RegularExpressions.Regex.Replace(cellContent, @"(<br>\s*){2,}", "<br>");
        
        // パイプ文字をエスケープ
        cellContent = cellContent.Replace("|", "\\|");
        
        return cellContent.Trim();
    }

    private static string PostProcessMarkdown(string markdown)
    {
        if (string.IsNullOrWhiteSpace(markdown))
            return "";
            
        var lines = markdown.Split('\n');
        var processedLines = new List<string>();
        
        bool previousWasEmpty = false;
        
        foreach (var line in lines)
        {
            var trimmed = line.Trim();
            
            // 単独の数字行を除外（ページ番号など）
            if (trimmed.Length > 0 && trimmed.All(char.IsDigit) && trimmed.Length <= 3)
            {
                continue;
            }
            
            if (string.IsNullOrWhiteSpace(trimmed))
            {
                if (!previousWasEmpty)
                {
                    processedLines.Add("");
                    previousWasEmpty = true;
                }
            }
            else
            {
                // ヘッダーの前に空行を追加（ただし最初の行でない場合）
                if (trimmed.StartsWith("#") && processedLines.Count > 0 && !previousWasEmpty)
                {
                    processedLines.Add("");
                }
                
                processedLines.Add(trimmed);
                previousWasEmpty = false;
            }
        }
        
        // 末尾の空行を削除
        while (processedLines.Count > 0 && string.IsNullOrWhiteSpace(processedLines.Last()))
        {
            processedLines.RemoveAt(processedLines.Count - 1);
        }
        
        return string.Join("\n", processedLines);
    }
}