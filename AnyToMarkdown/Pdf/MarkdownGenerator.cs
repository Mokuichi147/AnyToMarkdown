using System.Text;
using UglyToad.PdfPig.Content;

namespace AnyToMarkdown.Pdf;

internal static class MarkdownGenerator
{
    public static string GenerateMarkdown(DocumentStructure structure)
    {
        var sb = new StringBuilder();
        var elements = structure.Elements.Where(e => e.Type != ElementType.Empty).ToList();
        
        // 段落の統合処理
        var consolidatedElements = ConsolidateParagraphs(elements);
        
        for (int i = 0; i < consolidatedElements.Count; i++)
        {
            var element = consolidatedElements[i];
            var markdown = ConvertElementToMarkdown(element, consolidatedElements, i, structure.FontAnalysis);
            
            if (!string.IsNullOrWhiteSpace(markdown))
            {
                sb.AppendLine(markdown);
                
                // 要素間の適切な間隔を追加
                if (i < consolidatedElements.Count - 1)
                {
                    var nextElement = consolidatedElements[i + 1];
                    
                    // ヘッダーの後に空行を追加
                    if (element.Type == ElementType.Header)
                    {
                        sb.AppendLine();
                    }
                    // 段落の後に別の段落、テーブル、リストが来る場合も空行を追加
                    else if (element.Type == ElementType.Paragraph && 
                            (nextElement.Type == ElementType.Paragraph || nextElement.Type == ElementType.TableRow || nextElement.Type == ElementType.ListItem))
                    {
                        sb.AppendLine();
                    }
                }
            }
        }

        return PostProcessMarkdown(sb.ToString());
    }

    private static List<DocumentElement> ConsolidateParagraphs(List<DocumentElement> elements)
    {
        var consolidated = new List<DocumentElement>();
        
        for (int i = 0; i < elements.Count; i++)
        {
            var current = elements[i];
            
            if (current.Type == ElementType.Paragraph)
            {
                var paragraphBuilder = new StringBuilder(current.Content);
                var consolidatedWords = new List<Word>(current.Words);
                
                // より慎重な段落統合：条件を満たす場合のみ統合
                int j = i + 1;
                while (j < elements.Count && elements[j].Type == ElementType.Paragraph)
                {
                    var nextParagraph = elements[j];
                    
                    // 統合条件をチェック
                    if (!ShouldConsolidateParagraphs(current, nextParagraph))
                    {
                        break;
                    }
                    
                    paragraphBuilder.Append(" ").Append(nextParagraph.Content);
                    consolidatedWords.AddRange(nextParagraph.Words);
                    j++;
                }
                
                // 統合された段落要素を作成
                var consolidatedElement = new DocumentElement
                {
                    Type = ElementType.Paragraph,
                    Content = paragraphBuilder.ToString(),
                    FontSize = current.FontSize,
                    LeftMargin = current.LeftMargin,
                    IsIndented = current.IsIndented,
                    Words = consolidatedWords
                };
                
                consolidated.Add(consolidatedElement);
                i = j - 1; // 統合した要素数分進める
            }
            else
            {
                consolidated.Add(current);
            }
        }
        
        return consolidated;
    }
    
    private static bool ShouldConsolidateParagraphs(DocumentElement current, DocumentElement next)
    {
        // 書式設定が大きく異なる場合は統合しない
        if (Math.Abs(current.FontSize - next.FontSize) > 1.0)
        {
            return false;
        }
        
        // インデント状態が異なる場合は統合しない
        if (current.IsIndented != next.IsIndented)
        {
            return false;
        }
        
        // マージンが大きく異なる場合は統合しない
        if (Math.Abs(current.LeftMargin - next.LeftMargin) > 10.0)
        {
            return false;
        }
        
        // 書式設定マークを含む場合は慎重に判断
        var currentText = current.Content.Trim();
        var nextText = next.Content.Trim();
        
        // 現在の段落が完結している（句点で終わる）場合は統合しない
        if (currentText.EndsWith("。") || currentText.EndsWith("."))
        {
            return false;
        }
        
        // 次の段落が書式設定を含む場合は統合しない
        if (nextText.Contains("**") || nextText.Contains("*") || nextText.Contains("_"))
        {
            return false;
        }
        
        return true;
    }

    private static string ConvertElementToMarkdown(DocumentElement element, List<DocumentElement> allElements, int currentIndex, FontAnalysis fontAnalysis)
    {
        return element.Type switch
        {
            ElementType.Header => ConvertHeader(element, fontAnalysis),
            ElementType.ListItem => ConvertListItem(element),
            ElementType.TableRow => ConvertTableRow(element, allElements, currentIndex),
            ElementType.CodeBlock => ConvertCodeBlock(element, allElements, currentIndex),
            ElementType.QuoteBlock => ConvertQuoteBlock(element, allElements, currentIndex),
            ElementType.Paragraph => ConvertParagraph(element),
            _ => element.Content
        };
    }

    private static string ConvertHeader(DocumentElement element, FontAnalysis fontAnalysis)
    {
        var text = element.Content.Trim();
        
        // 既にMarkdownヘッダーの場合はそのまま
        if (text.StartsWith("#")) return text;
        
        // ヘッダーのフォーマットを除去してクリーンなテキストを取得
        var cleanText = StripMarkdownFormatting(text);
        
        var level = DetermineHeaderLevel(element, fontAnalysis);
        var prefix = new string('#', level);
        
        return $"{prefix} {cleanText}";
    }
    
    private static string StripMarkdownFormatting(string text)
    {
        // 太字フォーマットを除去
        text = System.Text.RegularExpressions.Regex.Replace(text, @"\*{1,3}([^*]+)\*{1,3}", "$1");
        
        // 斜体フォーマットを除去  
        text = System.Text.RegularExpressions.Regex.Replace(text, @"_([^_]+)_", "$1");
        
        return text.Trim();
    }

    private static int DetermineHeaderLevel(DocumentElement element, FontAnalysis fontAnalysis)
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
            return Math.Min(parts.Length, 4);
        }
        
        // フォント分析に基づく相対的なレベル判定
        var fontSizeRatio = element.FontSize / fontAnalysis.BaseFontSize;
        
        // すべてのフォントサイズから相対的な位置を計算
        var allSizes = fontAnalysis.AllFontSizes.Distinct().OrderByDescending(s => s).ToList();
        if (allSizes.Count > 0)
        {
            var currentSizeRank = allSizes.IndexOf(element.FontSize);
            if (currentSizeRank >= 0)
            {
                // フォントサイズの順位に基づいてレベルを決定
                var normalizedRank = (double)currentSizeRank / Math.Max(allSizes.Count - 1, 1);
                
                if (normalizedRank <= 0.2) return 1;  // 上位20%
                if (normalizedRank <= 0.4) return 2;  // 上位40%
                if (normalizedRank <= 0.6) return 3;  // 上位60%
                return 4;
            }
        }
        
        // フォールバック：フォントサイズ比に基づく判定
        if (fontSizeRatio >= 1.3) return 1;
        if (fontSizeRatio >= 1.2) return 2;
        if (fontSizeRatio >= 1.1) return 3;
        
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
            
        // 数字付きリストの処理（より柔軟に）
        var numberListMatch = System.Text.RegularExpressions.Regex.Match(text, @"^(\d{1,3})[\.\)](.*)");
        if (numberListMatch.Success)
        {
            var number = numberListMatch.Groups[1].Value;
            var content = numberListMatch.Groups[2].Value.Trim();
            return $"{number}. {content}";
        }
            
        // 括弧付き数字を変換
        var parenNumberMatch = System.Text.RegularExpressions.Regex.Match(text, @"^\((\d{1,3})\)(.*)");
        if (parenNumberMatch.Success)
        {
            var number = parenNumberMatch.Groups[1].Value;
            var content = parenNumberMatch.Groups[2].Value.Trim();
            return $"{number}. {content}";
        }
            
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

        // 各行をセルに分割（基本的な処理に戻す）
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
        if (allCells.Count < 1) return "";

        // 列数の正規化：最も一般的な列数に合わせる（より柔軟に）
        if (allCells.Count > 0)
        {
            var columnCounts = allCells.GroupBy(row => row.Count).OrderByDescending(g => g.Count());
            var normalizedColumnCount = columnCounts.First().Key;
            
            // 少なくとも2列は確保し、最大列数との差を2列まで許容
            maxColumns = Math.Max(Math.Min(maxColumns, normalizedColumnCount + 2), 2);
        }

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

        // 改良された適応的閾値設定
        var avgGap = gaps.Average();
        var sortedGaps = gaps.OrderBy(g => g).ToList();
        var medianGap = sortedGaps[sortedGaps.Count / 2];
        
        // より正確な閾値計算：大きなギャップと小さなギャップを区別
        var q75 = sortedGaps[(int)(sortedGaps.Count * 0.75)];
        var q25 = sortedGaps[(int)(sortedGaps.Count * 0.25)];
        var iqr = q75 - q25;
        
        // IQRベースの閾値またはメディアンベースの閾値の大きい方を使用
        var iqrThreshold = q75 + (iqr * 0.5);
        var medianThreshold = medianGap * 1.5;
        var threshold = Math.Max(Math.Max(iqrThreshold, medianThreshold), 15);

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
    
    private static string ConvertCodeBlock(DocumentElement element, List<DocumentElement> allElements, int currentIndex)
    {
        // 連続するコードブロック行を検出してまとめる
        var codeLines = new List<DocumentElement>();
        
        // 現在の行から後方のコードブロック行を収集
        for (int i = currentIndex; i < allElements.Count; i++)
        {
            if (allElements[i].Type == ElementType.CodeBlock)
            {
                codeLines.Add(allElements[i]);
            }
            else
            {
                break;
            }
        }

        // 最初の行の場合のみコードブロックを生成
        if (currentIndex == 0 || allElements[currentIndex - 1].Type != ElementType.CodeBlock)
        {
            return GenerateMarkdownCodeBlock(codeLines);
        }

        // 後続の行は空文字を返す（既にコードブロックに含まれている）
        return "";
    }
    
    private static string ConvertQuoteBlock(DocumentElement element, List<DocumentElement> allElements, int currentIndex)
    {
        // 連続する引用ブロック行を検出してまとめる
        var quoteLines = new List<DocumentElement>();
        
        // 現在の行から後方の引用ブロック行を収集
        for (int i = currentIndex; i < allElements.Count; i++)
        {
            if (allElements[i].Type == ElementType.QuoteBlock)
            {
                quoteLines.Add(allElements[i]);
            }
            else
            {
                break;
            }
        }

        // 最初の行の場合のみ引用ブロックを生成
        if (currentIndex == 0 || allElements[currentIndex - 1].Type != ElementType.QuoteBlock)
        {
            return GenerateMarkdownQuoteBlock(quoteLines);
        }

        // 後続の行は空文字を返す（既に引用ブロックに含まれている）
        return "";
    }
    
    private static string GenerateMarkdownCodeBlock(List<DocumentElement> codeLines)
    {
        if (codeLines.Count == 0) return "";
        
        var sb = new StringBuilder();
        
        // 言語の検出
        string language = DetectCodeLanguage(codeLines);
        
        sb.AppendLine($"```{language}");
        
        foreach (var line in codeLines)
        {
            var content = line.Content.Trim();
            
            // 既に``` で囲まれている場合は除去
            if (content.StartsWith("```")) 
            {
                content = content.Substring(3).Trim();
            }
            if (content.EndsWith("```"))
            {
                content = content.Substring(0, content.Length - 3).Trim();
            }
            
            sb.AppendLine(content);
        }
        
        sb.AppendLine("```");
        
        return sb.ToString();
    }
    
    private static string GenerateMarkdownQuoteBlock(List<DocumentElement> quoteLines)
    {
        if (quoteLines.Count == 0) return "";
        
        var sb = new StringBuilder();
        
        foreach (var line in quoteLines)
        {
            var content = line.Content.Trim();
            
            // 既に > で始まっている場合はそのまま
            if (content.StartsWith(">"))
            {
                sb.AppendLine(content);
            }
            else
            {
                sb.AppendLine($"> {content}");
            }
        }
        
        return sb.ToString();
    }
    
    private static string DetectCodeLanguage(List<DocumentElement> codeLines)
    {
        if (codeLines.Count == 0) return "";
        
        var allText = string.Join(" ", codeLines.Select(l => l.Content));
        
        // Python
        if (allText.Contains("def ") || allText.Contains("import ") || allText.Contains("from ") || allText.Contains("python"))
            return "python";
            
        // JavaScript/JSON
        if (allText.Contains("function") || allText.Contains("const ") || allText.Contains("let ") || allText.Contains("var "))
            return "javascript";
            
        // JSON
        if ((allText.Contains("{") && allText.Contains("}")) || allText.Contains("\"key\":"))
            return "json";
            
        // Bash/Shell
        if (allText.Contains("#!/bin/bash") || allText.Contains("sudo ") || allText.Contains("apt-get") || allText.Contains("yum "))
            return "bash";
            
        // C#
        if (allText.Contains("using ") || allText.Contains("namespace ") || allText.Contains("public class"))
            return "csharp";
            
        // HTML
        if (allText.Contains("<html") || allText.Contains("</html>") || allText.Contains("<div"))
            return "html";
            
        // CSS
        if (allText.Contains("{") && allText.Contains("}") && allText.Contains(":"))
            return "css";
        
        return "";
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
        
        // 既存の <br> タグを保持
        cellContent = System.Text.RegularExpressions.Regex.Replace(cellContent, @"<br>\s*<br>", "<br>");
        
        // テーブル内で改行が必要な場合の代替表現も対応
        // 長いテキストを自動的に改行に変換する
        if (cellContent.Length > 50 && !cellContent.Contains("<br>"))
        {
            // 文の区切りで自動改行を検出
            cellContent = System.Text.RegularExpressions.Regex.Replace(cellContent, @"([。！？])\s*([^\s])", "$1<br>$2");
        }
        
        // パイプ文字をエスケープ
        cellContent = cellContent.Replace("|", "\\|");
        
        return cellContent.Trim();
    }
    
    private static List<List<string>> ProcessMultiRowCells(List<DocumentElement> tableRows)
    {
        var allCells = new List<List<string>>();
        
        foreach (var row in tableRows)
        {
            var cells = ParseTableCells(row);
            allCells.Add(cells);
        }
        
        if (allCells.Count <= 1) return allCells;
        
        // 複数行セルの検出と統合
        var mergedCells = new List<List<string>>();
        var i = 0;
        
        while (i < allCells.Count)
        {
            var currentRow = allCells[i];
            var mergedRow = new List<string>(currentRow);
            
            // 次の行が現在の行と統合できるかチェック
            if (i + 1 < allCells.Count)
            {
                var nextRow = allCells[i + 1];
                
                // セル数が一致しない、または明らかに別のテーブル行である場合は統合しない
                if (ShouldMergeRows(currentRow, nextRow))
                {
                    // セルの内容を統合
                    for (int j = 0; j < Math.Min(mergedRow.Count, nextRow.Count); j++)
                    {
                        if (!string.IsNullOrWhiteSpace(nextRow[j]))
                        {
                            if (!string.IsNullOrWhiteSpace(mergedRow[j]))
                            {
                                mergedRow[j] += "<br>" + nextRow[j];
                            }
                            else
                            {
                                mergedRow[j] = nextRow[j];
                            }
                        }
                    }
                    i += 2; // 2行をスキップ
                }
                else
                {
                    i++;
                }
            }
            else
            {
                i++;
            }
            
            // 改行を含むセルを処理
            for (int j = 0; j < mergedRow.Count; j++)
            {
                mergedRow[j] = ProcessMultiLineCell(mergedRow[j]);
            }
            
            mergedCells.Add(mergedRow);
        }
        
        return mergedCells;
    }
    
    private static bool ShouldMergeRows(List<string> currentRow, List<string> nextRow)
    {
        // 基本的なヒューリスティック：
        // 1. 次の行の最初のセルが空で、他のセルに内容がある場合は統合
        // 2. 両方の行のセル数が同じで、次の行が明らかに継続内容の場合は統合
        
        if (nextRow.Count == 0) return false;
        
        // 次の行の最初のセルが空で、残りに内容がある場合（典型的な複数行セル）
        if (string.IsNullOrWhiteSpace(nextRow[0]) && nextRow.Skip(1).Any(c => !string.IsNullOrWhiteSpace(c)))
        {
            return true;
        }
        
        // セル数が一致し、次の行の内容が短い場合（継続行の可能性）
        if (currentRow.Count == nextRow.Count)
        {
            var avgCurrentLength = currentRow.Where(c => !string.IsNullOrWhiteSpace(c)).Average(c => c.Length);
            var avgNextLength = nextRow.Where(c => !string.IsNullOrWhiteSpace(c)).Average(c => c.Length);
            
            // 次の行の平均文字数が現在の行の半分以下の場合は継続行と判定
            return avgNextLength < avgCurrentLength * 0.5;
        }
        
        return false;
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
            
            // ヘッダー形式の単独数字も除外（# 1など）
            if (System.Text.RegularExpressions.Regex.IsMatch(trimmed, @"^#{1,6}\s*\d{1,3}$"))
            {
                continue;
            }
            
            // 非常に短い断片的なテキストを除外
            if (trimmed.Length > 0 && trimmed.Length <= 2 && !trimmed.StartsWith("#"))
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