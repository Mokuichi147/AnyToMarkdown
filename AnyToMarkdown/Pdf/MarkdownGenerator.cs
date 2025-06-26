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
        // 太字フォーマットを除去（複数回実行して入れ子や複数パターンを処理）
        while (text.Contains("**") || text.Contains("*"))
        {
            var before = text;
            text = System.Text.RegularExpressions.Regex.Replace(text, @"\*{1,3}([^*]*)\*{1,3}", "$1");
            if (before == text) break; // 変化がなければ終了
        }
        
        // 斜体フォーマットを除去  
        text = System.Text.RegularExpressions.Regex.Replace(text, @"_([^_]+)_", "$1");
        
        // 余分なスペースを統合
        text = System.Text.RegularExpressions.Regex.Replace(text, @"\s+", " ");
        
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
                // フォントサイズの順位に基づいてレベルを決定（より一般的なレベルを優先）
                var normalizedRank = (double)currentSizeRank / Math.Max(allSizes.Count - 1, 1);
                
                if (normalizedRank <= 0.2) return 1;  // 上位20%のみレベル1
                if (normalizedRank <= 0.6) return 2;  // 上位60%はレベル2
                if (normalizedRank <= 0.8) return 3;  // 上位80%はレベル3
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
            return "- " + text.Substring(1).Trim().Replace("\0", "");
        if (text.StartsWith("•") || text.StartsWith("◦"))
            return "- " + text.Substring(1).Trim().Replace("\0", "");
            
        // 数字付きリストの処理（より柔軟に）
        var numberListMatch = System.Text.RegularExpressions.Regex.Match(text, @"^(\d{1,3})[\.\)](.*)");
        if (numberListMatch.Success)
        {
            var number = numberListMatch.Groups[1].Value;
            var content = numberListMatch.Groups[2].Value.Trim();
            // null文字を除去
            content = content.Replace("\0", "");
            return $"{number}. {content}";
        }
            
        // 括弧付き数字を変換
        var parenNumberMatch = System.Text.RegularExpressions.Regex.Match(text, @"^\((\d{1,3})\)(.*)");
        if (parenNumberMatch.Success)
        {
            var number = parenNumberMatch.Groups[1].Value;
            var content = parenNumberMatch.Groups[2].Value.Trim();
            // null文字を除去
            content = content.Replace("\0", "");
            return $"{number}. {content}";
        }
            
        // その他はダッシュを付ける
        text = text.Replace("\0", "");
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

        // 複数行セルの検出と統合
        var processedCells = ProcessMultiRowCells(tableRows);
        
        // 各行をセルに分割（改良版）
        foreach (var cells in processedCells)
        {
            // 空のセルを除外してより正確なテーブルを作成
            if (cells.Count > 0 && cells.Any(c => !string.IsNullOrWhiteSpace(c)))
            {
                // セル内の<br>を適切に処理
                var processedCellRow = cells.Select(ProcessMultiLineCell).ToList();
                allCells.Add(processedCellRow);
                maxColumns = Math.Max(maxColumns, processedCellRow.Count);
            }
        }

        // テーブル行が不足している場合は空文字を返す
        if (allCells.Count < 1) return "";

        // より堅牢な列数正規化
        if (allCells.Count > 0)
        {
            // 列数の統計的分析
            var columnCounts = allCells.GroupBy(row => row.Count).OrderByDescending(g => g.Count());
            var mostFrequentColumnCount = columnCounts.First().Key;
            var secondMostFrequentCount = columnCounts.Count() > 1 ? columnCounts.Skip(1).First().Key : mostFrequentColumnCount;
            
            // 最頻値とその次を考慮して適応的に決定
            var targetColumnCount = mostFrequentColumnCount;
            
            // 非常に少ない行数の場合は最大列数を優先
            if (allCells.Count <= 2)
            {
                targetColumnCount = Math.Max(maxColumns, mostFrequentColumnCount);
            }
            // 複数の異なる列数がある場合は、より大きい値を選択（データ損失を避ける）
            else if (Math.Abs(mostFrequentColumnCount - secondMostFrequentCount) <= 1)
            {
                targetColumnCount = Math.Max(mostFrequentColumnCount, secondMostFrequentCount);
            }
            
            maxColumns = Math.Max(targetColumnCount, 2);
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
                if (string.IsNullOrWhiteSpace(cleanCell)) cleanCell = "";
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
        
        // 全ての置換文字を事前に除去
        text = text.Replace("￿", "").Replace("\uFFFD", "").Trim();
        
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

        // 統計的に堅牢な適応的閾値設定
        var sortedGaps = gaps.OrderBy(g => g).ToList();
        var medianGap = sortedGaps[sortedGaps.Count / 2];
        
        // 四分位数による外れ値検出
        var q1 = sortedGaps[(int)(sortedGaps.Count * 0.25)];
        var q3 = sortedGaps[(int)(sortedGaps.Count * 0.75)];
        var iqr = q3 - q1;
        
        // Tukey's fence法による閾値設定（外れ値検出の標準手法）
        var lowerFence = q1 - (1.5 * iqr);
        var upperFence = q3 + (1.5 * iqr);
        
        // 通常ギャップと大きなギャップを区別する閾値
        var threshold = Math.Max(upperFence, medianGap * 2.0);
        
        // 最小閾値を設定（非常に小さいテキストでも機能するように）
        threshold = Math.Max(threshold, 10);

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
        
        // より洗練された分割ロジック - 財務データパターンに特化
        var parts = new List<string>();
        
        // 密集した複合数値の特別処理
        var digitRatio = (double)text.Count(char.IsDigit) / text.Length;
        
        // 汎用的な数値パターン分離（言語・ドメイン非依存）
        var genericNumberPattern = @"(\d{1,4}[,.]?\d{0,3}%?|\+?-?\d{1,4}[,.]?\d{0,3}|\d{4}年|\w+\d+)";
        var numberMatches = System.Text.RegularExpressions.Regex.Matches(text, genericNumberPattern);
        
        if (numberMatches.Count >= 3 && digitRatio > 0.3 && text.Length > 15)
        {
            var extractedParts = new List<string>();
            var lastIndex = 0;
            
            foreach (System.Text.RegularExpressions.Match match in numberMatches)
            {
                // マッチ前のテキスト部分を追加
                if (match.Index > lastIndex)
                {
                    var beforeText = text.Substring(lastIndex, match.Index - lastIndex).Trim();
                    if (!string.IsNullOrWhiteSpace(beforeText))
                    {
                        extractedParts.Add(beforeText);
                    }
                }
                
                // マッチした数値部分を追加
                extractedParts.Add(match.Value);
                lastIndex = match.Index + match.Length;
            }
            
            // 最後の残りテキストを追加
            if (lastIndex < text.Length)
            {
                var remainingText = text.Substring(lastIndex).Trim();
                if (!string.IsNullOrWhiteSpace(remainingText))
                {
                    extractedParts.Add(remainingText);
                }
            }
            
            if (extractedParts.Count >= 3)
            {
                return extractedParts.Where(p => !string.IsNullOrWhiteSpace(p)).ToList();
            }
        }
        
        // フォールバック: 通常の分割
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
        
        // 禁止文字（置換文字）を除去
        cellContent = cellContent.Replace("￿", "").Replace("\uFFFD", "");
        
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
        // 1. 次の行の最初のセルが空、継続マーカー、または禁止文字で、他のセルに内容がある場合は統合
        // 2. 両方の行のセル数が同じで、次の行が明らかに継続内容の場合は統合
        
        if (nextRow.Count == 0) return false;
        
        var firstCell = nextRow[0]?.Replace("￿", "").Replace("\uFFFD", "").Trim() ?? "";
        
        // 改良された継続行判定：日本語セルの分離パターンを検出
        var isFirstCellEmpty = string.IsNullOrWhiteSpace(firstCell);
        
        // 次の行の内容分布を分析
        var nonEmptyNextCells = nextRow.Where(c => !string.IsNullOrWhiteSpace(c?.Replace("￿", "").Replace("\uFFFD", "").Trim())).Count();
        var nextCellLengths = nextRow.Select(c => (c?.Replace("￿", "").Replace("\uFFFD", "").Trim() ?? "").Length).ToList();
        
        // 前の行の内容分布を分析
        var nonEmptyCurrentCells = currentRow.Where(c => !string.IsNullOrWhiteSpace(c?.Replace("￿", "").Replace("\uFFFD", "").Trim())).Count();
        
        // 次の行が継続行である可能性の判定
        // 1. 最初のセルが空で、他のセルに短いコンテンツが含まれている
        // 2. 行全体のセル数が少ない（継続コンテンツの特徴）
        bool isPotentialContinuation = isFirstCellEmpty && nonEmptyNextCells > 0 && nonEmptyNextCells < currentRow.Count;
        
        // 短い断片的なコンテンツパターンを検出（日本語での単語分離）
        bool hasFragmentedContent = nextCellLengths.Where(l => l > 0).All(l => l < 10) && 
                                   nextCellLengths.Where(l => l > 0).Count() <= 2;
        
        var hasContentInOtherCells = nextRow.Skip(1).Any(c => 
        {
            var cleanCell = c?.Replace("￿", "").Replace("\uFFFD", "").Trim() ?? "";
            return !string.IsNullOrWhiteSpace(cleanCell) && cleanCell.Length > 0;
        });
        
        // 多段階の統合判定
        
        // パターン1: 最初のセルが空で、他のセルに継続コンテンツがある
        if (isPotentialContinuation && hasContentInOtherCells)
        {
            return true;
        }
        
        // パターン2: 断片的なコンテンツの継続（文字数制限で分割された可能性）
        if (hasFragmentedContent && isFirstCellEmpty)
        {
            return true;
        }
        
        // パターン3: セル数が一致し、次の行の内容が短い場合（継続行の可能性）
        if (currentRow.Count == nextRow.Count && nonEmptyNextCells > 0)
        {
            var avgCurrentLength = currentRow.Where(c => !string.IsNullOrWhiteSpace(c?.Trim())).DefaultIfEmpty("").Average(c => c.Length);
            var avgNextLength = nextRow.Where(c => !string.IsNullOrWhiteSpace(c?.Trim())).DefaultIfEmpty("").Average(c => c.Length);
            
            // 次の行の平均文字数が現在の行の半分以下で、短いテキストの場合は継続行と判定
            if (avgNextLength > 0 && avgCurrentLength > 0 && avgNextLength < avgCurrentLength * 0.6 && avgNextLength < 15)
            {
                return true;
            }
        }
        
        return false;
    }

    private static string PostProcessMarkdown(string markdown)
    {
        if (string.IsNullOrWhiteSpace(markdown))
            return "";
        
        // 禁止文字（null文字など）を除去
        markdown = RemoveForbiddenCharacters(markdown);
            
        var lines = markdown.Split('\n');
        
        // テーブルの分離された行を統合する前処理
        lines = MergeDisconnectedTableCells(lines);
        
        // 重複するテーブル区切り行を除去
        lines = RemoveDuplicateTableSeparators(lines);
        
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
    
    private static string RemoveForbiddenCharacters(string text)
    {
        if (string.IsNullOrEmpty(text)) return text;
        
        // null文字（U+0000）を除去
        text = text.Replace("\0", "");
        
        // 置換文字（U+FFFD, ￿）を除去 - これは文字化けを表す
        text = text.Replace("￿", "");
        text = text.Replace("\uFFFD", "");
        
        // その他の制御文字を除去（印刷可能文字、空白、改行、タブ以外）
        var cleanedText = new StringBuilder();
        foreach (char c in text)
        {
            // 印刷可能文字、空白類、改行、タブを保持
            if (char.IsControl(c))
            {
                if (c == '\n' || c == '\r' || c == '\t')
                {
                    cleanedText.Append(c);
                }
                // その他の制御文字は除去
            }
            else
            {
                cleanedText.Append(c);
            }
        }
        
        return cleanedText.ToString();
    }

    private static string[] MergeDisconnectedTableCells(string[] lines)
    {
        var result = new List<string>();
        
        for (int i = 0; i < lines.Length; i++)
        {
            var line = lines[i].Trim();
            
            // テーブル行（|で始まり終わる）の直後にある分離されたテキストを検出
            if (line.StartsWith("|") && line.EndsWith("|") && i + 1 < lines.Length)
            {
                var nextLines = new List<string>();
                int j = i + 1;
                
                // 空行をスキップ
                while (j < lines.Length && string.IsNullOrWhiteSpace(lines[j].Trim()))
                {
                    j++;
                }
                
                // 連続する非テーブル行（断片化されたセル内容の可能性）を収集
                while (j < lines.Length)
                {
                    var nextLine = lines[j].Trim();
                    
                    // ヘッダー、別のテーブル、区切り行に遭遇したら終了
                    if (nextLine.StartsWith("#") ||
                        nextLine.StartsWith("|") ||
                        nextLine.Contains("---") ||
                        nextLine.All(c => c == '-' || c == ' '))
                    {
                        break;
                    }
                    
                    // 空行の場合は収集を続けるが、連続する空行は終了
                    if (string.IsNullOrWhiteSpace(nextLine))
                    {
                        // 次の行が空でない場合は継続、そうでなければ終了
                        if (j + 1 < lines.Length && !string.IsNullOrWhiteSpace(lines[j + 1].Trim()))
                        {
                            j++;
                            continue;
                        }
                        else
                        {
                            break;
                        }
                    }
                    
                    // テキスト行を収集（より寛容な条件）
                    if (nextLine.Length > 0 && nextLine.Length <= 50 && !nextLine.Contains("|"))
                    {
                        nextLines.Add(nextLine);
                        j++;
                    }
                    else if (nextLine.Length > 50)
                    {
                        // 長いテキストも1行だけ収集
                        nextLines.Add(nextLine);
                        j++;
                        break;
                    }
                    else
                    {
                        break;
                    }
                    
                    // 最大5行まで統合（より多くの行を許可）
                    if (nextLines.Count >= 5) break;
                }
                
                // 分離されたテキストがある場合は、前のテーブル行のセルに統合
                if (nextLines.Count > 0)
                {
                    var enhancedTableRow = MergeFragmentsIntoTableRow(line, nextLines);
                    result.Add(enhancedTableRow);
                    
                    // 空行を追加してスキップした行数分進める
                    while (i + 1 < j && i + 1 < lines.Length)
                    {
                        i++;
                        if (string.IsNullOrWhiteSpace(lines[i].Trim()))
                        {
                            // 空行は維持
                            if (result.Count > 0 && !string.IsNullOrWhiteSpace(result.Last()))
                            {
                                result.Add("");
                            }
                        }
                    }
                    i = j - 1; // 処理した行まで進める
                }
                else
                {
                    result.Add(lines[i]);
                }
            }
            else
            {
                result.Add(lines[i]);
            }
        }
        
        return result.ToArray();
    }
    
    private static string MergeFragmentsIntoTableRow(string tableRow, List<string> fragments)
    {
        if (fragments.Count == 0) return tableRow;
        
        // テーブル行をセルに分割
        var cells = tableRow.Split('|').ToList();
        if (cells.Count < 3) return tableRow; // 無効なテーブル行
        
        // 最初と最後の空要素を除去
        if (cells.Count > 0 && string.IsNullOrWhiteSpace(cells[0])) cells.RemoveAt(0);
        if (cells.Count > 0 && string.IsNullOrWhiteSpace(cells[cells.Count - 1])) cells.RemoveAt(cells.Count - 1);
        
        // 全ての断片を結合して説明セル（通常2番目のセル）に配置
        if (fragments.Count > 0 && cells.Count >= 2)
        {
            var combinedFragments = string.Join(" ", fragments.Select(f => f.Trim()).Where(f => !string.IsNullOrWhiteSpace(f)));
            
            if (!string.IsNullOrWhiteSpace(combinedFragments))
            {
                // 説明セル（インデックス1）に統合
                var descriptionCellIndex = 1;
                
                if (!string.IsNullOrWhiteSpace(cells[descriptionCellIndex].Trim()))
                {
                    // 既存の内容がある場合は<br>で結合
                    cells[descriptionCellIndex] = " " + cells[descriptionCellIndex].Trim() + "<br>" + combinedFragments + " ";
                }
                else
                {
                    // 空のセルの場合は直接配置
                    cells[descriptionCellIndex] = " " + combinedFragments + " ";
                }
            }
        }
        
        // テーブル行を再構築
        return "|" + string.Join("|", cells) + "|";
    }

    private static string[] RemoveDuplicateTableSeparators(string[] lines)
    {
        var result = new List<string>();
        bool lastWasTableSeparator = false;
        
        for (int i = 0; i < lines.Length; i++)
        {
            var line = lines[i].Trim();
            
            // テーブル区切り行パターン（| --- | --- | など）
            bool isTableSeparator = System.Text.RegularExpressions.Regex.IsMatch(line, @"^\|\s*---\s*(\|\s*---\s*)*\|?$");
            
            if (isTableSeparator)
            {
                // 前の行もテーブル区切りの場合はスキップ
                if (!lastWasTableSeparator)
                {
                    result.Add(lines[i]);
                    lastWasTableSeparator = true;
                }
                // else: 重複する区切り行をスキップ
            }
            else
            {
                result.Add(lines[i]);
                lastWasTableSeparator = false;
            }
        }
        
        return result.ToArray();
    }
}