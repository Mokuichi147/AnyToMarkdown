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
        
        // 現在の行から後方のテーブル行を収集（ヘッダー要素は除外）
        for (int i = currentIndex; i < allElements.Count; i++)
        {
            var currentElement = allElements[i];
            
            if (currentElement.Type == ElementType.TableRow)
            {
                // 汎用的なヘッダー的内容をテーブル行から除外
                if (IsStandaloneHeaderInTable(currentElement))
                {
                    break; // ヘッダー的要素に遭遇したらテーブル終了
                }
                tableRows.Add(currentElement);
            }
            else if (currentElement.Type == ElementType.Header)
            {
                break; // ヘッダー要素に遭遇したらテーブル終了
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
        
        // 空のテーブル（データがない場合）は空文字を返す
        var hasAnyData = allCells.Any(row => row.Any(cell => !string.IsNullOrWhiteSpace(cell)));
        if (!hasAnyData) return "";

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
            
            // 最小列数を3に設定（元の表構造に合わせて）
            maxColumns = Math.Max(targetColumnCount, 3);
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

        // 純粋な座標ベースの適応的閾値設定
        var sortedGaps = gaps.OrderBy(g => g).ToList();
        var avgGap = gaps.Average();
        var maxGap = gaps.Max();
        
        // フォントサイズに基づく基準値
        var avgFontSize = words.Average(w => w.BoundingBox.Height);
        var fontBasedThreshold = avgFontSize * 0.8;
        
        // 統計的閾値計算（四分位数ベース）
        var q1 = sortedGaps.Count > 3 ? sortedGaps[(int)(sortedGaps.Count * 0.25)] : avgGap;
        var q3 = sortedGaps.Count > 3 ? sortedGaps[(int)(sortedGaps.Count * 0.75)] : avgGap;
        var iqr = q3 - q1;
        
        // より敏感な閾値決定（より多くのセル境界を検出）
        double threshold;
        
        // 大きなギャップの検出（セル境界の候補）
        var largeGaps = gaps.Where(g => g > avgGap * 1.5).ToList();
        
        if (largeGaps.Count > 0)
        {
            // 明確なセル境界が存在する場合、最小の大きなギャップを基準とする
            threshold = largeGaps.Min() * 0.7;
        }
        else if (iqr > fontBasedThreshold * 0.5)
        {
            // 中程度の分散がある場合
            threshold = q1 + (iqr * 0.3);
        }
        else
        {
            // ギャップが均等な場合は中央値を基準とする
            var medianGap = sortedGaps[sortedGaps.Count / 2];
            threshold = medianGap * 0.8;
        }
        
        // より低い最小閾値を設定（より多くの分離を許可）
        var minThreshold = fontBasedThreshold * 0.3;
        var maxThreshold = avgFontSize * 2.5;
        
        threshold = Math.Max(threshold, minThreshold);
        threshold = Math.Min(threshold, maxThreshold);

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

        // フォールバック: セル数が少なすぎる場合はより積極的に分割
        if (cells.Count < 3)
        {
            return SplitTextIntoCells(text);
        }

        return cells;
    }
    

    private static List<string> SplitTextIntoCells(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return new List<string>();
        
        // 汎用的な単語分割ロジック - 位置情報なしでの分割
        var parts = new List<string>();
        
        // より厳密な単語境界検出
        var words = text.Split(new[] { ' ', '\t', '\u00A0' }, StringSplitOptions.RemoveEmptyEntries);
        
        // 各単語の特徴分析
        var wordAnalyses = words.Select((word, index) => new
        {
            Word = word,
            Index = index,
            IsNumeric = word.Any(char.IsDigit),
            HasSpecialChars = word.Any(c => !char.IsLetterOrDigit(c) && !char.IsWhiteSpace(c)),
            Length = word.Length
        }).ToList();
        
        // 目標列数に向けて最適な分割ポイントを決定
        var targetColumns = 4;
        var currentColumn = new List<string>();
        var segmentCount = 0;
        
        foreach (var analysis in wordAnalyses)
        {
            currentColumn.Add(analysis.Word);
            
            // セル境界の決定条件：より柔軟な基準
            bool shouldSplit = false;
            
            if (segmentCount < targetColumns - 1)
            {
                // 数値や特殊文字で区切る
                if (analysis.IsNumeric && currentColumn.Count > 1)
                {
                    shouldSplit = true;
                }
                // 一定長さで区切る
                else if (currentColumn.Count >= 3 && string.Join("", currentColumn).Length > 10)
                {
                    shouldSplit = true;
                }
                // 最後の数語は一つのセルに
                else if (analysis.Index >= words.Length - 2)
                {
                    shouldSplit = false;
                }
            }
            
            if (shouldSplit || segmentCount == targetColumns - 1)
            {
                parts.Add(string.Join(" ", currentColumn).Trim());
                currentColumn.Clear();
                segmentCount++;
                
                if (segmentCount >= targetColumns) break;
            }
        }
        
        // 残りの単語をまとめる
        if (currentColumn.Count > 0)
        {
            parts.Add(string.Join(" ", currentColumn).Trim());
        }
        
        // フォールバック: より積極的な分割で多列対応
        var candidates = text.Split(new[] { ' ', '\t', '　' }, StringSplitOptions.RemoveEmptyEntries);
        
        // より細かい分割を試行（3-4列を目指す）
        if (candidates.Length >= 2)
        {
            var result = new List<string>();
            var currentGroup = new List<string>();
            
            foreach (var candidate in candidates)
            {
                // 汎用的な独立要素判定（数値、年度、記号パターン）
                if (System.Text.RegularExpressions.Regex.IsMatch(candidate, @"^\d+$|^\d{4}年|^\+?\-?\d+\.?\d*%?$|^[A-Z]\d+$"))
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
        // 統合条件を厳格化して誤統合を防ぐ
        if (nextRow.Count == 0 || currentRow.Count == 0) return false;
        
        // セル数が大きく異なる場合は統合しない
        if (Math.Abs(currentRow.Count - nextRow.Count) > 1) return false;
        
        var firstCellNext = nextRow[0]?.Replace("￿", "").Replace("\uFFFD", "").Trim() ?? "";
        var firstCellCurrent = currentRow[0]?.Replace("￿", "").Replace("\uFFFD", "").Trim() ?? "";
        
        // 次の行の最初のセルが空でない場合は新しい行として扱う
        if (!string.IsNullOrWhiteSpace(firstCellNext))
        {
            return false;
        }
        
        // 現在の行の最初のセルが空の場合も統合しない（両方とも継続行の可能性）
        if (string.IsNullOrWhiteSpace(firstCellCurrent))
        {
            return false;
        }
        
        // 次の行に意味のあるコンテンツがあるかチェック
        var hasValidContent = nextRow.Skip(1).Any(c => 
        {
            var cleanCell = c?.Replace("￿", "").Replace("\uFFFD", "").Trim() ?? "";
            return !string.IsNullOrWhiteSpace(cleanCell) && cleanCell.Length > 1;
        });
        
        // 条件1: 最初のセルが空で、2番目以降のセルに内容がある
        if (string.IsNullOrWhiteSpace(firstCellNext) && hasValidContent)
        {
            // 次の行のセル数が少ない場合のみ統合（継続行の特徴）
            var nonEmptyNextCells = nextRow.Count(c => !string.IsNullOrWhiteSpace(c?.Replace("￿", "").Replace("\uFFFD", "").Trim()));
            return nonEmptyNextCells <= 2 && nonEmptyNextCells < currentRow.Count;
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
        
        // テーブル内の太字ヘッダーをMarkdownヘッダーに変換
        lines = ExtractHeadersFromTables(lines);
        
        // より積極的な後処理：残存する太字ヘッダー行を除去
        lines = CleanupRemainingBoldHeaders(lines);
        
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
                    
                    // テキスト行を収集（さらに寛容な条件）
                    if (nextLine.Length > 0 && !nextLine.Contains("|"))
                    {
                        nextLines.Add(nextLine);
                        j++;
                        
                        // 長いテキストの場合は1行で終了
                        if (nextLine.Length > 50) break;
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
        
        // 意味的な断片分析を行って適切なセルに配置
        if (fragments.Count > 0 && cells.Count >= 2)
        {
            var processedFragments = AnalyzeAndProcessFragments(fragments, cells);
            
            for (int i = 0; i < Math.Min(processedFragments.Count, cells.Count); i++)
            {
                if (!string.IsNullOrWhiteSpace(processedFragments[i]))
                {
                    if (!string.IsNullOrWhiteSpace(cells[i].Trim()))
                    {
                        // 既存の内容がある場合は<br>で結合
                        cells[i] = " " + cells[i].Trim() + "<br>" + processedFragments[i] + " ";
                    }
                    else
                    {
                        // 空のセルの場合は直接配置
                        cells[i] = " " + processedFragments[i] + " ";
                    }
                }
            }
        }
        
        // テーブル行を再構築
        return "|" + string.Join("|", cells) + "|";
    }

    private static string[] ExtractHeadersFromTables(string[] lines)
    {
        var result = new List<string>();
        
        for (int i = 0; i < lines.Length; i++)
        {
            var line = lines[i].Trim();
            
            // テーブル行かつ太字ヘッダー要素を含む場合
            if (line.StartsWith("|") && line.Contains("**"))
            {
                var headerExtracted = ExtractAndConvertBoldHeader(line);
                if (headerExtracted.IsHeader)
                {
                    // ヘッダーをMarkdown形式で追加
                    result.Add($"## {headerExtracted.HeaderText}");
                    result.Add("");
                    
                    // 残りのテーブル内容があれば追加
                    if (!string.IsNullOrWhiteSpace(headerExtracted.RemainingContent))
                    {
                        result.Add(headerExtracted.RemainingContent);
                    }
                    continue;
                }
            }
            
            result.Add(lines[i]);
        }
        
        return result.ToArray();
    }
    
    private static (bool IsHeader, string HeaderText, string RemainingContent) ExtractAndConvertBoldHeader(string tableLine)
    {
        // テーブル行から太字ヘッダーを抽出
        var cells = tableLine.Split('|').Select(c => c.Trim()).Where(c => !string.IsNullOrEmpty(c)).ToList();
        
        foreach (var cell in cells)
        {
            var boldMatch = System.Text.RegularExpressions.Regex.Match(cell, @"\*\*([^*]+)\*\*");
            if (boldMatch.Success)
            {
                var boldText = boldMatch.Groups[1].Value.Trim();
                
                // 汎用的ヘッダー判定
                if (IsLikelyStandaloneHeader(boldText, cells))
                {
                    return (true, boldText, "");
                }
            }
        }
        
        return (false, "", tableLine);
    }
    
    private static bool IsLikelyStandaloneHeader(string boldText, List<string> cells)
    {
        // 汎用的な判定条件（文字数制限なし）：
        // 1. セル構造の分析：他のセルが空または非構造的
        // 2. コンテンツ密度の分析：太字部分がコンテンツの主要部分
        // 3. 表形式パターンの不在
        
        var nonBoldCells = cells.Where(c => !c.Contains("**") || string.IsNullOrWhiteSpace(c.Replace("**", "").Trim())).Count();
        var totalCells = cells.Count;
        
        // セル密度分析：空セルが多い場合はヘッダー構造
        var emptyRatio = (double)nonBoldCells / totalCells;
        
        // コンテンツ分布分析：太字コンテンツが行の主要部分を占める
        var totalContent = string.Join("", cells).Replace("**", "");
        var boldContentRatio = (double)boldText.Length / Math.Max(totalContent.Length, 1);
        
        return emptyRatio >= 0.6 && boldContentRatio >= 0.5;
    }

    private static string[] CleanupRemainingBoldHeaders(string[] lines)
    {
        var result = new List<string>();
        
        for (int i = 0; i < lines.Length; i++)
        {
            var line = lines[i];
            var trimmed = line.Trim();
            
            // テーブル行内の太字ヘッダーパターンを検出（データも含まれているかチェック）
            if (trimmed.StartsWith("|") && trimmed.Contains("**") && trimmed.EndsWith("|"))
            {
                var cells = trimmed.Split('|').Where(c => !string.IsNullOrWhiteSpace(c)).ToList();
                var hasHeadersOnly = true;
                var headers = new List<string>();
                var nonHeaderCells = new List<string>();
                
                foreach (var cell in cells)
                {
                    var cellContent = cell.Trim();
                    if (cellContent.StartsWith("**") && cellContent.EndsWith("**") && cellContent.Length > 4)
                    {
                        var headerText = cellContent.Substring(2, cellContent.Length - 4).Trim();
                        if (!string.IsNullOrWhiteSpace(headerText))
                        {
                            headers.Add(headerText);
                        }
                    }
                    else if (!string.IsNullOrWhiteSpace(cellContent))
                    {
                        hasHeadersOnly = false;
                        nonHeaderCells.Add(cellContent);
                    }
                }
                
                // ヘッダーだけの行の場合のみ抽出
                if (hasHeadersOnly && headers.Count > 0 && nonHeaderCells.Count == 0)
                {
                    foreach (var header in headers)
                    {
                        result.Add($"## {header}");
                        result.Add("");
                    }
                    continue; // 元のテーブル行は除去
                }
            }
            
            result.Add(line);
        }
        
        return result.ToArray();
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

    private static List<string> AnalyzeAndProcessFragments(List<string> fragments, List<string> existingCells)
    {
        var processedFragments = new List<string>(new string[existingCells.Count]);
        
        if (fragments.Count == 0) return processedFragments;
        
        // 断片を意味的に分析
        var analysisResult = AnalyzeFragmentMeaning(fragments, existingCells);
        
        // 最も適切なセルに配置
        for (int i = 0; i < analysisResult.Count; i++)
        {
            var fragment = analysisResult[i];
            if (string.IsNullOrWhiteSpace(fragment.Content)) continue;
            
            var targetCellIndex = fragment.TargetCellIndex;
            if (targetCellIndex >= 0 && targetCellIndex < processedFragments.Count)
            {
                if (string.IsNullOrWhiteSpace(processedFragments[targetCellIndex]))
                {
                    processedFragments[targetCellIndex] = fragment.Content;
                }
                else
                {
                    // 複数の断片が同じセルに配置される場合
                    if (fragment.IsLineBreak)
                    {
                        processedFragments[targetCellIndex] += "<br>" + fragment.Content;
                    }
                    else if (fragment.IsWordFragment)
                    {
                        // 単語断片は直接結合（スペースなし）
                        processedFragments[targetCellIndex] += fragment.Content;
                    }
                    else
                    {
                        // 通常の単語はスペースで結合
                        processedFragments[targetCellIndex] += " " + fragment.Content;
                    }
                }
            }
        }
        
        return processedFragments;
    }
    
    private static List<FragmentAnalysis> AnalyzeFragmentMeaning(List<string> fragments, List<string> existingCells)
    {
        var results = new List<FragmentAnalysis>();
        
        for (int i = 0; i < fragments.Count; i++)
        {
            var fragment = fragments[i].Trim();
            if (string.IsNullOrWhiteSpace(fragment)) continue;
            
            var analysis = new FragmentAnalysis
            {
                Content = fragment,
                OriginalIndex = i
            };
            
            // セル配置の決定ロジック
            analysis.TargetCellIndex = DetermineTargetCell(fragment, existingCells, i);
            
            // 改行判定ロジック
            analysis.IsLineBreak = ShouldBeLineBreak(fragment, fragments, i);
            
            // 単語完成性の判定
            analysis.IsWordFragment = IsWordFragment(fragment);
            
            results.Add(analysis);
        }
        
        // 単語断片を結合
        results = ConsolidateWordFragments(results);
        
        return results;
    }
    
    private static int DetermineTargetCell(string fragment, List<string> existingCells, int fragmentIndex)
    {
        // デフォルトは最初のセル
        if (existingCells.Count < 2) return 0;
        
        // 汎用的なパターン分析による配置決定（言語・ドメイン非依存）
        
        // 数値のみのフラグメント
        if (System.Text.RegularExpressions.Regex.IsMatch(fragment, @"^\d{1,4}$"))
        {
            return Math.Min(1, existingCells.Count - 1);
        }
        
        // 短いフラグメント（1-3文字）は文字種に基づく分散
        if (fragment.Length <= 3)
        {
            return fragmentIndex % existingCells.Count;
        }
        
        // 長いフラグメント（4文字以上）の場合
        if (fragment.Length >= 4)
        {
            // 文字の種類分析：数字が多い場合
            var digitRatio = (double)fragment.Count(char.IsDigit) / fragment.Length;
            if (digitRatio > 0.5)
            {
                return Math.Min(1, existingCells.Count - 1);
            }
            
            // アルファベットが多い場合
            var letterRatio = (double)fragment.Count(c => char.IsLetter(c) && c < 128) / fragment.Length;
            if (letterRatio > 0.5)
            {
                return Math.Min(2, existingCells.Count - 1);
            }
        }
        
        // フラグメントのインデックスに基づく分散配置（汎用的）
        return fragmentIndex % existingCells.Count;
    }
    
    private static bool ShouldBeLineBreak(string fragment, List<string> allFragments, int currentIndex)
    {
        // 前の断片との関連性を判定
        if (currentIndex > 0)
        {
            var prevFragment = allFragments[currentIndex - 1].Trim();
            
            // 文章の継続でない場合は改行
            if (prevFragment.EndsWith("。") || prevFragment.EndsWith("！") || prevFragment.EndsWith("？"))
            {
                return true;
            }
            
            // 異なる意味的カテゴリの場合は改行
            if (IsDifferentSemanticCategory(prevFragment, fragment))
            {
                return true;
            }
        }
        
        return false;
    }
    
    private static bool IsWordFragment(string fragment)
    {
        // 非常に短い断片は単語の一部の可能性
        if (fragment.Length <= 1) return true;
        
        // ひらがな1-2文字のみの断片（助詞など）
        if (fragment.Length <= 2 && System.Text.RegularExpressions.Regex.IsMatch(fragment, @"^[ひ-ん]+$"))
        {
            return true;
        }
        
        // カタカナの断片
        if (fragment.Length <= 3 && System.Text.RegularExpressions.Regex.IsMatch(fragment, @"^[ァ-ヶ]+$"))
        {
            return true;
        }
        
        return false;
    }
    
    private static bool IsDifferentSemanticCategory(string fragment1, string fragment2)
    {
        var categories1 = GetSemanticCategories(fragment1);
        var categories2 = GetSemanticCategories(fragment2);
        
        return !categories1.Intersect(categories2).Any();
    }
    
    private static HashSet<string> GetSemanticCategories(string text)
    {
        var categories = new HashSet<string>();
        
        if (System.Text.RegularExpressions.Regex.IsMatch(text, @"\d")) categories.Add("numeric");
        if (System.Text.RegularExpressions.Regex.IsMatch(text, @"[一-龯]")) categories.Add("kanji");
        if (System.Text.RegularExpressions.Regex.IsMatch(text, @"[ひ-ん]")) categories.Add("hiragana");
        if (System.Text.RegularExpressions.Regex.IsMatch(text, @"[ァ-ヶ]")) categories.Add("katakana");
        if (System.Text.RegularExpressions.Regex.IsMatch(text, @"[a-zA-Z]")) categories.Add("alphabet");
        
        return categories;
    }
    
    private static List<FragmentAnalysis> ConsolidateWordFragments(List<FragmentAnalysis> analyses)
    {
        var consolidated = new List<FragmentAnalysis>();
        
        for (int i = 0; i < analyses.Count; i++)
        {
            var current = analyses[i];
            
            if (current.IsWordFragment && i + 1 < analyses.Count)
            {
                var next = analyses[i + 1];
                if (next.TargetCellIndex == current.TargetCellIndex)
                {
                    // 同じセルの断片を結合
                    current.Content += next.Content;
                    current.IsWordFragment = false;
                    i++; // 次の断片をスキップ
                }
            }
            
            consolidated.Add(current);
        }
        
        return consolidated;
    }

    private static bool IsStandaloneHeaderInTable(DocumentElement element)
    {
        var content = element.Content.Trim();
        
        // 太字フォーマットされたテキストの汎用的分析
        if (content.Contains("**"))
        {
            var boldMatch = System.Text.RegularExpressions.Regex.Match(content, @"\*\*([^*]+)\*\*");
            if (boldMatch.Success)
            {
                var boldText = boldMatch.Groups[1].Value.Trim();
                var cells = ParseTableCells(element);
                
                // 汎用的判定：
                // 1. 太字テキストが主要な内容（全体の50%以上）
                // 2. セル数が少ない（独立したヘッダー的構造）
                // 3. 空セルが多い（ヘッダー行の特徴）
                var boldRatio = (double)boldText.Length / content.Replace("**", "").Length;
                var emptyRatio = (double)cells.Count(c => string.IsNullOrWhiteSpace(c)) / Math.Max(cells.Count, 1);
                
                return boldRatio > 0.7 && cells.Count <= 2 && emptyRatio >= 0.5;
            }
        }
        
        return false;
    }

    private class FragmentAnalysis
    {
        public string Content { get; set; } = "";
        public int TargetCellIndex { get; set; }
        public bool IsLineBreak { get; set; }
        public bool IsWordFragment { get; set; }
        public int OriginalIndex { get; set; }
    }
}