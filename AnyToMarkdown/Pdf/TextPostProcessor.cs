using System.Linq;
using System.Text;

namespace AnyToMarkdown.Pdf;

internal static class TextPostProcessor
{
    public static string PostProcessMarkdown(string markdown)
    {
        if (string.IsNullOrWhiteSpace(markdown))
            return "";
        
        // 禁止文字（null文字など）を除去
        markdown = RemoveForbiddenCharacters(markdown);
        
        // HTMLタグの除去（最優先で処理）
        markdown = RemoveHtmlTags(markdown);
        
        // エスケープ文字の復元
        markdown = RestoreEscapeCharacters(markdown);
        
        // 特殊文字の正規化
        markdown = NormalizeSpecialCharacters(markdown);
            
        var lines = markdown.Split('\n');
        
        // テーブル内の太字ヘッダーをMarkdownヘッダーに変換
        lines = ExtractHeadersFromTables(lines);
        
        // より積極的な後処理：残存する太字ヘッダー行を除去
        lines = CleanupRemainingBoldHeaders(lines);
        
        // テーブルの分離された行を統合する前処理
        lines = TableProcessor.MergeDisconnectedTableCells(lines);
        
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
            
            // 空行の処理
            if (string.IsNullOrWhiteSpace(trimmed))
            {
                if (!previousWasEmpty)
                {
                    processedLines.Add("");
                    previousWasEmpty = true;
                }
                continue;
            }
            
            processedLines.Add(line);
            previousWasEmpty = false;
        }
        
        // 末尾の空行を除去
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
        
        // 置換文字（U+FFFD）を除去
        text = text.Replace("￿", "").Replace("\uFFFD", "");
        
        // より包括的な制御文字除去
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
    
    private static string RemoveHtmlTags(string text)
    {
        if (string.IsNullOrEmpty(text)) return text;
        
        // 各種HTMLタグの除去（大文字小文字を問わず）
        text = System.Text.RegularExpressions.Regex.Replace(text, @"<br\s*/?>\s*", " ", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        text = System.Text.RegularExpressions.Regex.Replace(text, @"<BR\s*/?>\s*", " ", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        text = System.Text.RegularExpressions.Regex.Replace(text, @"<p\s*/?>\s*", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        text = System.Text.RegularExpressions.Regex.Replace(text, @"</p>\s*", " ", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        text = System.Text.RegularExpressions.Regex.Replace(text, @"<div\s*/?>\s*", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        text = System.Text.RegularExpressions.Regex.Replace(text, @"</div>\s*", " ", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        
        // 一般的なHTMLタグの除去
        text = System.Text.RegularExpressions.Regex.Replace(text, @"<[^>]+>", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        
        // HTMLエンティティのデコード
        text = System.Text.RegularExpressions.Regex.Replace(text, @"&nbsp;", " ");
        text = System.Text.RegularExpressions.Regex.Replace(text, @"&amp;", "&");
        text = System.Text.RegularExpressions.Regex.Replace(text, @"&lt;", "<");
        text = System.Text.RegularExpressions.Regex.Replace(text, @"&gt;", ">");
        text = System.Text.RegularExpressions.Regex.Replace(text, @"&quot;", "\"");
        
        // 余分な空白を統合（改行は保持）
        text = System.Text.RegularExpressions.Regex.Replace(text, @"[ \t]+", " ");
        
        return text.Trim();
    }

    private static string RestoreEscapeCharacters(string text)
    {
        if (string.IsNullOrEmpty(text)) return text;
        
        // バックスラッシュエスケープの復元
        text = text.Replace("\\\\", "\uE000"); // 一時的な置換
        text = text.Replace("\\*", "*");
        text = text.Replace("\\_", "_");
        text = text.Replace("\\#", "#");
        text = text.Replace("\\[", "[");
        text = text.Replace("\\]", "]");
        text = text.Replace("\\(", "(");
        text = text.Replace("\\)", ")");
        text = text.Replace("\\|", "|");
        text = text.Replace("\uE000", "\\"); // 元の\\を復元
        
        return text;
    }

    private static string NormalizeSpecialCharacters(string text)
    {
        if (string.IsNullOrEmpty(text)) return text;
        
        // Unicode正規化
        text = text.Normalize(NormalizationForm.FormC);
        
        // 特殊な引用符を標準的なものに置換
        text = text.Replace(""", "\"").Replace(""", "\"");
        text = text.Replace("'", "'").Replace("'", "'");
        
        // 特殊なダッシュを標準的なものに置換
        text = text.Replace("—", "--").Replace("–", "-");
        
        // 特殊な空白文字を通常のスペースに置換
        text = System.Text.RegularExpressions.Regex.Replace(text, @"[\u00A0\u2000-\u200B\u2028\u2029]", " ");
        
        return text;
    }

    private static string[] ExtractHeadersFromTables(string[] lines)
    {
        var result = new List<string>();
        
        for (int i = 0; i < lines.Length; i++)
        {
            var line = lines[i].Trim();
            
            // テーブル行で太字のヘッダーっぽいものを検出
            if (line.StartsWith("|") && line.EndsWith("|"))
            {
                var cells = line.Split('|')
                    .Where(cell => !string.IsNullOrWhiteSpace(cell))
                    .Select(cell => cell.Trim())
                    .ToList();
                
                // 全てのセルが太字で短い場合はヘッダーとして抽出
                if (cells.Count > 0 && cells.All(cell => 
                    (cell.StartsWith("**") && cell.EndsWith("**") && cell.Length <= 30) ||
                    (cell.StartsWith("***") && cell.EndsWith("***") && cell.Length <= 30)))
                {
                    // 最初のセルをヘッダーとして抽出
                    var headerText = cells[0].Replace("**", "").Replace("***", "").Trim();
                    if (!string.IsNullOrWhiteSpace(headerText))
                    {
                        result.Add($"# {headerText}");
                        
                        // 残りのセルがある場合は通常のテーブル行として処理
                        if (cells.Count > 1)
                        {
                            var remainingCells = cells.Skip(1).ToList();
                            var newTableRow = "| " + string.Join(" | ", remainingCells) + " |";
                            result.Add(newTableRow);
                        }
                        continue;
                    }
                }
            }
            
            result.Add(line);
        }
        
        return result.ToArray();
    }

    private static string[] CleanupRemainingBoldHeaders(string[] lines)
    {
        var result = new List<string>();
        
        foreach (var line in lines)
        {
            var trimmed = line.Trim();
            
            // 単独の太字テキスト行がヘッダーの可能性
            if ((trimmed.StartsWith("**") && trimmed.EndsWith("**") && !trimmed.Contains("|")) ||
                (trimmed.StartsWith("***") && trimmed.EndsWith("***") && !trimmed.Contains("|")))
            {
                var headerText = trimmed.Replace("**", "").Replace("***", "").Trim();
                if (!string.IsNullOrWhiteSpace(headerText) && headerText.Length <= 50)
                {
                    result.Add($"# {headerText}");
                    continue;
                }
            }
            
            result.Add(line);
        }
        
        return result.ToArray();
    }

    private static string[] RemoveDuplicateTableSeparators(string[] lines)
    {
        var result = new List<string>();
        string? previousLine = null;
        
        foreach (var line in lines)
        {
            var trimmed = line.Trim();
            
            // テーブル区切り行の検出
            if (System.Text.RegularExpressions.Regex.IsMatch(trimmed, @"^[\|\-\s]+$"))
            {
                // 前の行も区切り行の場合はスキップ
                if (previousLine != null && 
                    System.Text.RegularExpressions.Regex.IsMatch(previousLine.Trim(), @"^[\|\-\s]+$"))
                {
                    continue;
                }
            }
            
            result.Add(line);
            previousLine = line;
        }
        
        return result.ToArray();
    }
}