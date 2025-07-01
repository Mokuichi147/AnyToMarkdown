using System.Text;

namespace AnyToMarkdown.Pdf;

internal static class BlockProcessor
{
    public static string ConvertListItem(DocumentElement element)
    {
        var content = element.Content.Trim();
        
        // 既にMarkdownリスト形式の場合はそのまま返す
        if (content.StartsWith("- ") || content.StartsWith("* ") || content.StartsWith("+ ") ||
            System.Text.RegularExpressions.Regex.IsMatch(content, @"^\d+\.\s+"))
        {
            return content;
        }
        
        // インデントレベルを計算
        var indentLevel = CalculateListIndentLevel(element);
        var indent = new string(' ', indentLevel * 2);
        
        // リストマーカーを正規化
        var normalizedContent = NormalizeListMarker(content);
        
        return $"{indent}{normalizedContent}";
    }

    private static int CalculateListIndentLevel(DocumentElement element)
    {
        // 左マージンに基づいてインデントレベルを計算
        var basMargin = 30.0; // ベースマージン
        var indentUnit = 20.0; // インデント単位
        
        var normalizedMargin = Math.Max(0, element.LeftMargin - basMargin);
        var level = (int)(normalizedMargin / indentUnit);
        
        return Math.Min(Math.Max(level, 0), 5); // 最大5レベルまで
    }

    private static string NormalizeListMarker(string content)
    {
        // 様々なリストマーカーを標準的なMarkdown形式に変換
        
        // 太字装飾が含まれている場合の処理を最初に行う
        if (content.StartsWith("*") && content.Contains("*"))
        {
            // *•項目* のような形式を処理
            var asteriskPattern = System.Text.RegularExpressions.Regex.Match(content, @"^\*([•・‒–—\-\*\+])(.*?)\*(.*)");
            if (asteriskPattern.Success)
            {
                var restContent = asteriskPattern.Groups[2].Value + asteriskPattern.Groups[3].Value;
                return "- " + restContent.TrimStart();
            }
        }
        
        // 日本語の箇条書き記号
        if (content.StartsWith("・") || content.StartsWith("•"))
        {
            return "- " + content.Substring(1).TrimStart();
        }
        
        // ダッシュ系の記号
        if (content.StartsWith("‒") || content.StartsWith("–") || content.StartsWith("—"))
        {
            return "- " + content.Substring(1).TrimStart();
        }
        
        // 太字のリストマーカー
        if (content.StartsWith("**•**") || content.StartsWith("**・**"))
        {
            return "- " + content.Substring(5).TrimStart();
        }
        
        if (content.StartsWith("**-**") || content.StartsWith("**‒**"))
        {
            return "- " + content.Substring(5).TrimStart();
        }
        
        // 番号付きリスト（既に正しい形式の場合）
        if (System.Text.RegularExpressions.Regex.IsMatch(content, @"^\d+\.\s+"))
        {
            return content;
        }
        
        // アルファベット順リスト
        var alphaMatch = System.Text.RegularExpressions.Regex.Match(content, @"^([a-zA-Z])\)\s*(.*)");
        if (alphaMatch.Success)
        {
            return $"- {alphaMatch.Groups[2].Value}";
        }
        
        // 括弧付きアルファベット
        var parenAlphaMatch = System.Text.RegularExpressions.Regex.Match(content, @"^\(([a-zA-Z])\)\s*(.*)");
        if (parenAlphaMatch.Success)
        {
            return $"- {parenAlphaMatch.Groups[2].Value}";
        }
        
        // デフォルト：既存のコンテンツの前に "- " を追加（フォーマット記号を除去）
        var cleanContent = content;
        
        // 開始と終了の*を除去
        if (cleanContent.StartsWith("*") && cleanContent.EndsWith("*"))
        {
            cleanContent = cleanContent.Substring(1, cleanContent.Length - 2);
        }
        
        return "- " + cleanContent.TrimStart();
    }

    public static string ConvertCodeBlock(DocumentElement element, List<DocumentElement> allElements, int currentIndex)
    {
        var content = element.Content.Trim();
        
        // 単一行のインラインコード
        if (content.StartsWith("`") && content.EndsWith("`") && !content.Contains("\n"))
        {
            return content;
        }
        
        // 複数行のコードブロックかチェック
        var codeLines = new List<string> { content };
        
        // 続くコードブロック要素を収集
        for (int i = currentIndex + 1; i < allElements.Count; i++)
        {
            if (allElements[i].Type == ElementType.CodeBlock)
            {
                codeLines.Add(allElements[i].Content.Trim());
            }
            else if (allElements[i].Type == ElementType.Empty)
            {
                continue; // 空行は無視
            }
            else
            {
                break;
            }
        }
        
        return GenerateMarkdownCodeBlock(codeLines);
    }

    private static string GenerateMarkdownCodeBlock(List<string> codeLines)
    {
        if (codeLines.Count == 0) return "";
        
        var result = new StringBuilder();
        
        // 言語を検出
        var language = DetectCodeLanguage(string.Join("\n", codeLines));
        
        // コードブロックの開始
        result.AppendLine($"```{language}");
        
        // コード内容
        foreach (var line in codeLines)
        {
            var cleanedLine = CleanCodeContent(line);
            result.AppendLine(cleanedLine);
        }
        
        // コードブロックの終了
        result.AppendLine("```");
        
        return result.ToString();
    }

    private static string DetectCodeLanguage(string code)
    {
        if (string.IsNullOrWhiteSpace(code)) return "";
        
        // プログラミング言語のキーワードパターン
        var patterns = new Dictionary<string, string[]>
        {
            ["javascript"] = ["function", "var", "let", "const", "=>", "console.log"],
            ["python"] = ["def ", "import ", "from ", "class ", "if __name__"],
            ["csharp"] = ["using ", "namespace ", "class ", "public ", "private "],
            ["java"] = ["public class", "public static", "import java"],
            ["html"] = ["<html", "<div", "<span", "<!DOCTYPE"],
            ["css"] = ["{", "}", ":", ";", "@media"],
            ["sql"] = ["SELECT", "FROM", "WHERE", "INSERT", "UPDATE"],
            ["xml"] = ["<?xml", "<", "/>", "xmlns"],
            ["json"] = ["{", "}", "[", "]", "\":", ","]
        };
        
        var lowerCode = code.ToLowerInvariant();
        
        foreach (var pattern in patterns)
        {
            var matchCount = pattern.Value.Count(keyword => lowerCode.Contains(keyword.ToLowerInvariant()));
            if (matchCount >= 2) // 2つ以上のキーワードがマッチした場合
            {
                return pattern.Key;
            }
        }
        
        return ""; // 言語を特定できない場合は空文字
    }

    private static string CleanCodeContent(string code)
    {
        if (string.IsNullOrEmpty(code)) return code;
        
        // コード内容のクリーンアップ
        
        // 先頭と末尾の余分な空白を除去
        code = code.Trim();
        
        // タブを4つのスペースに変換
        code = code.Replace("\t", "    ");
        
        // 不要なHTMLタグの除去
        code = System.Text.RegularExpressions.Regex.Replace(code, @"<[^>]+>", "");
        
        // 特殊文字のデコード
        code = code.Replace("&lt;", "<")
                   .Replace("&gt;", ">")
                   .Replace("&amp;", "&")
                   .Replace("&quot;", "\"");
        
        return code;
    }

    public static string ConvertQuoteBlock(DocumentElement element, List<DocumentElement> allElements, int currentIndex)
    {
        var content = element.Content.Trim();
        
        // 単一行の引用
        if (content.StartsWith("> "))
        {
            return content;
        }
        
        // 複数行の引用ブロックかチェック
        var quoteLines = new List<string> { content };
        
        // 続く引用ブロック要素を収集
        for (int i = currentIndex + 1; i < allElements.Count; i++)
        {
            if (allElements[i].Type == ElementType.QuoteBlock)
            {
                quoteLines.Add(allElements[i].Content.Trim());
            }
            else if (allElements[i].Type == ElementType.Empty)
            {
                continue; // 空行は無視
            }
            else
            {
                break;
            }
        }
        
        return GenerateMarkdownQuoteBlock(quoteLines);
    }

    private static string GenerateMarkdownQuoteBlock(List<string> quoteLines)
    {
        if (quoteLines.Count == 0) return "";
        
        var result = new StringBuilder();
        
        foreach (var line in quoteLines)
        {
            var cleanedLine = CleanQuoteContent(line);
            
            // 既に > で始まっている場合はそのまま
            if (cleanedLine.StartsWith(">"))
            {
                result.AppendLine(cleanedLine);
            }
            else
            {
                result.AppendLine($"> {cleanedLine}");
            }
        }
        
        return result.ToString();
    }

    private static string CleanQuoteContent(string quote)
    {
        if (string.IsNullOrEmpty(quote)) return quote;
        
        // 引用符の除去
        if (quote.StartsWith("\"") && quote.EndsWith("\""))
        {
            quote = quote.Substring(1, quote.Length - 2);
        }
        
        if (quote.StartsWith("'") && quote.EndsWith("'"))
        {
            quote = quote.Substring(1, quote.Length - 2);
        }
        
        // 日本語の引用符の除去
        if (quote.StartsWith("「") && quote.EndsWith("」"))
        {
            quote = quote.Substring(1, quote.Length - 2);
        }
        
        if (quote.StartsWith("『") && quote.EndsWith("』"))
        {
            quote = quote.Substring(1, quote.Length - 2);
        }
        
        return quote.Trim();
    }

    public static string ConvertHorizontalLine(DocumentElement element)
    {
        // 水平線を標準的なMarkdown形式に変換
        return "---";
    }
}