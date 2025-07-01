using System.Text;
using UglyToad.PdfPig.Content;

namespace AnyToMarkdown.Pdf;

internal static class MarkdownGenerator
{
    public static string GenerateMarkdown(DocumentStructure structure)
    {
        var sb = new StringBuilder();
        var elements = structure.Elements.Where(e => e.Type != ElementType.Empty).ToList();
        
        // 段落の統合処理（改良版）
        var consolidatedElements = ConsolidateParagraphsImproved(elements);
        
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

        return TextPostProcessor.PostProcessMarkdown(sb.ToString());
    }

    private static List<DocumentElement> ConsolidateParagraphsImproved(List<DocumentElement> elements)
    {
        var consolidated = new List<DocumentElement>();
        
        for (int i = 0; i < elements.Count; i++)
        {
            var current = elements[i];
            
            if (current.Type == ElementType.Paragraph)
            {
                var paragraphBuilder = new StringBuilder(current.Content);
                var consolidatedWords = new List<Word>(current.Words);
                
                // より精密な統合条件
                int j = i + 1;
                while (j < elements.Count && elements[j].Type == ElementType.Paragraph)
                {
                    var nextParagraph = elements[j];
                    
                    // 統合適用性を慎重に判定
                    if (!ShouldConsolidateParagraphsImproved(current, nextParagraph))
                    {
                        break;
                    }
                    
                    // スペース追加の改良
                    var separator = DetermineParagraphSeparator(current.Content, nextParagraph.Content);
                    paragraphBuilder.Append(separator).Append(nextParagraph.Content);
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
    
    private static string DetermineParagraphSeparator(string currentText, string nextText)
    {
        return TextFormatter.DetermineParagraphSeparator(currentText, nextText);
    }
    
    private static bool ShouldConsolidateParagraphsImproved(DocumentElement current, DocumentElement next)
    {
        // より保守的な統合アプローチ
        var currentText = current.Content.Trim();
        var nextText = next.Content.Trim();
        
        // 文章の区切りを示す文字で終わっている場合は統合しない（より厳格に）
        if (currentText.EndsWith(".") || currentText.EndsWith("。") || 
            currentText.EndsWith("!") || currentText.EndsWith("？") ||
            currentText.EndsWith("?") || currentText.EndsWith("！") ||
            currentText.EndsWith("です。") || currentText.EndsWith("ます。"))
        {
            return false;
        }
        
        // 異なる文章パターンは統合しない
        if (currentText.StartsWith("これは") && nextText.Contains("**"))
        {
            return false;
        }
        
        // 次の文が大文字で始まっている場合は新しい文の可能性
        if (nextText.Length > 0 && char.IsUpper(nextText[0]))
        {
            return false;
        }
        
        // フォントサイズ差が大きい場合は統合しない
        if (Math.Abs(current.FontSize - next.FontSize) > 1.0)
        {
            return false;
        }
        
        // Markdown記法の特殊文字を含む場合は統合しない
        if (ContainsMarkdownSyntax(currentText) || ContainsMarkdownSyntax(nextText))
        {
            return false;
        }
        
        // 段落の統合を慎重に有効化（改善された条件）
        
        // フォント分析の差が大きい場合は統合しない
        if (Math.Abs(current.FontSize - next.FontSize) > 2.0)
        {
            return false;
        }
        
        // 短い独立した文は統合しない
        if (currentText.Length <= 25 || nextText.Length <= 25)
        {
            return false;
        }
        
        return true;
        
        // 垂直距離による判定
        if (current.Words.Count > 0 && next.Words.Count > 0)
        {
            var currentBottom = current.Words.Min(w => w.BoundingBox.Bottom);
            var nextTop = next.Words.Max(w => w.BoundingBox.Top);
            var verticalGap = Math.Abs(currentBottom - nextTop);
            
            // 適度な垂直ギャップ制限
            if (verticalGap > 15.0)
            {
                return false;
            }
        }
        
        return true;
    }
    
    private static bool ContainsMarkdownSyntax(string text)
    {
        // Markdownの特殊記法を検出
        return text.Contains("**") || text.Contains("__") || 
               text.Contains("```") || text.Contains("`") ||
               text.Contains("[") && text.Contains("](") ||
               text.StartsWith("#") || text.StartsWith(">") ||
               text.StartsWith("-") || text.StartsWith("*") || text.StartsWith("+") ||
               text.Contains("---") || text.Contains("***") || text.Contains("___");
    }

    private static string ConvertElementToMarkdown(DocumentElement element, List<DocumentElement> allElements, int currentIndex, FontAnalysis fontAnalysis)
    {
        return element.Type switch
        {
            ElementType.Header => HeaderProcessor.ConvertHeader(element, fontAnalysis),
            ElementType.ListItem => BlockProcessor.ConvertListItem(element),
            ElementType.TableRow => TableProcessor.ConvertTableRow(element, allElements, currentIndex),
            ElementType.CodeBlock => BlockProcessor.ConvertCodeBlock(element, allElements, currentIndex),
            ElementType.QuoteBlock => BlockProcessor.ConvertQuoteBlock(element, allElements, currentIndex),
            ElementType.HorizontalLine => BlockProcessor.ConvertHorizontalLine(element),
            ElementType.Paragraph => ConvertParagraph(element),
            _ => element.Content
        };
    }

    private static string ConvertParagraph(DocumentElement element)
    {
        return element.Content?.Trim() ?? "";
    }
}