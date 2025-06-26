using System.Text;
using UglyToad.PdfPig.Content;

namespace AnyToMarkdown.Pdf;

internal class PdfStructureAnalyzer
{
    public static DocumentStructure AnalyzePageStructure(Page page, double horizontalTolerance = 5.0, double verticalTolerance = 5.0)
    {
        var words = page.GetWords()
            .OrderByDescending(x => x.BoundingBox.Bottom)
            .ThenBy(x => x.BoundingBox.Left)
            .ToList();

        var lines = PdfWordProcessor.GroupWordsIntoLines(words, verticalTolerance);
        var documentStructure = new DocumentStructure();

        // より詳細なフォント分析
        var fontAnalysis = AnalyzeFontDistribution(words);
        
        // 図形情報を抽出してテーブルの境界を検出
        var graphicsInfo = ExtractGraphicsInfo(page);
        
        var elements = new List<DocumentElement>();
        for (int i = 0; i < lines.Count; i++)
        {
            var element = AnalyzeLine(lines[i], fontAnalysis, horizontalTolerance);
            elements.Add(element);
        }

        // 後処理：コンテキスト情報を活用した要素分類の改善（テーブル検出前）
        elements = PostProcessElementClassification(elements, fontAnalysis);
        
        // 後処理：図形情報と連続する行の構造パターンを分析してテーブルを検出
        elements = PostProcessTableDetection(elements, graphicsInfo);
        
        documentStructure.Elements.AddRange(elements);
        documentStructure.FontAnalysis = fontAnalysis;
        return documentStructure;
    }

    private static FontAnalysis AnalyzeFontDistribution(List<Word> words)
    {
        if (words.Count == 0)
        {
            return new FontAnalysis
            {
                BaseFontSize = 12.0,
                LargeFontThreshold = 15.6,
                SmallFontThreshold = 9.6,
                DominantFont = "unknown",
                AllFontSizes = [12.0]
            };
        }

        var fontSizes = words.GroupBy(w => Math.Round(w.BoundingBox.Height, 1))
                            .OrderByDescending(g => g.Count())
                            .ToList();

        var fontNames = words.GroupBy(w => w.FontName ?? "unknown")
                            .OrderByDescending(g => g.Count())
                            .ToList();

        // 最も頻度の高いフォントサイズをベースとして使用
        double baseFontSize = fontSizes.First().Key;
        string dominantFont = fontNames.First().Key;

        // 段落レベルのテキストを探してより正確なベースサイズを決定
        var paragraphWords = words.Where(w => 
        {
            var text = w.Text?.Trim();
            return !string.IsNullOrEmpty(text) && 
                   text.Length > 3 && 
                   !text.All(char.IsDigit) &&
                   !(text.StartsWith("#") || text.StartsWith("-") || text.StartsWith("*"));
        }).ToList();

        if (paragraphWords.Count > 0)
        {
            var paragraphFontSizes = paragraphWords.GroupBy(w => Math.Round(w.BoundingBox.Height, 1))
                                                   .OrderByDescending(g => g.Count())
                                                   .ToList();
            baseFontSize = paragraphFontSizes.First().Key;
        }

        // より保守的な閾値設定
        double largeFontThreshold = baseFontSize * 1.15;
        double smallFontThreshold = baseFontSize * 0.85;

        return new FontAnalysis
        {
            BaseFontSize = baseFontSize,
            LargeFontThreshold = largeFontThreshold,
            SmallFontThreshold = smallFontThreshold,
            DominantFont = dominantFont,
            AllFontSizes = [.. fontSizes.Select(g => g.Key)]
        };
    }

    private static DocumentElement AnalyzeLine(List<Word> line, FontAnalysis fontAnalysis, double horizontalTolerance)
    {
        if (line.Count == 0)
            return new DocumentElement { Type = ElementType.Empty, Content = "" };

        var mergedWords = PdfWordProcessor.MergeWordsInLine(line, horizontalTolerance);
        
        // フォント検出機能付きのテキスト生成（完全な実装）
        var formattedText = BuildFormattedText(mergedWords);

        if (string.IsNullOrWhiteSpace(formattedText))
            return new DocumentElement { Type = ElementType.Empty, Content = "" };

        // フォントサイズ分析
        var avgFontSize = line.Average(w => w.BoundingBox.Height);
        var maxFontSize = line.Max(w => w.BoundingBox.Height);
        
        // 位置分析
        var leftMargin = line.Min(w => w.BoundingBox.Left);
        var isIndented = leftMargin > 50; // 基準マージンから50pt以上右

        // 要素タイプの判定
        var elementType = DetermineElementType(formattedText, avgFontSize, maxFontSize, fontAnalysis, isIndented, line);

        return new DocumentElement
        {
            Type = elementType,
            Content = formattedText,
            FontSize = avgFontSize,
            LeftMargin = leftMargin,
            IsIndented = isIndented,
            Words = line
        };
    }

    private static ElementType DetermineElementType(string text, double avgFontSize, double maxFontSize, 
        FontAnalysis fontAnalysis, bool isIndented, List<Word> words)
    {
        // フォーマット済みテキストから元のテキストを抽出
        var cleanText = ExtractCleanTextForAnalysis(text);
        
        // 明確なMarkdownパターン
        if (cleanText.StartsWith("#")) return ElementType.Header;
        if (cleanText.StartsWith(">")) return ElementType.QuoteBlock;
        if (cleanText.StartsWith("```")) return ElementType.CodeBlock;
        
        // ヘッダー判定を最優先（フォントサイズと内容の両方を考慮）
        bool isLargeFont = maxFontSize > fontAnalysis.LargeFontThreshold;
        bool hasHeaderContent = IsHeaderLike(cleanText);
        bool isShortText = cleanText.Length <= 20; // 短いテキストはヘッダーの可能性が高い
        
        // 強力なヘッダー判定条件
        if ((isLargeFont && hasHeaderContent) || 
            (isLargeFont && isShortText && !cleanText.Contains("|")) ||
            (hasHeaderContent && text.Contains("**") && cleanText.Length <= 15))
        {
            return ElementType.Header;
        }
        
        // 太字フォーマットを含む数字パターンを最優先でヘッダーとして扱う
        if (HasBoldNumberPattern(text))
        {
            return ElementType.Header;
        }
        
        // 太字テキストで短い行の場合はヘッダーとして扱う
        if (IsPotentialBoldHeader(text))
        {
            return ElementType.Header;
        }
        
        if (cleanText.StartsWith("-") || cleanText.StartsWith("*") || cleanText.StartsWith("+")) return ElementType.ListItem;
        if (cleanText.StartsWith("・") || cleanText.StartsWith("•")) return ElementType.ListItem;

        // コードブロックの判定（モノスペースフォントやインデント）
        if (IsCodeBlockLike(cleanText, words, fontAnalysis)) return ElementType.CodeBlock;
        
        // 引用ブロックの判定
        if (IsQuoteBlockLike(cleanText)) return ElementType.QuoteBlock;

        // リストアイテムの判定を最優先（数字付きリストを含む）
        if (IsListItemLike(cleanText)) return ElementType.ListItem;

        // テーブル行の判定（座標とギャップベース）- ヘッダー判定より先に実行
        if (IsTableRowLike(cleanText, words)) return ElementType.TableRow;

        // フォントサイズベースのヘッダー判定（改良版）
        var fontSizeRatio = maxFontSize / fontAnalysis.BaseFontSize;
        
        // 明らかに大きなフォントサイズの場合
        if (fontSizeRatio > 1.15 && !cleanText.EndsWith("。") && !cleanText.EndsWith("."))
        {
            return ElementType.Header;
        }
        
        // 中程度のフォントサイズでヘッダーパターンを持つ場合
        if (fontSizeRatio >= 1.05 && IsHeaderLike(cleanText))
        {
            return ElementType.Header;
        }
        
        // フォントサイズが小さくても明確なヘッダーパターンがある場合
        if (fontSizeRatio >= 1.0 && IsStrongHeaderPattern(cleanText))
        {
            return ElementType.Header;
        }
        
        // 太字フォーマットを含む数字パターンを強制的にヘッダーとして扱う
        if (HasBoldNumberPattern(text))
        {
            return ElementType.Header;
        }
        
        // 太字テキストで短い行の場合はヘッダーとして扱う
        if (IsPotentialBoldHeader(text))
        {
            return ElementType.Header;
        }

        // 位置ベースの判定
        if (isIndented && IsListItemLike(cleanText))
        {
            return ElementType.ListItem;
        }

        return ElementType.Paragraph;
    }
    
    private static string ExtractCleanTextForAnalysis(string text)
    {
        // Markdownフォーマットを除去してテキスト分析を行う
        var cleanText = text;
        
        // 太字フォーマットを除去（複数回実行して入れ子を処理）
        while (cleanText.Contains("**") || cleanText.Contains("*"))
        {
            var before = cleanText;
            cleanText = System.Text.RegularExpressions.Regex.Replace(cleanText, @"\*{1,3}([^*]*)\*{1,3}", "$1");
            if (before == cleanText) break; // 変化がなければ終了
        }
        
        // 斜体フォーマットを除去  
        cleanText = System.Text.RegularExpressions.Regex.Replace(cleanText, @"_([^_]+)_", "$1");
        
        // 余分なスペースを統合
        cleanText = System.Text.RegularExpressions.Regex.Replace(cleanText, @"\s+", " ");
        
        return cleanText.Trim();
    }

    private static bool IsHeaderLike(string text)
    {
        // 明確なMarkdownヘッダーパターン
        if (text.StartsWith("#")) return true;

        // 文末の表現はヘッダーではない
        if (text.EndsWith("。") || text.EndsWith(".") || text.Contains("、")) return false;
        
        // 短い単一の単語/フレーズはヘッダーの可能性が高い（12文字以下）
        var cleanText = text.Trim();
        if (cleanText.Length <= 12 && !cleanText.Contains(" ") && 
            !cleanText.All(char.IsDigit) && !cleanText.Contains("|"))
        {
            return true;
        }
        
        // "基本的な"、"空欄を含む"などの修飾詞＋名詞パターン
        if (System.Text.RegularExpressions.Regex.IsMatch(cleanText, @"^.{1,8}な.{1,10}$") ||
            System.Text.RegularExpressions.Regex.IsMatch(cleanText, @"^.{1,8}を含む.{1,10}$"))
        {
            return true;
        }
        
        // 階層的な数字パターンはヘッダー (1.1, 1.1.1)
        var hierarchicalPattern = @"^\d+\.\d+";
        if (System.Text.RegularExpressions.Regex.IsMatch(text, hierarchicalPattern)) return true;
        
        // 章番号とタイトルのパターン (1. 概要, 2.1 システム要件など)
        if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^\d+[\.\s]+\S+")) return true;

        return false;
    }

    private static bool IsStrongHeaderPattern(string text)
    {
        // 明確なMarkdownヘッダーパターン
        if (text.StartsWith("#")) return true;
        
        // 数字のみの章番号パターン
        if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^\d{1,3}$")) return true;
        
        // 階層的な数字パターン (1.1, 1.1.1, 1.1.1.1)
        if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^\d+(\.\d+){1,3}$")) return true;
        
        // 章番号とタイトルのパターン (1. 概要, 2.1 システム要件など)
        if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^\d+(\.\d+)*[\.\s]+\S+")) return true;
        
        // 大文字のタイトル（英語）
        if (text.All(c => char.IsUpper(c) || char.IsWhiteSpace(c) || char.IsDigit(c)) && text.Any(char.IsLetter)) return true;
        
        return false;
    }

    private static bool IsListItemLike(string text)
    {
        text = text.Trim();
        if (string.IsNullOrEmpty(text)) return false;
        
        // 明確なリストマーカー
        if (text.StartsWith("・") || text.StartsWith("•") || text.StartsWith("◦")) return true;
        if (text.StartsWith("-") || text.StartsWith("*") || text.StartsWith("+")) return true;
        
        // 数字付きリスト
        if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^\d+[\.\)]")) return true;
        
        // 括弧付き数字
        if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^\(\d+\)")) return true;
        
        // アルファベット付きリスト
        if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^[a-zA-Z][\.\)]")) return true;
        
        // Unicode数字記号（丸数字など）
        if (!string.IsNullOrEmpty(text))
        {
            var firstChar = text[0];
            if (firstChar >= '\u2460' && firstChar <= '\u2473') return true; // 丸数字
            if (firstChar >= '\u2474' && firstChar <= '\u2487') return true; // 括弧付き数字
        }

        return false;
    }

    private static bool IsTableRowLike(string text, List<Word> words)
    {
        // 禁止文字を除去してからチェック
        text = text.Replace("￿", "").Replace("\uFFFD", "").Trim();
        if (string.IsNullOrEmpty(text)) return false;
        
        var parts = text.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length < 2) return false;

        // パイプ文字がある（既にMarkdownテーブル）
        if (text.Contains("|")) return true;

        // 数値密度による表判定（言語・ドメイン非依存）
        var digitRatio = (double)text.Count(char.IsDigit) / text.Length;
        if (digitRatio > 0.2 && parts.Length >= 2) return true;

        // 数値の比率が高い場合
        int numericParts = parts.Count(p => double.TryParse(p.Replace(",", ""), out _) || 
            p.Contains("%") || p.Contains(",") || System.Text.RegularExpressions.Regex.IsMatch(p, @"^\+?-?\d+\.?\d*$"));
        if (parts.Length >= 3 && (double)numericParts / parts.Length > 0.4) return true;

        // 単語間の距離による判定（最も重要な指標）
        if (words != null && words.Count >= 3)
        {
            var gaps = new List<double>();
            for (int i = 1; i < words.Count; i++)
            {
                gaps.Add(words[i].BoundingBox.Left - words[i-1].BoundingBox.Right);
            }
            
            if (gaps.Count >= 2)
            {
                var avgGap = gaps.Average();
                var largeGaps = gaps.Count(g => g > Math.Max(avgGap * 1.5, 20));
                
                // 大きなギャップが複数ある場合
                if (largeGaps >= 2) return true;
            }
        }

        return false;
    }
    
    private static bool IsCodeBlockLike(string text, List<Word> words, FontAnalysis fontAnalysis)
    {
        if (string.IsNullOrWhiteSpace(text)) return false;
        
        // 明確なコードブロック開始・終了マーカー
        if (text.StartsWith("```") || text.EndsWith("```")) return true;
        
        // プログラミング言語のキーワードパターン
        var codeKeywords = new[] { "def ", "class ", "function ", "import ", "from ", "if __name__", "try:", "except:", "finally:", "with ", "async def", "await ", "return ", "yield ", "break", "continue", "pass", "raise", "assert", "del", "global", "nonlocal", "lambda", "for ", "while ", "if ", "elif ", "else:", "and ", "or ", "not ", "in ", "is ", "True", "False", "None" };
        if (codeKeywords.Any(keyword => text.Contains(keyword))) return true;
        
        // JSON/設定ファイルパターン
        if ((text.Contains("{") && text.Contains("}")) || (text.Contains("[") && text.Contains("]"))) return true;
        if (text.Contains("\"key\":") || text.Contains("'key':")) return true;
        
        // コマンドラインパターン
        if (text.StartsWith("$") || text.StartsWith("#") || text.Contains("--")) return true;
        
        // インデントが深い場合（4スペース以上）
        if (text.StartsWith("    ") || text.StartsWith("\t")) return true;
        
        // モノスペースフォントの検出
        if (words != null && words.Count > 0)
        {
            var fontNames = words.Where(w => !string.IsNullOrEmpty(w.FontName))
                                .GroupBy(w => w.FontName)
                                .OrderByDescending(g => g.Count())
                                .ToList();
                                
            if (fontNames.Count > 0)
            {
                var dominantFont = fontNames.First().Key.ToLower();
                // 一般的なモノスペースフォント名
                var monospaceFonts = new[] { "courier", "consolas", "monaco", "menlo", "source code", "dejavu sans mono", "liberation mono", "ubuntu mono" };
                if (monospaceFonts.Any(font => dominantFont.Contains(font))) return true;
            }
        }
        
        return false;
    }
    
    private static bool IsQuoteBlockLike(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return false;
        
        // 明確な引用マーカー
        if (text.StartsWith(">")) return true;
        
        // 日本語の引用表現
        if (text.StartsWith("「") && text.EndsWith("」")) return true;
        if (text.StartsWith("『") && text.EndsWith("』")) return true;
        
        // 英語の引用符
        if (text.StartsWith("\"") && text.EndsWith("\"")) return true;
        if (text.StartsWith("'") && text.EndsWith("'")) return true;
        
        // 引用を示すフレーズ
        var quoteIndicators = new[] { "注意", "重要", "警告", "Note:", "Important:", "Warning:", "Tip:", "記:", "備考" };
        if (quoteIndicators.Any(indicator => text.StartsWith(indicator))) return true;
        
        return false;
    }
    
    private static bool HasBoldNumberPattern(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return false;
        
        // **1.** **概要** や **1.1目的** のようなパターンを検出
        var boldNumberPatterns = new[]
        {
            @"\*\*\d+\.\*\*\s*\*\*[^\*]+\*\*", // **1.** **概要**
            @"\*\*\d+\.\d+[^\*]*\*\*",          // **1.1目的**
            @"\*\*\d+\.\*\*\s*[^\*]+",          // **1.** 概要 (一部太字)
            @"\*\*\d+\.\d+\*\*\s*[^\*]+",       // **1.1** システム要件
            @"\*\*\d+\.\*\*",                   // **1.**
            @"\*\*\d+\.\d+\*\*"                 // **1.1**
        };
        
        return boldNumberPatterns.Any(pattern => 
            System.Text.RegularExpressions.Regex.IsMatch(text, pattern));
    }
    
    private static bool IsPotentialBoldHeader(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return false;
        
        // 太字を含むテキストをヘッダー候補とする（文字数に依存しない）
        if (text.Contains("**"))
        {
            // リストアイテムのマーカーは除外
            if (text.StartsWith("- ") || text.StartsWith("* ") || text.StartsWith("+ ")) return false;
            
            // 数字リストの項目は除外
            if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^\d+\.\s*\*\*")) return false;
            
            // 完全な文（句点で終わる）は除外
            var cleanText = text.Replace("**", "").Trim();
            if (cleanText.EndsWith("。") || cleanText.EndsWith(".")) return false;
            
            // 日本語の項目マーカーは除外
            if (cleanText.StartsWith("•") || cleanText.StartsWith("‒")) return false;
            
            // 数字パターンを含む太字テキストで、かつセクション番号形式
            if (System.Text.RegularExpressions.Regex.IsMatch(text, @"\*\*\d+(\.\d+)*[\.\s]"))
            {
                return true;
            }
            
            // 明確なセクションタイトルパターン（数字以外の見出し）
            if (System.Text.RegularExpressions.Regex.IsMatch(text, @"\*\*[^\d\*\s][^\*]*\*\*$"))
            {
                return true;
            }
        }
        
        return false;
    }
    
    private static string BuildFormattedText(List<List<Word>> mergedWords)
    {
        if (mergedWords.Count == 0) return "";
        
        var result = new System.Text.StringBuilder();
        
        foreach (var wordGroup in mergedWords)
        {
            if (wordGroup.Count == 0) continue;
            
            // 単語レベルでの書式設定を処理
            var groupResult = BuildFormattedWordGroup(wordGroup);
            if (string.IsNullOrWhiteSpace(groupResult)) continue;
            
            // スペースの追加 - 最初の単語でない場合のみ
            if (result.Length > 0) result.Append(" ");
            result.Append(groupResult);
        }
        
        var finalResult = result.ToString();
        
        // 禁止文字（null文字と置換文字など）を除去し、Unicode正規化を適用
        finalResult = finalResult.Replace("\0", "").Replace("￿", "").Replace("\uFFFD", "");
        
        // Unicode正規化（NFC: 正規化形式C）
        finalResult = finalResult.Normalize(System.Text.NormalizationForm.FormC);
        
        return finalResult;
    }
    
    private static string BuildFormattedWordGroup(List<Word> wordGroup)
    {
        if (wordGroup.Count == 0) return "";
        
        var result = new System.Text.StringBuilder();
        FontFormatting? currentFormatting = null;
        var currentSegment = new System.Text.StringBuilder();
        
        foreach (var word in wordGroup)
        {
            var wordFormatting = AnalyzeFontFormatting([word]);
            
            // 書式が変わった場合、現在のセグメントを確定
            if (currentFormatting != null && !FormattingEqual(currentFormatting, wordFormatting))
            {
                if (currentSegment.Length > 0)
                {
                    result.Append(ApplyFormatting(currentSegment.ToString(), currentFormatting));
                    currentSegment.Clear();
                }
            }
            
            currentFormatting = wordFormatting;
            // null文字と置換文字を除去してからテキストを追加
            var cleanText = word.Text?.Replace("\0", "").Replace("￿", "").Replace("\uFFFD", "") ?? "";
            currentSegment.Append(cleanText);
        }
        
        // 最後のセグメントを処理
        if (currentSegment.Length > 0 && currentFormatting != null)
        {
            result.Append(ApplyFormatting(currentSegment.ToString(), currentFormatting));
        }
        
        return result.ToString();
    }
    
    private static bool FormattingEqual(FontFormatting a, FontFormatting b)
    {
        return a.IsBold == b.IsBold && a.IsItalic == b.IsItalic;
    }
    
    private static string BuildFormattedTextSimple(List<List<Word>> mergedWords)
    {
        if (mergedWords.Count == 0) return "";
        
        var result = new System.Text.StringBuilder();
        
        foreach (var wordGroup in mergedWords)
        {
            if (wordGroup.Count == 0) continue;
            
            var groupText = string.Join("", wordGroup.Select(w => w.Text));
            if (string.IsNullOrWhiteSpace(groupText)) continue;
            
            // 簡単なフォント検出 - 太字のみ
            bool isBold = wordGroup.Any(w => 
            {
                var fontName = w.FontName?.ToLowerInvariant() ?? "";
                return fontName.Contains("w5") || fontName.Contains("bold");
            });
            
            // 太字の場合のみフォーマットを適用
            var formattedText = isBold ? $"**{groupText}**" : groupText;
            
            if (result.Length > 0) result.Append(" ");
            result.Append(formattedText);
        }
        
        return result.ToString();
    }
    
    private static FontFormatting AnalyzeFontFormatting(List<Word> words)
    {
        var formatting = new FontFormatting();
        
        foreach (var word in words)
        {
            var fontName = word.FontName?.ToLowerInvariant() ?? "";
            
            // 日本語PDFではイタリックがフォント名に反映されない場合が多い
            // そのため、太字検出に重点を置く
            
            // 改良されたフォント検出パターン
            
            // 太字判定：より包括的なパターン
            var boldPattern = @"(bold|black|heavy|semibold|demibold|medium|[6789]00|w[5-9])";
            if (System.Text.RegularExpressions.Regex.IsMatch(fontName, boldPattern))
            {
                formatting.IsBold = true;
            }
            
            // 斜体判定：より包括的で柔軟なパターン（大文字小文字を無視）
            var italicPattern = @"(italic|oblique|slanted|cursive|emphasis|stress|kursiv|inclined|tilted)";
            if (System.Text.RegularExpressions.Regex.IsMatch(fontName, italicPattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase))
            {
                formatting.IsItalic = true;
            }
            
            // フォント名に基づく追加判定 - より幅広い検出
            if (fontName.Contains("-italic") || fontName.Contains("_italic") || 
                fontName.EndsWith("-i") || fontName.EndsWith("_i") ||
                fontName.Contains("-oblique") || fontName.Contains("-slant") ||
                fontName.Contains("italic") || fontName.Contains("oblique") ||
                fontName.Contains("italicmt") || fontName.Contains("-it") || 
                fontName.EndsWith("it") || fontName.EndsWith("-i.ttf") ||
                fontName.Contains("minion") && fontName.Contains("it"))
            {
                formatting.IsItalic = true;
            }
            
            
            // PostScriptフォント名のサブセットタグを除去して再判定
            // 例: "EOODIA+Poetica-Bold" -> "Poetica-Bold"
            var cleanedFontName = System.Text.RegularExpressions.Regex.Replace(fontName, @"^[A-Z]{6}\+", "");
            if (cleanedFontName != fontName)
            {
                if (System.Text.RegularExpressions.Regex.IsMatch(cleanedFontName, boldPattern))
                {
                    formatting.IsBold = true;
                }
                if (System.Text.RegularExpressions.Regex.IsMatch(cleanedFontName, italicPattern))
                {
                    formatting.IsItalic = true;
                }
            }
        }
        
        return formatting;
    }
    
    private static string ApplyFormatting(string text, FontFormatting formatting)
    {
        // null文字と置換文字を除去
        text = text.Replace("\0", "").Replace("￿", "").Replace("\uFFFD", "");
        
        if (formatting.IsBold && formatting.IsItalic)
        {
            return $"***{text}***";
        }
        else if (formatting.IsBold)
        {
            return $"**{text}**";
        }
        else if (formatting.IsItalic)
        {
            return $"*{text}*";
        }
        
        return text;
    }


    private static bool HasConsistentTableStructure(List<DocumentElement> candidates)
    {
        if (candidates.Count < 2) return false;

        var columnCounts = candidates.Select(c => 
            c.Content.Split(new[] { ' ', '\t', '　' }, StringSplitOptions.RemoveEmptyEntries).Length
        ).ToList();

        // 列数の一貫性をチェック（±1の差を許容）
        var minColumns = columnCounts.Min();
        var maxColumns = columnCounts.Max();
        
        return maxColumns - minColumns <= 1 && minColumns >= 2;
    }

    private static double CalculateAverageWordGap(List<Word> words)
    {
        if (words.Count < 2) return 0;

        var gaps = new List<double>();
        for (int i = 1; i < words.Count; i++)
        {
            gaps.Add(words[i].BoundingBox.Left - words[i-1].BoundingBox.Right);
        }

        return gaps.Average();
    }

    private static GraphicsInfo ExtractGraphicsInfo(Page page)
    {
        var graphicsInfo = new GraphicsInfo();
        
        try
        {
            var horizontalLines = new List<LineSegment>();
            var verticalLines = new List<LineSegment>();
            var rectangles = new List<UglyToad.PdfPig.Core.PdfRectangle>();
            
            // 現在のPdfPigバージョンでは直接的なパス操作が制限されるため
            // 代替アプローチとして、ページの画像情報とWordの位置から推測する
            
            // 将来的に詳細な図形情報取得が可能になるまでの基本実装
            var words = page.GetWords();
            
            // 単語の配置パターンから表構造を推測
            var tableStructure = InferTableStructureFromWordPositions(words);
            horizontalLines.AddRange(tableStructure.horizontalLines);
            verticalLines.AddRange(tableStructure.verticalLines);
            rectangles.AddRange(tableStructure.rectangles);
            
            graphicsInfo.HorizontalLines = horizontalLines;
            graphicsInfo.VerticalLines = verticalLines;
            graphicsInfo.Rectangles = rectangles;
        }
        catch
        {
            // 図形情報の抽出に失敗した場合は空の情報を返す
        }

        return graphicsInfo;
    }

    private static (List<LineSegment> horizontalLines, List<LineSegment> verticalLines, List<UglyToad.PdfPig.Core.PdfRectangle> rectangles) InferTableStructureFromWordPositions(IEnumerable<UglyToad.PdfPig.Content.Word> words)
    {
        var horizontalLines = new List<LineSegment>();
        var verticalLines = new List<LineSegment>();
        var rectangles = new List<UglyToad.PdfPig.Core.PdfRectangle>();
        
        try
        {
            var wordList = words.ToList();
            if (wordList.Count < 4) return (horizontalLines, verticalLines, rectangles);
            
            // 単語をY座標でグループ化（行を識別）
            var tolerance = 5.0;
            var rows = GroupWordsByRows(wordList, tolerance);
            
            // 各行で単語が整列している場合、表構造を推測
            if (rows.Count >= 2)
            {
                // 行間の境界線を推測（水平線）
                for (int i = 0; i < rows.Count - 1; i++)
                {
                    var currentRow = rows[i];
                    var nextRow = rows[i + 1];
                    
                    var currentBottom = currentRow.Min(w => w.BoundingBox.Bottom);
                    var nextTop = nextRow.Max(w => w.BoundingBox.Top);
                    var lineY = (currentBottom + nextTop) / 2;
                    
                    var leftMost = Math.Min(currentRow.Min(w => w.BoundingBox.Left), nextRow.Min(w => w.BoundingBox.Left));
                    var rightMost = Math.Max(currentRow.Max(w => w.BoundingBox.Right), nextRow.Max(w => w.BoundingBox.Right));
                    
                    horizontalLines.Add(new LineSegment
                    {
                        From = new UglyToad.PdfPig.Core.PdfPoint(leftMost, lineY),
                        To = new UglyToad.PdfPig.Core.PdfPoint(rightMost, lineY)
                    });
                }
                
                // 列の境界線を推測（垂直線）
                var allColumns = DetectColumnBoundaries(rows);
                foreach (var columnX in allColumns)
                {
                    var topMost = rows.Max(row => row.Max(w => w.BoundingBox.Top));
                    var bottomMost = rows.Min(row => row.Min(w => w.BoundingBox.Bottom));
                    
                    verticalLines.Add(new LineSegment
                    {
                        From = new UglyToad.PdfPig.Core.PdfPoint(columnX, bottomMost),
                        To = new UglyToad.PdfPig.Core.PdfPoint(columnX, topMost)
                    });
                }
            }
        }
        catch
        {
            // エラーが発生した場合は空のリストを返す
        }
        
        return (horizontalLines, verticalLines, rectangles);
    }
    
    private static List<List<UglyToad.PdfPig.Content.Word>> GroupWordsByRows(List<UglyToad.PdfPig.Content.Word> words, double tolerance)
    {
        var rows = new List<List<UglyToad.PdfPig.Content.Word>>();
        
        foreach (var word in words.OrderByDescending(w => w.BoundingBox.Bottom))
        {
            bool addedToExistingRow = false;
            
            foreach (var row in rows)
            {
                var rowY = row.First().BoundingBox.Bottom;
                if (Math.Abs(word.BoundingBox.Bottom - rowY) <= tolerance)
                {
                    row.Add(word);
                    addedToExistingRow = true;
                    break;
                }
            }
            
            if (!addedToExistingRow)
            {
                rows.Add(new List<UglyToad.PdfPig.Content.Word> { word });
            }
        }
        
        // 各行を左から右へソート
        foreach (var row in rows)
        {
            row.Sort((w1, w2) => w1.BoundingBox.Left.CompareTo(w2.BoundingBox.Left));
        }
        
        return rows;
    }
    
    private static List<double> DetectColumnBoundaries(List<List<UglyToad.PdfPig.Content.Word>> rows)
    {
        var boundaries = new HashSet<double>();
        
        foreach (var row in rows)
        {
            foreach (var word in row)
            {
                boundaries.Add(word.BoundingBox.Left);
                boundaries.Add(word.BoundingBox.Right);
            }
        }
        
        return boundaries.OrderBy(x => x).ToList();
    }
    
    private static bool IsHorizontalLine(LineSegment line)
    {
        var tolerance = 2.0; // Y座標の許容差
        return Math.Abs(line.From.Y - line.To.Y) <= tolerance && Math.Abs(line.From.X - line.To.X) > tolerance;
    }
    
    private static bool IsVerticalLine(LineSegment line)
    {
        var tolerance = 2.0; // X座標の許容差
        return Math.Abs(line.From.X - line.To.X) <= tolerance && Math.Abs(line.From.Y - line.To.Y) > tolerance;
    }
    
    private static (List<LineSegment> horizontal, List<LineSegment> vertical) ExtractLinesFromRectangle(UglyToad.PdfPig.Core.PdfRectangle rect)
    {
        var horizontal = new List<LineSegment>();
        var vertical = new List<LineSegment>();
        
        // 矩形の4辺を線分として抽出
        
        // 上辺（水平線）
        horizontal.Add(new LineSegment
        {
            From = new UglyToad.PdfPig.Core.PdfPoint(rect.Left, rect.Top),
            To = new UglyToad.PdfPig.Core.PdfPoint(rect.Right, rect.Top)
        });
        
        // 下辺（水平線）
        horizontal.Add(new LineSegment
        {
            From = new UglyToad.PdfPig.Core.PdfPoint(rect.Left, rect.Bottom),
            To = new UglyToad.PdfPig.Core.PdfPoint(rect.Right, rect.Bottom)
        });
        
        // 左辺（垂直線）
        vertical.Add(new LineSegment
        {
            From = new UglyToad.PdfPig.Core.PdfPoint(rect.Left, rect.Bottom),
            To = new UglyToad.PdfPig.Core.PdfPoint(rect.Left, rect.Top)
        });
        
        // 右辺（垂直線）
        vertical.Add(new LineSegment
        {
            From = new UglyToad.PdfPig.Core.PdfPoint(rect.Right, rect.Bottom),
            To = new UglyToad.PdfPig.Core.PdfPoint(rect.Right, rect.Top)
        });
        
        return (horizontal, vertical);
    }

    private static List<DocumentElement> PostProcessTableDetection(List<DocumentElement> elements, GraphicsInfo graphicsInfo)
    {
        var result = new List<DocumentElement>();
        
        // まず線情報からテーブル領域を特定
        var tableRegions = IdentifyTableRegions(elements, graphicsInfo);
        
        var i = 0;
        while (i < elements.Count)
        {
            var current = elements[i];
            
            // 現在の要素がテーブル領域内にあるかチェック
            var tableRegion = tableRegions.FirstOrDefault(region => region.Contains(current));
            
            if (tableRegion != null)
            {
                // テーブル領域内の要素をグループ化してテーブル行として処理
                var tableElements = ProcessTableRegion(tableRegion, graphicsInfo);
                result.AddRange(tableElements);
                
                // テーブル領域の要素数分進める
                i += tableRegion.Elements.Count;
            }
            else
            {
                // 通常の単一要素処理
                var tableCandidate = FindTableSequence(elements, i, graphicsInfo);
                
                if (tableCandidate.Count >= 2)
                {
                    foreach (var candidate in tableCandidate)
                    {
                        candidate.Type = ElementType.TableRow;
                        result.Add(candidate);
                    }
                    i += tableCandidate.Count;
                }
                else
                {
                    result.Add(current);
                    i++;
                }
            }
        }

        return result;
    }

    private static List<DocumentElement> FindTableSequence(List<DocumentElement> elements, int startIndex, GraphicsInfo graphicsInfo)
    {
        var tableCandidate = new List<DocumentElement>();
        var i = startIndex;

        while (i < elements.Count)
        {
            var current = elements[i];
            
            // 空の要素はスキップ
            if (current.Type == ElementType.Empty)
            {
                i++;
                continue;
            }
            
            // ヘッダー要素はテーブル行から除外
            if (current.Type == ElementType.Header)
            {
                // ヘッダーに遭遇したらテーブル検出を終了
                break;
            }
            
            // テーブル行の可能性をチェック（図形情報も考慮）
            if (IsLikelyTableRow(current, graphicsInfo))
            {
                tableCandidate.Add(current);
            }
            else
            {
                // 表形式でない行に遭遇したら終了
                break;
            }
            
            i++;
        }

        // 構造の一貫性をチェック
        if (tableCandidate.Count >= 2 && HasConsistentTableStructure(tableCandidate))
        {
            return tableCandidate;
        }

        return new List<DocumentElement>();
    }

    private static bool IsLikelyTableRow(DocumentElement element, GraphicsInfo graphicsInfo)
    {
        var text = element.Content.Trim();
        
        // 明らかにヘッダーではない行
        if (!text.EndsWith("。") && !text.EndsWith("."))
        {
            // 複数の単語/要素に分かれている
            var parts = text.Split(new[] { ' ', '\t', '　' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length >= 2) // 複数列のテーブル
            {
                // 図形情報による補強判定
                if (HasTableBoundaries(element, graphicsInfo))
                {
                    return true;
                }

                // 単語間に適切な間隔がある（Words情報を使用）
                if (element.Words != null && element.Words.Count >= 2)
                {
                    var avgGap = CalculateAverageWordGap(element.Words);
                    return avgGap > 10; // 10pt以上の間隔
                }
                return true;
            }
        }

        return false;
    }

    private static bool HasTableBoundaries(DocumentElement element, GraphicsInfo graphicsInfo)
    {
        if (element.Words == null || element.Words.Count == 0) return false;

        var elementBounds = new UglyToad.PdfPig.Core.PdfRectangle(
            element.Words.Min(w => w.BoundingBox.Left),
            element.Words.Min(w => w.BoundingBox.Bottom),
            element.Words.Max(w => w.BoundingBox.Right),
            element.Words.Max(w => w.BoundingBox.Top)
        );

        // 要素の周囲に水平線があるかチェック
        var hasHorizontalBoundary = graphicsInfo.HorizontalLines.Any(line =>
            Math.Abs(line.From.Y - elementBounds.Bottom) < 5 ||
            Math.Abs(line.From.Y - elementBounds.Top) < 5);

        // 要素の周囲に垂直線があるかチェック
        var hasVerticalBoundary = graphicsInfo.VerticalLines.Any(line =>
            line.From.X >= elementBounds.Left - 5 && line.From.X <= elementBounds.Right + 5 &&
            line.From.Y <= elementBounds.Top + 5 && line.To.Y >= elementBounds.Bottom - 5);

        return hasHorizontalBoundary || hasVerticalBoundary;
    }

    private static List<TableRegion> IdentifyTableRegions(List<DocumentElement> elements, GraphicsInfo graphicsInfo)
    {
        var tableRegions = new List<TableRegion>();
        
        // 強化された罫線ベーステーブル検出
        if (HasSufficientBordersForTable(graphicsInfo))
        {
            var borderBasedRegions = DetectTablesByBorders(elements, graphicsInfo);
            tableRegions.AddRange(borderBasedRegions);
        }
        
        // 従来のグリッドベース検出（厳密な矩形グリッドの場合）
        if (graphicsInfo.HorizontalLines.Count >= 2 && graphicsInfo.VerticalLines.Count >= 2)
        {
            var gridCells = DetectGridCells(graphicsInfo);
            
            foreach (var cell in gridCells)
            {
                var cellElements = elements.Where(element => IsElementInCell(element, cell)).ToList();
                
                if (cellElements.Count > 0)
                {
                    var region = new TableRegion
                    {
                        Bounds = cell,
                        Elements = cellElements
                    };
                    tableRegions.Add(region);
                }
            }
        }
        
        // 線ベースで十分なテーブルが検出されない場合、ギャップベース検出を併用
        if (tableRegions.Count == 0)
        {
            var gapBasedRegions = DetectTableByGapAnalysis(elements);
            tableRegions.AddRange(gapBasedRegions);
        }
        
        return MergeAdjacentTableRegions(tableRegions);
    }

    private static List<UglyToad.PdfPig.Core.PdfRectangle> DetectGridCells(GraphicsInfo graphicsInfo)
    {
        var cells = new List<UglyToad.PdfPig.Core.PdfRectangle>();
        
        // 水平線と垂直線をソート
        var hLines = graphicsInfo.HorizontalLines.OrderBy(l => l.From.Y).ToList();
        var vLines = graphicsInfo.VerticalLines.OrderBy(l => l.From.X).ToList();
        
        // 隣接する線の組み合わせでセルを作成
        for (int i = 0; i < hLines.Count - 1; i++)
        {
            for (int j = 0; j < vLines.Count - 1; j++)
            {
                var topLine = hLines[i + 1];
                var bottomLine = hLines[i];
                var leftLine = vLines[j];
                var rightLine = vLines[j + 1];
                
                // 線が交差してセルを形成するかチェック
                if (LinesFormCell(topLine, bottomLine, leftLine, rightLine))
                {
                    var cell = new UglyToad.PdfPig.Core.PdfRectangle(
                        leftLine.From.X,
                        bottomLine.From.Y,
                        rightLine.From.X,
                        topLine.From.Y
                    );
                    cells.Add(cell);
                }
            }
        }
        
        return cells;
    }

    private static bool LinesFormCell(LineSegment top, LineSegment bottom, LineSegment left, LineSegment right)
    {
        var tolerance = 5.0;
        
        // 線が適切な位置にあるかチェック
        return Math.Abs(top.From.Y - top.To.Y) < tolerance && // 水平線
               Math.Abs(bottom.From.Y - bottom.To.Y) < tolerance && // 水平線
               Math.Abs(left.From.X - left.To.X) < tolerance && // 垂直線
               Math.Abs(right.From.X - right.To.X) < tolerance && // 垂直線
               top.From.Y > bottom.From.Y && // 上下関係
               right.From.X > left.From.X; // 左右関係
    }

    private static bool IsElementInCell(DocumentElement element, UglyToad.PdfPig.Core.PdfRectangle cell)
    {
        if (element.Words == null || element.Words.Count == 0) return false;
        
        var elementBounds = new UglyToad.PdfPig.Core.PdfRectangle(
            element.Words.Min(w => w.BoundingBox.Left),
            element.Words.Min(w => w.BoundingBox.Bottom),
            element.Words.Max(w => w.BoundingBox.Right),
            element.Words.Max(w => w.BoundingBox.Top)
        );
        
        // 要素がセル内に含まれるかチェック（マージンを考慮）
        var margin = 2.0;
        return elementBounds.Left >= cell.Left - margin &&
               elementBounds.Right <= cell.Right + margin &&
               elementBounds.Bottom >= cell.Bottom - margin &&
               elementBounds.Top <= cell.Top + margin;
    }

    private static List<TableRegion> MergeAdjacentTableRegions(List<TableRegion> regions)
    {
        // 隣接するテーブル領域をマージして、より大きなテーブルを形成
        // 簡略化のため、現在は元のリストをそのまま返す
        return regions;
    }

    private static List<DocumentElement> ProcessTableRegion(TableRegion region, GraphicsInfo graphicsInfo)
    {
        var result = new List<DocumentElement>();
        
        // より柔軟な行グループ化を使用
        var rowGroups = GroupElementsByFlexibleRows(region.Elements);
        
        foreach (var rowElements in rowGroups)
        {
            // 各行内の要素を左から右へソート
            var sortedElements = rowElements.OrderBy(e => e.Words?.FirstOrDefault()?.BoundingBox.Left ?? 0).ToList();
            
            // 要素を列に分割
            var columns = GroupElementsIntoColumns(sortedElements);
            
            if (columns.Count > 0 && columns.Any(col => col.Any(e => !string.IsNullOrWhiteSpace(e.Content))))
            {
                var cellContents = columns.Select(col => 
                {
                    var content = string.Join(" ", col.Select(e => e.Content.Trim()).Where(c => !string.IsNullOrWhiteSpace(c)));
                    return content;
                }).ToList();
                
                var tableRow = new DocumentElement
                {
                    Type = ElementType.TableRow,
                    Content = string.Join(" | ", cellContents),
                    Words = sortedElements.SelectMany(e => e.Words ?? []).ToList()
                };
                result.Add(tableRow);
            }
        }
        
        return result;
    }
    
    private static List<List<DocumentElement>> GroupElementsByFlexibleRows(List<DocumentElement> elements)
    {
        if (!elements.Any()) return [];
        
        var rows = new List<List<DocumentElement>>();
        var tolerance = 5.0; // Y座標の許容差
        
        // Y座標でソート（上から下へ）
        var sortedElements = elements.OrderByDescending(e => e.Words?.FirstOrDefault()?.BoundingBox.Bottom ?? 0).ToList();
        
        var currentRow = new List<DocumentElement> { sortedElements[0] };
        var currentY = sortedElements[0].Words?.FirstOrDefault()?.BoundingBox.Bottom ?? 0;
        
        for (int i = 1; i < sortedElements.Count; i++)
        {
            var element = sortedElements[i];
            var elementY = element.Words?.FirstOrDefault()?.BoundingBox.Bottom ?? 0;
            
            if (Math.Abs(elementY - currentY) <= tolerance)
            {
                // 同じ行
                currentRow.Add(element);
            }
            else
            {
                // 新しい行
                if (currentRow.Count > 0)
                {
                    rows.Add(currentRow);
                }
                currentRow = new List<DocumentElement> { element };
                currentY = elementY;
            }
        }
        
        if (currentRow.Count > 0)
        {
            rows.Add(currentRow);
        }
        
        return rows;
    }
    
    private static List<List<DocumentElement>> GroupElementsIntoColumns(List<DocumentElement> rowElements)
    {
        if (!rowElements.Any()) return [];
        
        var columns = new List<List<DocumentElement>>();
        var threshold = 30.0; // 列分離の閾値
        
        foreach (var element in rowElements)
        {
            var elementX = element.Words?.FirstOrDefault()?.BoundingBox.Left ?? 0;
            
            // 既存の列に追加できるかチェック
            bool addedToExistingColumn = false;
            
            for (int i = 0; i < columns.Count; i++)
            {
                var columnX = columns[i].FirstOrDefault()?.Words?.FirstOrDefault()?.BoundingBox.Left ?? 0;
                
                if (Math.Abs(elementX - columnX) <= threshold)
                {
                    columns[i].Add(element);
                    addedToExistingColumn = true;
                    break;
                }
            }
            
            if (!addedToExistingColumn)
            {
                // 新しい列を作成
                columns.Add(new List<DocumentElement> { element });
            }
        }
        
        // 列をX座標でソート
        columns = columns.OrderBy(col => col.FirstOrDefault()?.Words?.FirstOrDefault()?.BoundingBox.Left ?? 0).ToList();
        
        return columns;
    }

    private static Dictionary<UglyToad.PdfPig.Core.PdfRectangle, List<DocumentElement>> GroupElementsByCells(TableRegion region, GraphicsInfo graphicsInfo)
    {
        var cellGroups = new Dictionary<UglyToad.PdfPig.Core.PdfRectangle, List<DocumentElement>>();
        
        // 各要素を適切なセルに割り当て
        foreach (var element in region.Elements)
        {
            var cell = FindContainingCell(element, graphicsInfo);
            if (cell.HasValue)
            {
                if (!cellGroups.ContainsKey(cell.Value))
                {
                    cellGroups[cell.Value] = new List<DocumentElement>();
                }
                cellGroups[cell.Value].Add(element);
            }
        }
        
        return cellGroups;
    }

    private static UglyToad.PdfPig.Core.PdfRectangle? FindContainingCell(DocumentElement element, GraphicsInfo graphicsInfo)
    {
        if (element.Words == null || element.Words.Count == 0) return null;
        
        var elementCenter = new UglyToad.PdfPig.Core.PdfPoint(
            element.Words.Average(w => w.BoundingBox.Left + w.BoundingBox.Width / 2),
            element.Words.Average(w => w.BoundingBox.Bottom + w.BoundingBox.Height / 2)
        );
        
        // 最も近いセルを検索
        var gridCells = DetectGridCells(graphicsInfo);
        return gridCells.FirstOrDefault(cell => 
            elementCenter.X >= cell.Left && elementCenter.X <= cell.Right &&
            elementCenter.Y >= cell.Bottom && elementCenter.Y <= cell.Top);
    }

    private static List<DocumentElement> PostProcessElementClassification(List<DocumentElement> elements, FontAnalysis fontAnalysis)
    {
        var result = new List<DocumentElement>(elements);
        
        // コンテキスト情報を使用して要素分類を改善
        for (int i = 0; i < result.Count; i++)
        {
            var current = result[i];
            var previous = i > 0 ? result[i - 1] : null;
            var next = i < result.Count - 1 ? result[i + 1] : null;
            
            // 段落の継続性チェック
            if (current.Type == ElementType.Paragraph && 
                previous?.Type == ElementType.Paragraph &&
                ShouldMergeParagraphs(previous, current))
            {
                // 段落を統合
                previous.Content = previous.Content.Trim() + " " + current.Content.Trim();
                previous.Words.AddRange(current.Words);
                result.RemoveAt(i);
                i--; // インデックス調整
                continue;
            }
            
            // ヘッダー分類の改善
            if (current.Type == ElementType.Paragraph && 
                CouldBeHeader(current, previous, next, fontAnalysis))
            {
                current.Type = ElementType.Header;
            }
            
            // リストアイテムの継続性チェック
            if (current.Type == ElementType.Paragraph &&
                previous?.Type == ElementType.ListItem &&
                IsListContinuation(current, previous))
            {
                current.Type = ElementType.ListItem;
            }
            
            // テーブル行の改善
            if (current.Type == ElementType.Paragraph &&
                IsPartOfTableSequence(current, result, i))
            {
                current.Type = ElementType.TableRow;
            }
        }
        
        return result;
    }
    
    private static bool ShouldMergeParagraphs(DocumentElement previous, DocumentElement current)
    {
        var prevText = previous.Content.Trim();
        var currText = current.Content.Trim();
        
        // フォントサイズが類似している
        var fontSizeDiff = Math.Abs(previous.FontSize - current.FontSize);
        if (fontSizeDiff > 2.0) return false;
        
        // インデント状況が類似している
        var indentDiff = Math.Abs(previous.LeftMargin - current.LeftMargin);
        if (indentDiff > 10.0) return false;
        
        // 前の段落が文で終わっていない（継続の可能性）
        if (!prevText.EndsWith("。") && !prevText.EndsWith(".") && 
            !prevText.EndsWith("!") && !prevText.EndsWith("?"))
        {
            return true;
        }
        
        // 現在の段落が小文字や続きを示す語句で始まる
        if (currText.StartsWith("と") || currText.StartsWith("が") || 
            currText.StartsWith("で") || currText.StartsWith("に"))
        {
            return true;
        }
        
        return false;
    }
    
    private static bool CouldBeHeader(DocumentElement current, DocumentElement? previous, DocumentElement? next, FontAnalysis fontAnalysis)
    {
        var text = current.Content.Trim();
        var cleanText = ExtractCleanTextForAnalysis(text);
        
        // 明らかにヘッダーではないパターンを除外
        if (cleanText.EndsWith("。") || cleanText.EndsWith(".") || cleanText.Contains("、")) return false;
        
        // フォントサイズがベースより大きいかチェック
        var fontSizeRatio = current.FontSize / fontAnalysis.BaseFontSize;
        
        // 強力なヘッダー指標：太字でフォーマットされた短いテキスト
        bool isBoldFormatted = text.Contains("**") && cleanText.Length <= 20;
        
        // 短い名詞句パターン（テーブルヘッダーなど）
        bool isShortNounPhrase = cleanText.Length <= 15 && !cleanText.Contains(" ") && !cleanText.All(char.IsDigit);
        
        // "基本的な〜"、"〜を含む〜"パターン
        bool isDescriptivePhrase = System.Text.RegularExpressions.Regex.IsMatch(cleanText, @"^.{1,8}[なの].{1,10}$") ||
                                  System.Text.RegularExpressions.Regex.IsMatch(cleanText, @"^.{1,8}を含む.{1,10}$") ||
                                  System.Text.RegularExpressions.Regex.IsMatch(cleanText, @"^.{1,8}[なの].{1,8}[ーテ].{1,5}$");
        
        // 前後の要素との関係を考慮
        bool hasSpaceBefore = previous == null || previous.Type != ElementType.Paragraph;
        bool hasSpaceAfter = next == null || next.Type != ElementType.Paragraph;
        
        // 強力なヘッダー判定
        if (isBoldFormatted || (fontSizeRatio > 1.15 && (isShortNounPhrase || isDescriptivePhrase)))
        {
            return true;
        }
        
        // 従来のヘッダー的なパターン
        if (System.Text.RegularExpressions.Regex.IsMatch(cleanText, @"^\d+\.\s*\w+") || // "1. 概要"
            System.Text.RegularExpressions.Regex.IsMatch(cleanText, @"^第\d+章") ||     // "第1章"
            System.Text.RegularExpressions.Regex.IsMatch(cleanText, @"^\d+\s*[-－]\s*\w+")) // "1 - 概要"
        {
            return true;
        }
        
        // フォントサイズが十分に大きく、前後に空白がある場合
        if (fontSizeRatio > 1.2 && (hasSpaceBefore || hasSpaceAfter))
        {
            return true;
        }
        
        return false;
    }
    
    private static bool IsListContinuation(DocumentElement current, DocumentElement previous)
    {
        var currText = current.Content.Trim();
        var prevText = previous.Content.Trim();
        
        // インデントが類似している
        var indentDiff = Math.Abs(current.LeftMargin - previous.LeftMargin);
        if (indentDiff > 15.0) return false;
        
        // 前のリストアイテムが完了していない
        if (!prevText.EndsWith("。") && !prevText.EndsWith("."))
        {
            return true;
        }
        
        // 継続を示す語句で始まる
        if (currText.StartsWith("また") || currText.StartsWith("さらに") ||
            currText.StartsWith("ただし") || currText.StartsWith("なお"))
        {
            return true;
        }
        
        return false;
    }
    
    private static bool IsPartOfTableSequence(DocumentElement current, List<DocumentElement> allElements, int currentIndex)
    {
        // 前後にテーブル行がある場合
        var hasPrevTable = currentIndex > 0 && allElements[currentIndex - 1].Type == ElementType.TableRow;
        var hasNextTable = currentIndex < allElements.Count - 1 && allElements[currentIndex + 1].Type == ElementType.TableRow;
        
        if (!hasPrevTable && !hasNextTable) return false;
        
        var text = current.Content.Trim();
        
        // テーブル行っぽい特徴
        var parts = text.Split(new[] { ' ', '\t', '　' }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length < 2) return false;
        
        // 数値が多い
        var numericParts = parts.Count(p => double.TryParse(p, out _) || 
            p.Contains("%") || p.Contains(",") || p.Contains("¥") || p.Contains("$"));
        
        if ((double)numericParts / parts.Length > 0.3) return true;
        
        // 単語間のギャップパターンが一致する
        if (current.Words != null && current.Words.Count >= 3)
        {
            var gaps = new List<double>();
            for (int i = 1; i < current.Words.Count; i++)
            {
                gaps.Add(current.Words[i].BoundingBox.Left - current.Words[i-1].BoundingBox.Right);
            }
            
            var largeGaps = gaps.Count(g => g > 15);
            if (largeGaps >= 2) return true;
        }
        
        return false;
    }

    private static List<TableRegion> DetectTableByGapAnalysis(List<DocumentElement> elements)
    {
        var tableRegions = new List<TableRegion>();
        
        // Y座標でグループ化して行を識別
        var rowGroups = elements
            .Where(e => e.Type == ElementType.Paragraph || e.Type == ElementType.TableRow)
            .GroupBy(e => Math.Round(e.Words?.FirstOrDefault()?.BoundingBox.Bottom ?? 0, 1))
            .Where(g => g.Count() >= 2) // 最低2つの要素がある行のみ
            .OrderByDescending(g => g.Key)
            .ToList();
        
        if (rowGroups.Count < 2) return tableRegions;
        
        var tableRows = new List<List<DocumentElement>>();
        
        foreach (var rowGroup in rowGroups)
        {
            var rowElements = rowGroup.OrderBy(e => e.Words?.FirstOrDefault()?.BoundingBox.Left ?? 0).ToList();
            
            // 水平ギャップ分析で列を検出
            if (HasConsistentColumnStructure(rowElements))
            {
                tableRows.Add(rowElements);
            }
        }
        
        // 連続する行がテーブルを形成するかチェック
        if (tableRows.Count >= 2)
        {
            // 全ての行を含む単一のテーブル領域を作成
            var allElements = tableRows.SelectMany(row => row).ToList();
            
            if (allElements.Count > 0)
            {
                var minX = allElements.SelectMany(e => e.Words ?? []).Min(w => w.BoundingBox.Left);
                var maxX = allElements.SelectMany(e => e.Words ?? []).Max(w => w.BoundingBox.Right);
                var minY = allElements.SelectMany(e => e.Words ?? []).Min(w => w.BoundingBox.Bottom);
                var maxY = allElements.SelectMany(e => e.Words ?? []).Max(w => w.BoundingBox.Top);
                
                var bounds = new UglyToad.PdfPig.Core.PdfRectangle(minX, minY, maxX, maxY);
                
                var region = new TableRegion
                {
                    Bounds = bounds,
                    Elements = allElements
                };
                tableRegions.Add(region);
            }
        }
        
        return tableRegions;
    }
    
    private static bool HasConsistentColumnStructure(List<DocumentElement> rowElements)
    {
        if (rowElements.Count < 2) return false;
        
        // 要素間の水平ギャップを計算
        var gaps = new List<double>();
        for (int i = 0; i < rowElements.Count - 1; i++)
        {
            var current = rowElements[i].Words?.LastOrDefault()?.BoundingBox.Right ?? 0;
            var next = rowElements[i + 1].Words?.FirstOrDefault()?.BoundingBox.Left ?? 0;
            var gap = next - current;
            if (gap > 5) gaps.Add(gap); // 意味のあるギャップのみ
        }
        
        // 十分な数のギャップがある場合はテーブル行と判定
        return gaps.Count >= 1 && gaps.Any(g => g > 15);
    }
    
    private static bool HasSufficientBordersForTable(GraphicsInfo graphicsInfo)
    {
        // 罫線が表として認識できる最小条件
        // 最低限1本以上の水平線と垂直線があれば表の可能性がある
        return graphicsInfo.HorizontalLines.Count >= 1 && graphicsInfo.VerticalLines.Count >= 1;
    }
    
    private static List<TableRegion> DetectTablesByBorders(List<DocumentElement> elements, GraphicsInfo graphicsInfo)
    {
        var tableRegions = new List<TableRegion>();
        
        try
        {
            // 罫線で囲まれた領域を特定
            var borderedRegions = FindBorderedRegions(graphicsInfo);
            
            foreach (var region in borderedRegions)
            {
                // 領域内のテキスト要素を取得
                var regionElements = elements.Where(e => IsElementInRegion(e, region)).ToList();
                
                if (regionElements.Count >= 2) // 最低2つの要素があれば表として認識
                {
                    var tableRegion = new TableRegion
                    {
                        Bounds = region,
                        Elements = regionElements
                    };
                    tableRegions.Add(tableRegion);
                }
            }
        }
        catch
        {
            // 罫線ベース検出でエラーが発生した場合は空のリストを返す
        }
        
        return tableRegions;
    }
    
    private static List<UglyToad.PdfPig.Core.PdfRectangle> FindBorderedRegions(GraphicsInfo graphicsInfo)
    {
        var regions = new List<UglyToad.PdfPig.Core.PdfRectangle>();
        
        // 矩形として定義された領域
        regions.AddRange(graphicsInfo.Rectangles);
        
        // 線の交点から形成される矩形領域を検出
        var intersectionRegions = DetectRegionsFromLineIntersections(graphicsInfo);
        regions.AddRange(intersectionRegions);
        
        // 重複除去
        return DeduplicateRegions(regions);
    }
    
    private static List<UglyToad.PdfPig.Core.PdfRectangle> DetectRegionsFromLineIntersections(GraphicsInfo graphicsInfo)
    {
        var regions = new List<UglyToad.PdfPig.Core.PdfRectangle>();
        
        try
        {
            // 水平線と垂直線の交点を見つけて矩形を形成
            foreach (var hLine1 in graphicsInfo.HorizontalLines)
            {
                foreach (var hLine2 in graphicsInfo.HorizontalLines)
                {
                    if (hLine1 == hLine2) continue;
                    
                    // 2本の水平線の間にある垂直線を見つける
                    var topY = Math.Max(hLine1.From.Y, hLine2.From.Y);
                    var bottomY = Math.Min(hLine1.From.Y, hLine2.From.Y);
                    
                    if (Math.Abs(topY - bottomY) < 5) continue; // 線が近すぎる場合はスキップ
                    
                    foreach (var vLine1 in graphicsInfo.VerticalLines)
                    {
                        foreach (var vLine2 in graphicsInfo.VerticalLines)
                        {
                            if (vLine1 == vLine2) continue;
                            
                            var leftX = Math.Min(vLine1.From.X, vLine2.From.X);
                            var rightX = Math.Max(vLine1.From.X, vLine2.From.X);
                            
                            if (Math.Abs(rightX - leftX) < 10) continue; // 線が近すぎる場合はスキップ
                            
                            // 4本の線で矩形が形成されるかチェック
                            if (LinesFormValidRectangle(hLine1, hLine2, vLine1, vLine2))
                            {
                                var rect = new UglyToad.PdfPig.Core.PdfRectangle(leftX, bottomY, rightX, topY);
                                regions.Add(rect);
                            }
                        }
                    }
                }
            }
        }
        catch
        {
            // 交点検出でエラーが発生した場合は継続
        }
        
        return regions;
    }
    
    private static bool LinesFormValidRectangle(LineSegment hLine1, LineSegment hLine2, LineSegment vLine1, LineSegment vLine2)
    {
        var tolerance = 5.0;
        
        // 水平線が実際に垂直線と交差するかチェック
        var h1IntersectsV1 = LineIntersects(hLine1, vLine1, tolerance);
        var h1IntersectsV2 = LineIntersects(hLine1, vLine2, tolerance);
        var h2IntersectsV1 = LineIntersects(hLine2, vLine1, tolerance);
        var h2IntersectsV2 = LineIntersects(hLine2, vLine2, tolerance);
        
        // 4つの交点があれば矩形
        return h1IntersectsV1 && h1IntersectsV2 && h2IntersectsV1 && h2IntersectsV2;
    }
    
    private static bool LineIntersects(LineSegment line1, LineSegment line2, double tolerance)
    {
        // 線分の交点計算（簡易版）
        
        // 水平線と垂直線の場合
        if (IsHorizontalLine(line1) && IsVerticalLine(line2))
        {
            var hLine = line1;
            var vLine = line2;
            
            var hLeft = Math.Min(hLine.From.X, hLine.To.X);
            var hRight = Math.Max(hLine.From.X, hLine.To.X);
            var hY = hLine.From.Y;
            
            var vBottom = Math.Min(vLine.From.Y, vLine.To.Y);
            var vTop = Math.Max(vLine.From.Y, vLine.To.Y);
            var vX = vLine.From.X;
            
            return vX >= hLeft - tolerance && vX <= hRight + tolerance &&
                   hY >= vBottom - tolerance && hY <= vTop + tolerance;
        }
        
        if (IsVerticalLine(line1) && IsHorizontalLine(line2))
        {
            return LineIntersects(line2, line1, tolerance);
        }
        
        return false;
    }
    
    private static List<UglyToad.PdfPig.Core.PdfRectangle> DeduplicateRegions(List<UglyToad.PdfPig.Core.PdfRectangle> regions)
    {
        var deduplicated = new List<UglyToad.PdfPig.Core.PdfRectangle>();
        
        foreach (var region in regions)
        {
            bool isDuplicate = false;
            
            foreach (var existing in deduplicated)
            {
                if (Math.Abs(region.Left - existing.Left) < 5 &&
                    Math.Abs(region.Right - existing.Right) < 5 &&
                    Math.Abs(region.Bottom - existing.Bottom) < 5 &&
                    Math.Abs(region.Top - existing.Top) < 5)
                {
                    isDuplicate = true;
                    break;
                }
            }
            
            if (!isDuplicate)
            {
                deduplicated.Add(region);
            }
        }
        
        return deduplicated;
    }
    
    private static bool IsElementInRegion(DocumentElement element, UglyToad.PdfPig.Core.PdfRectangle region)
    {
        if (element.Words == null || !element.Words.Any()) return false;
        
        var elementBounds = GetElementBounds(element);
        
        // 要素が領域内に含まれるかチェック（マージン許容）
        var margin = 3.0;
        return elementBounds.Left >= region.Left - margin &&
               elementBounds.Right <= region.Right + margin &&
               elementBounds.Bottom >= region.Bottom - margin &&
               elementBounds.Top <= region.Top + margin;
    }
    
    private static UglyToad.PdfPig.Core.PdfRectangle GetElementBounds(DocumentElement element)
    {
        if (element.Words == null || !element.Words.Any())
        {
            return new UglyToad.PdfPig.Core.PdfRectangle(0, 0, 0, 0);
        }
        
        var left = element.Words.Min(w => w.BoundingBox.Left);
        var right = element.Words.Max(w => w.BoundingBox.Right);
        var bottom = element.Words.Min(w => w.BoundingBox.Bottom);
        var top = element.Words.Max(w => w.BoundingBox.Top);
        
        return new UglyToad.PdfPig.Core.PdfRectangle(left, bottom, right, top);
    }
}

public class TableRegion
{
    public UglyToad.PdfPig.Core.PdfRectangle Bounds { get; set; }
    public List<DocumentElement> Elements { get; set; } = new List<DocumentElement>();
    
    public bool Contains(DocumentElement element)
    {
        return Elements.Contains(element);
    }
}

public class GraphicsInfo
{
    public List<LineSegment> HorizontalLines { get; set; } = new List<LineSegment>();
    public List<LineSegment> VerticalLines { get; set; } = new List<LineSegment>();
    public List<UglyToad.PdfPig.Core.PdfRectangle> Rectangles { get; set; } = new List<UglyToad.PdfPig.Core.PdfRectangle>();
}

public class LineSegment
{
    public UglyToad.PdfPig.Core.PdfPoint From { get; set; }
    public UglyToad.PdfPig.Core.PdfPoint To { get; set; }
}

public class FontFormatting
{
    public bool IsBold { get; set; }
    public bool IsItalic { get; set; }
}

public class FontAnalysis
{
    public double BaseFontSize { get; set; }
    public double LargeFontThreshold { get; set; }
    public double SmallFontThreshold { get; set; }
    public string DominantFont { get; set; } = "";
    public List<double> AllFontSizes { get; set; } = [];
}

public class DocumentStructure
{
    public List<DocumentElement> Elements { get; set; } = [];
    public FontAnalysis FontAnalysis { get; set; } = new FontAnalysis();
}

public class DocumentElement
{
    public ElementType Type { get; set; }
    public string Content { get; set; } = "";
    public double FontSize { get; set; }
    public double LeftMargin { get; set; }
    public bool IsIndented { get; set; }
    public List<Word> Words { get; set; } = [];
}

public enum ElementType
{
    Empty,
    Header,
    Paragraph,
    ListItem,
    TableRow,
    CodeBlock,
    QuoteBlock
}