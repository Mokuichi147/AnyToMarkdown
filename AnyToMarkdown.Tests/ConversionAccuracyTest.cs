using System.Text.RegularExpressions;

namespace AnyToMarkdown.Tests;

public class MarkdownStructureAnalysis
{
    public string FileName { get; set; } = string.Empty;
    public List<string> OriginalHeaders { get; set; } = [];
    public List<string> ConvertedHeaders { get; set; } = [];
    public List<string> OriginalSections { get; set; } = [];
    public List<string> ConvertedSections { get; set; } = [];
    public List<string> OriginalTables { get; set; } = [];
    public List<string> ConvertedTables { get; set; } = [];
    public List<string> OriginalLists { get; set; } = [];
    public List<string> ConvertedLists { get; set; } = [];
}

public class ConversionAccuracyTest
{


    [Theory]
    [InlineData("test-basic")]
    [InlineData("test-table")]
    [InlineData("test-complex")]
    [InlineData("test-japanese")]
    [InlineData("test-multiline-table")]
    [InlineData("test-complex-table")]
    [InlineData("test-advanced-document")]
    [InlineData("test-scientific-paper")]
    [InlineData("test-financial-report")]
    [InlineData("test-mixed-content")]
    [InlineData("test-comprehensive-markdown")]
    public void RoundTripConversionTest(string fileName)
    {
        // 元のMarkdownファイルを読み込み
        string originalMdPath = $"./Resources/{fileName}.md";
        string generatedPdfPath = $"./Resources/{fileName}.pdf";
        string convertedMdPath = $"./test-output-{fileName}.md";

        Assert.True(File.Exists(originalMdPath), $"Original markdown file should exist: {originalMdPath}");
        
        // PDFファイルが存在することを確認（pandocで事前生成が必要）
        if (!File.Exists(generatedPdfPath))
        {
            Assert.Fail($"PDF file not found: {generatedPdfPath}. Please run: pandoc {originalMdPath} -o {generatedPdfPath} --pdf-engine=xelatex -V CJKmainfont=\"Hiragino Sans\" -V mainfont=\"Hiragino Sans\"");
        }

        // PDFからMarkdownに変換
        var result = AnyConverter.Convert(generatedPdfPath);
        
        Assert.NotNull(result.Text);
        Assert.NotEmpty(result.Text);
        
        // 変換結果をファイルに保存
        File.WriteAllText(convertedMdPath, result.Text);
        
        // 元のMarkdownを読み込み
        string originalContent = File.ReadAllText(originalMdPath);
        
        // 基本的な内容が保持されているかチェック
        VerifyContentSimilarity(originalContent, result.Text, fileName);
        
        // 変換統計を出力
        Console.WriteLine($"[{fileName}] Original: {originalContent.Length} chars, Converted: {result.Text.Length} chars, Warnings: {result.Warnings.Count}");
    }

    private void VerifyContentSimilarity(string original, string converted, string fileName)
    {
        var analysis = AnalyzeMarkdownStructure(original, converted, fileName);
        
        // 構造評価の実施
        EvaluateStructuralSimilarity(analysis, fileName);
    }

    private static MarkdownStructureAnalysis AnalyzeMarkdownStructure(string original, string converted, string fileName)
    {
        return new MarkdownStructureAnalysis
        {
            FileName = fileName,
            OriginalHeaders = ExtractHeaders(original),
            ConvertedHeaders = ExtractHeaders(converted),
            OriginalSections = ExtractSections(original),
            ConvertedSections = ExtractSections(converted),
            OriginalTables = ExtractTables(original),
            ConvertedTables = ExtractTables(converted),
            OriginalLists = ExtractLists(original),
            ConvertedLists = ExtractLists(converted)
        };
    }

    private void EvaluateStructuralSimilarity(MarkdownStructureAnalysis analysis, string fileName)
    {
        var testResults = new List<(string TestName, bool Passed, string Details)>();

        // 1. ヘッダー構造テスト
        var headerResult = EvaluateHeaderStructureAsTest(analysis);
        testResults.Add(("Header Structure", headerResult.Passed, headerResult.Details));

        // 2. セクション構造テスト
        var sectionResult = EvaluateSectionStructureAsTest(analysis);
        testResults.Add(("Section Structure", sectionResult.Passed, sectionResult.Details));

        // 3. テーブル構造テスト
        var tableResult = EvaluateTableStructureAsTest(analysis);
        testResults.Add(("Table Structure", tableResult.Passed, tableResult.Details));

        // 4. リスト構造テスト
        var listResult = EvaluateListStructureAsTest(analysis);
        testResults.Add(("List Structure", listResult.Passed, listResult.Details));

        // 5. 強調記法テスト（太字・斜体）
        var (emphasisPassed, emphasisDetails) = EvaluateEmphasisAsTest(analysis);
        testResults.Add(("Emphasis (Bold/Italic)", emphasisPassed, emphasisDetails));

        // 6. リンク記法テスト
        var (linkPassed, linkDetails) = EvaluateLinksAsTest(analysis);
        testResults.Add(("Link Syntax", linkPassed, linkDetails));

        // 7. コード記法テスト（インライン・ブロック）
        var (codePassed, codeDetails) = EvaluateCodeAsTest(analysis);
        testResults.Add(("Code Syntax", codePassed, codeDetails));

        // 8. 引用記法テスト
        var (quotePassed, quoteDetails) = EvaluateQuotesAsTest(analysis);
        testResults.Add(("Quote Syntax", quotePassed, quoteDetails));

        // 9. その他記法テスト（水平線・エスケープ）
        var (otherPassed, otherDetails) = EvaluateOtherSyntaxAsTest(analysis);
        testResults.Add(("Other Syntax", otherPassed, otherDetails));

        // 結果の表示
        int passedTests = 0;
        int totalTests = testResults.Count;

        Console.WriteLine($"[{fileName}] Detailed Test Results:");
        foreach (var (testName, passed, details) in testResults)
        {
            string status = passed ? "PASS" : "FAIL";
            Console.WriteLine($"  {status}: {testName} - {details}");
            if (passed) passedTests++;
        }

        double passRate = (double)passedTests / totalTests * 100.0;
        Console.WriteLine($"[{fileName}] Overall pass rate: {passedTests}/{totalTests} ({passRate:F1}%)");

        // 厳格な基準: 全項目100%合格が必要
        double requiredPassRate = 100.0;
        Assert.True(passRate >= requiredPassRate, 
            $"[{fileName}] Pass rate ({passRate:F1}%) below required threshold ({requiredPassRate:F1}%)");
    }


    private static (bool Passed, string Details) EvaluateHeaderStructureAsTest(MarkdownStructureAnalysis analysis)
    {
        if (analysis.OriginalHeaders.Count == 0) 
            return (true, "No headers to validate");

        int properMarkdownHeaders = 0;
        int totalHeaders = analysis.OriginalHeaders.Count;

        Console.WriteLine($"[DEBUG] Checking {totalHeaders} headers for proper Markdown preservation:");

        foreach (var originalHeader in analysis.OriginalHeaders)
        {
            var headerText = ExtractHeaderText(originalHeader);
            var headerLevel = GetHeaderLevel(originalHeader);

            Console.WriteLine($"[DEBUG] Looking for: '{originalHeader}' (level {headerLevel})");

            // 厳密な基準: 正確なMarkdownヘッダーとして保持されているかをチェック
            bool exactMatch = analysis.ConvertedHeaders.Any(h => 
                GetHeaderLevel(h) == headerLevel && 
                ExtractHeaderText(h).Equals(headerText, StringComparison.OrdinalIgnoreCase));

            // レベルは違うがMarkdownヘッダーとして保持されている
            bool headerWithDifferentLevel = !exactMatch && analysis.ConvertedHeaders.Any(h => 
                ExtractHeaderText(h).Equals(headerText, StringComparison.OrdinalIgnoreCase));

            if (exactMatch)
            {
                properMarkdownHeaders++;
                Console.WriteLine($"[DEBUG] ✓ Perfect match: '{headerText}' at level {headerLevel}");
            }
            else if (headerWithDifferentLevel)
            {
                properMarkdownHeaders++;
                Console.WriteLine($"[DEBUG] ○ Header preserved but level changed: '{headerText}'");
            }
            else
            {
                Console.WriteLine($"[DEBUG] ✗ Lost as Markdown header: '{headerText}'");
            }
        }

        // 厳格な基準: 80%以上のヘッダーがMarkdownヘッダーとして保持されていること
        bool passed = (double)properMarkdownHeaders / totalHeaders >= 0.8;
        string details = $"{properMarkdownHeaders}/{totalHeaders} headers preserved as Markdown headers";
        
        return (passed, details);
    }

    private static (bool Passed, string Details) EvaluateSectionStructureAsTest(MarkdownStructureAnalysis analysis)
    {
        if (analysis.OriginalSections.Count <= 1) 
            return (true, "Single section document");

        int matchedSections = 0;
        int totalSections = analysis.OriginalSections.Count;

        foreach (var originalSection in analysis.OriginalSections)
        {
            var originalBlocks = ExtractContentBlocks(originalSection);
            bool sectionMatched = false;

            foreach (var originalBlock in originalBlocks)
            {
                if (string.IsNullOrWhiteSpace(originalBlock)) continue;

                foreach (var convertedSection in analysis.ConvertedSections)
                {
                    var convertedBlocks = ExtractContentBlocks(convertedSection);
                    
                    if (convertedBlocks.Any(cb => CalculateBlockSimilarity(originalBlock, cb) > 0.7)) // 厳格な閾値
                    {
                        sectionMatched = true;
                        break;
                    }
                }
                
                if (sectionMatched) break;
            }

            if (sectionMatched) matchedSections++;
        }

        // 厳格な基準: 80%以上のセクションが適切に保持されていること
        bool passed = (double)matchedSections / totalSections >= 0.8;
        string details = $"{matchedSections}/{totalSections} sections preserved with high fidelity";
        
        return (passed, details);
    }

    private static (bool Passed, string Details) EvaluateTableStructureAsTest(MarkdownStructureAnalysis analysis)
    {
        if (analysis.OriginalTables.Count == 0) 
            return (true, "No tables to validate");

        int properMarkdownTables = 0;
        foreach (var originalTable in analysis.OriginalTables)
        {
            var originalTableContent = ExtractTableContentWords(originalTable);
            
            // 厳格な基準: Markdownテーブルとして適切に保持されているか
            bool preservedAsMarkdownTable = analysis.ConvertedTables.Any(convertedTable =>
                CalculateTableContentSimilarity(originalTableContent, ExtractTableContentWords(convertedTable)) > 0.8); // 厳格な閾値

            if (preservedAsMarkdownTable)
                properMarkdownTables++;
        }

        // 厳格な基準: 80%以上のテーブルがMarkdownテーブルとして保持されていること
        bool passed = (double)properMarkdownTables / analysis.OriginalTables.Count >= 0.8;
        string details = $"{properMarkdownTables}/{analysis.OriginalTables.Count} tables preserved as Markdown tables";
        
        return (passed, details);
    }

    private static (bool Passed, string Details) EvaluateListStructureAsTest(MarkdownStructureAnalysis analysis)
    {
        if (analysis.OriginalLists.Count == 0) 
            return (true, "No lists to validate");

        int properMarkdownLists = 0;
        foreach (var originalList in analysis.OriginalLists)
        {
            var originalListItems = ExtractListItems(originalList);
            
            // 厳格な基準: Markdownリストとして適切に保持されているか
            bool preservedAsMarkdownList = analysis.ConvertedLists.Any(convertedList =>
                CalculateListSimilarity(originalListItems, ExtractListItems(convertedList)) > 0.8); // 厳格な閾値

            if (preservedAsMarkdownList)
                properMarkdownLists++;
        }

        // 厳格な基準: 80%以上のリストがMarkdownリストとして保持されていること
        bool passed = (double)properMarkdownLists / analysis.OriginalLists.Count >= 0.8;
        string details = $"{properMarkdownLists}/{analysis.OriginalLists.Count} lists preserved as Markdown lists";
        
        return (passed, details);
    }

    private (bool Passed, string Details) EvaluateMarkdownFormattingAsTest(MarkdownStructureAnalysis analysis)
    {
        string originalContent = string.Join("\n", analysis.OriginalSections);
        string convertedContent = string.Join("\n", analysis.ConvertedSections);

        int formattingTests = 0;
        int preservedFormatting = 0;

        // 太字の保持確認 - 厳格な基準
        var boldMatches = Regex.Matches(originalContent, @"\*\*([^*]+)\*\*");
        if (boldMatches.Count > 0)
        {
            formattingTests++;
            // 厳格な基準: 太字がMarkdown記法として保持されているか、または視覚的に強調されているか
            bool boldPreserved = convertedContent.Contains("**") || // Markdown形式で保持
                                boldMatches.Cast<Match>()
                                .All(match => convertedContent.Contains(match.Groups[1].Value, StringComparison.OrdinalIgnoreCase));
            
            if (boldPreserved) preservedFormatting++;
        }

        // 斜体の保持確認 - 厳格な基準
        var italicMatches = Regex.Matches(originalContent, @"(?<!\*)\*([^*]+)\*(?!\*)");
        if (italicMatches.Count > 0)
        {
            formattingTests++;
            bool italicPreserved = convertedContent.Contains("*") || // Markdown形式で保持
                                 italicMatches.Cast<Match>()
                                 .All(match => convertedContent.Contains(match.Groups[1].Value, StringComparison.OrdinalIgnoreCase));
            
            if (italicPreserved) preservedFormatting++;
        }

        // コードの保持確認 - 厳格な基準
        var codeMatches = Regex.Matches(originalContent, @"`([^`]+)`");
        if (codeMatches.Count > 0)
        {
            formattingTests++;
            bool codePreserved = convertedContent.Contains("`") || // Markdown形式で保持
                               codeMatches.Cast<Match>()
                               .All(match => convertedContent.Contains(match.Groups[1].Value, StringComparison.OrdinalIgnoreCase));
            
            if (codePreserved) preservedFormatting++;
        }

        if (formattingTests == 0)
            return (true, "No formatting to validate");

        // 厳格な基準: すべての書式タイプが保持されていること
        bool passed = preservedFormatting == formattingTests;
        string details = $"{preservedFormatting}/{formattingTests} formatting types fully preserved";
        
        return (passed, details);
    }

    private static (bool Passed, string Details) EvaluateEmphasisAsTest(MarkdownStructureAnalysis analysis)
    {
        string originalContent = string.Join("\n", analysis.OriginalSections);
        string convertedContent = string.Join("\n", analysis.ConvertedSections);

        int emphasisTests = 0;
        int preservedEmphasis = 0;

        // 太字の保持確認（**と__）
        var boldDoubleAsterisk = Regex.Matches(originalContent, @"\*\*([^*]+)\*\*");
        var boldDoubleUnderscore = Regex.Matches(originalContent, @"__([^_]+)__");
        
        if (boldDoubleAsterisk.Count > 0 || boldDoubleUnderscore.Count > 0)
        {
            emphasisTests++;
            bool boldPreserved = convertedContent.Contains("**") || convertedContent.Contains("__");
            if (boldPreserved) preservedEmphasis++;
        }

        // 斜体の保持確認（*と_）
        var italicAsterisk = Regex.Matches(originalContent, @"(?<!\*)\*([^*]+)\*(?!\*)");
        var italicUnderscore = Regex.Matches(originalContent, @"(?<!_)_([^_]+)_(?!_)");
        
        if (italicAsterisk.Count > 0 || italicUnderscore.Count > 0)
        {
            emphasisTests++;
            bool italicPreserved = convertedContent.Contains("*") || convertedContent.Contains("_");
            if (italicPreserved) preservedEmphasis++;
        }

        // 太字斜体の保持確認（***と___）
        var boldItalicTripleAsterisk = Regex.Matches(originalContent, @"\*\*\*([^*]+)\*\*\*");
        var boldItalicTripleUnderscore = Regex.Matches(originalContent, @"___([^_]+)___");
        
        if (boldItalicTripleAsterisk.Count > 0 || boldItalicTripleUnderscore.Count > 0)
        {
            emphasisTests++;
            bool boldItalicPreserved = convertedContent.Contains("***") || convertedContent.Contains("___");
            if (boldItalicPreserved) preservedEmphasis++;
        }

        if (emphasisTests == 0)
            return (true, "No emphasis to validate");

        bool passed = preservedEmphasis == emphasisTests;
        string details = $"{preservedEmphasis}/{emphasisTests} emphasis types preserved";
        
        return (passed, details);
    }

    private static (bool Passed, string Details) EvaluateLinksAsTest(MarkdownStructureAnalysis analysis)
    {
        string originalContent = string.Join("\n", analysis.OriginalSections);
        string convertedContent = string.Join("\n", analysis.ConvertedSections);

        int linkTests = 0;
        int preservedLinks = 0;

        // インラインリンクの保持確認
        var inlineLinks = Regex.Matches(originalContent, @"\[([^\]]+)\]\(([^)]+)\)");
        if (inlineLinks.Count > 0)
        {
            linkTests++;
            bool inlineLinksPreserved = convertedContent.Contains("[") && convertedContent.Contains("](");
            if (inlineLinksPreserved) preservedLinks++;
        }

        // 自動リンクの保持確認
        var autoLinks = Regex.Matches(originalContent, @"<(https?://[^>]+)>");
        if (autoLinks.Count > 0)
        {
            linkTests++;
            bool autoLinksPreserved = convertedContent.Contains("<http") || 
                                    autoLinks.Cast<Match>()
                                    .Any(match => convertedContent.Contains(match.Groups[1].Value));
            if (autoLinksPreserved) preservedLinks++;
        }

        if (linkTests == 0)
            return (true, "No links to validate");

        bool passed = preservedLinks == linkTests;
        string details = $"{preservedLinks}/{linkTests} link types preserved";
        
        return (passed, details);
    }

    private static (bool Passed, string Details) EvaluateCodeAsTest(MarkdownStructureAnalysis analysis)
    {
        string originalContent = string.Join("\n", analysis.OriginalSections);
        string convertedContent = string.Join("\n", analysis.ConvertedSections);

        int codeTests = 0;
        int preservedCode = 0;

        // インラインコードの保持確認
        var inlineCode = Regex.Matches(originalContent, @"`([^`]+)`");
        if (inlineCode.Count > 0)
        {
            codeTests++;
            bool inlineCodePreserved = convertedContent.Contains("`");
            if (inlineCodePreserved) preservedCode++;
        }

        // コードブロックの保持確認
        var codeBlocks = Regex.Matches(originalContent, @"```[^`]*```", RegexOptions.Singleline);
        if (codeBlocks.Count > 0)
        {
            codeTests++;
            bool codeBlocksPreserved = convertedContent.Contains("```");
            if (codeBlocksPreserved) preservedCode++;
        }

        if (codeTests == 0)
            return (true, "No code to validate");

        bool passed = preservedCode == codeTests;
        string details = $"{preservedCode}/{codeTests} code types preserved";
        
        return (passed, details);
    }

    private static (bool Passed, string Details) EvaluateQuotesAsTest(MarkdownStructureAnalysis analysis)
    {
        string originalContent = string.Join("\n", analysis.OriginalSections);
        string convertedContent = string.Join("\n", analysis.ConvertedSections);

        int quoteTests = 0;
        int preservedQuotes = 0;

        // 引用の保持確認
        var quotes = Regex.Matches(originalContent, @"^>.*$", RegexOptions.Multiline);
        if (quotes.Count > 0)
        {
            quoteTests++;
            bool quotesPreserved = convertedContent.Contains(">");
            if (quotesPreserved) preservedQuotes++;
        }

        if (quoteTests == 0)
            return (true, "No quotes to validate");

        bool passed = preservedQuotes == quoteTests;
        string details = $"{preservedQuotes}/{quoteTests} quote types preserved";
        
        return (passed, details);
    }

    private static (bool Passed, string Details) EvaluateOtherSyntaxAsTest(MarkdownStructureAnalysis analysis)
    {
        string originalContent = string.Join("\n", analysis.OriginalSections);
        string convertedContent = string.Join("\n", analysis.ConvertedSections);

        int otherTests = 0;
        int preservedOther = 0;

        // 水平線の保持確認
        var horizontalRules = Regex.Matches(originalContent, @"^(---+|\*\*\*+|___+)$", RegexOptions.Multiline);
        if (horizontalRules.Count > 0)
        {
            otherTests++;
            bool horizontalRulesPreserved = convertedContent.Contains("---") || convertedContent.Contains("***") || convertedContent.Contains("___");
            if (horizontalRulesPreserved) preservedOther++;
        }

        if (otherTests == 0)
            return (true, "No other syntax to validate");

        bool passed = preservedOther == otherTests;
        string details = $"{preservedOther}/{otherTests} other syntax types preserved";
        
        return (passed, details);
    }






    private static List<string> ExtractHeaders(string content)
    {
        return content.Split('\n')
            .Where(line => line.TrimStart().StartsWith('#'))
            .Select(line => line.Trim())
            .ToList();
    }

    private static List<string> ExtractSections(string content)
    {
        var lines = content.Split('\n');
        var sections = new List<string>();
        var currentSection = new List<string>();

        foreach (var line in lines)
        {
            if (line.TrimStart().StartsWith('#') && currentSection.Count > 0)
            {
                sections.Add(string.Join("\n", currentSection));
                currentSection.Clear();
            }
            currentSection.Add(line);
        }

        if (currentSection.Count > 0)
        {
            sections.Add(string.Join("\n", currentSection));
        }

        return sections;
    }

    private static List<string> ExtractTables(string content)
    {
        var lines = content.Split('\n');
        var tables = new List<string>();
        var currentTable = new List<string>();
        bool inTable = false;

        foreach (var line in lines)
        {
            if (line.Contains('|'))
            {
                inTable = true;
                currentTable.Add(line);
            }
            else if (inTable && string.IsNullOrWhiteSpace(line))
            {
                if (currentTable.Count > 0)
                {
                    tables.Add(string.Join("\n", currentTable));
                    currentTable.Clear();
                }
                inTable = false;
            }
            else if (inTable)
            {
                currentTable.Add(line);
            }
        }

        if (currentTable.Count > 0)
        {
            tables.Add(string.Join("\n", currentTable));
        }

        return tables;
    }

    private static List<string> ExtractLists(string content)
    {
        var lines = content.Split('\n');
        var lists = new List<string>();
        var currentList = new List<string>();
        bool inList = false;

        foreach (var line in lines)
        {
            var trimmed = line.TrimStart();
            if (trimmed.StartsWith('-') || trimmed.StartsWith('*') || trimmed.StartsWith('+') ||
                (trimmed.Length > 2 && char.IsDigit(trimmed[0]) && trimmed[1] == '.'))
            {
                inList = true;
                currentList.Add(line);
            }
            else if (inList && string.IsNullOrWhiteSpace(line))
            {
                if (currentList.Count > 0)
                {
                    lists.Add(string.Join("\n", currentList));
                    currentList.Clear();
                }
                inList = false;
            }
            else if (inList && (trimmed.StartsWith(' ') || trimmed.StartsWith('\t')))
            {
                currentList.Add(line);
            }
            else if (inList)
            {
                if (currentList.Count > 0)
                {
                    lists.Add(string.Join("\n", currentList));
                    currentList.Clear();
                }
                inList = false;
            }
        }

        if (currentList.Count > 0)
        {
            lists.Add(string.Join("\n", currentList));
        }

        return lists;
    }

    private static string ExtractHeaderText(string header)
    {
        return header.Replace("#", "").Trim();
    }

    private static int GetHeaderLevel(string header)
    {
        return header.TakeWhile(c => c == '#').Count();
    }

    private static double CalculateTextSimilarity(string text1, string text2)
    {
        if (string.IsNullOrEmpty(text1) || string.IsNullOrEmpty(text2))
            return 0.0;

        var words1 = text1.Split(new char[] { ' ', '\n', '\t' }, StringSplitOptions.RemoveEmptyEntries)
                         .Select(w => w.ToLowerInvariant()).ToHashSet();
        var words2 = text2.Split(new char[] { ' ', '\n', '\t' }, StringSplitOptions.RemoveEmptyEntries)
                         .Select(w => w.ToLowerInvariant()).ToHashSet();

        var intersection = words1.Intersect(words2).Count();
        var union = words1.Union(words2).Count();

        return union == 0 ? 0.0 : (double)intersection / union;
    }

    private static bool HasTabularData(string content)
    {
        // テーブル的な構造を示すパターンを検索
        var lines = content.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        
        // 複数の列データが存在するかチェック
        int tabularLines = 0;
        foreach (var line in lines)
        {
            // 複数の単語やデータが空白で区切られている行をカウント
            var words = line.Trim().Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            if (words.Length >= 3) // 3列以上のデータがある行
            {
                tabularLines++;
            }
        }
        
        // 3行以上の表形式データがあればテーブルと判定
        return tabularLines >= 3;
    }

    private static List<string> ExtractTableContentWords(string tableContent)
    {
        var words = new List<string>();
        var lines = tableContent.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        
        foreach (var line in lines)
        {
            // テーブル区切り文字を除去してワードを抽出
            var cleanLine = line.Replace("|", " ").Replace("-", " ").Replace(":", " ");
            var lineWords = cleanLine.Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries)
                                   .Where(w => w.Length > 1) // 1文字の単語は除外
                                   .Select(w => w.Trim())
                                   .Where(w => !string.IsNullOrEmpty(w));
            words.AddRange(lineWords);
        }
        
        return words.Distinct().ToList();
    }

    private static double CalculateTableContentSimilarity(List<string> content1, List<string> content2)
    {
        if (content1.Count == 0 || content2.Count == 0)
            return 0.0;

        var set1 = content1.Select(w => w.ToLowerInvariant()).ToHashSet();
        var set2 = content2.Select(w => w.ToLowerInvariant()).ToHashSet();

        var intersection = set1.Intersect(set2).Count();
        var union = set1.Union(set2).Count();

        return union == 0 ? 0.0 : (double)intersection / union;
    }

    private static List<string> ExtractListItems(string listContent)
    {
        var items = new List<string>();
        var lines = listContent.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        
        foreach (var line in lines)
        {
            var trimmed = line.TrimStart();
            if (trimmed.StartsWith('-') || trimmed.StartsWith('*') || trimmed.StartsWith('+') ||
                (trimmed.Length > 2 && char.IsDigit(trimmed[0]) && trimmed[1] == '.'))
            {
                // リストマーカーを除去してアイテムテキストを抽出
                var itemText = trimmed[(trimmed.IndexOfAny([' ', '\t']) + 1)..].Trim();
                if (!string.IsNullOrEmpty(itemText))
                {
                    items.Add(itemText);
                }
            }
        }
        
        return items;
    }

    private static List<string> ExtractListItemsFromText(string text)
    {
        var items = new List<string>();
        var lines = text.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        
        foreach (var line in lines)
        {
            var trimmed = line.Trim();
            // 箇条書き的なパターンを検出
            if (trimmed.StartsWith('•') || trimmed.StartsWith('◦') || trimmed.StartsWith('▪'))
            {
                var itemText = trimmed.Substring(1).Trim();
                if (!string.IsNullOrEmpty(itemText))
                {
                    items.Add(itemText);
                }
            }
        }
        
        return items;
    }

    private static bool HasListLikeStructure(string text)
    {
        var lines = text.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        int listLikeLines = 0;
        
        foreach (var line in lines)
        {
            var trimmed = line.Trim();
            if (trimmed.StartsWith('•') || trimmed.StartsWith('◦') || trimmed.StartsWith('▪') ||
                trimmed.StartsWith('-') || trimmed.StartsWith('*') || trimmed.StartsWith('+'))
            {
                listLikeLines++;
            }
        }
        
        return listLikeLines >= 2; // 2行以上の箇条書きがあればリスト的構造と判定
    }

    private static double CalculateListSimilarity(List<string> items1, List<string> items2)
    {
        if (items1.Count == 0 || items2.Count == 0)
            return 0.0;

        int matchedItems = 0;
        foreach (var item1 in items1)
        {
            if (items2.Any(item2 => item2.Contains(item1, StringComparison.OrdinalIgnoreCase) ||
                                  item1.Contains(item2, StringComparison.OrdinalIgnoreCase)))
            {
                matchedItems++;
            }
        }

        return (double)matchedItems / Math.Max(items1.Count, items2.Count);
    }

    private static List<string> ExtractContentBlocks(string content)
    {
        // 空行で分割してブロック単位に分ける
        var blocks = content.Split(new string[] { "\n\n", "\r\n\r\n" }, StringSplitOptions.RemoveEmptyEntries)
                           .Select(block => block.Trim())
                           .Where(block => !string.IsNullOrWhiteSpace(block))
                           .ToList();
        
        // 単一行で構成されている場合は改行で分割
        if (blocks.Count == 1)
        {
            blocks = content.Split('\n')
                           .Select(line => line.Trim())
                           .Where(line => !string.IsNullOrWhiteSpace(line))
                           .ToList();
        }
        
        return blocks;
    }

    private static double CalculateBlockSimilarity(string block1, string block2)
    {
        if (string.IsNullOrWhiteSpace(block1) || string.IsNullOrWhiteSpace(block2))
            return 0.0;

        // Markdownの記号を除去して比較
        var cleanBlock1 = CleanMarkdownForComparison(block1);
        var cleanBlock2 = CleanMarkdownForComparison(block2);

        // 完全一致または部分一致をチェック
        if (cleanBlock1.Equals(cleanBlock2, StringComparison.OrdinalIgnoreCase))
            return 1.0;

        if (cleanBlock1.Contains(cleanBlock2, StringComparison.OrdinalIgnoreCase) ||
            cleanBlock2.Contains(cleanBlock1, StringComparison.OrdinalIgnoreCase))
            return 0.8;

        // 単語レベルでの一致度を計算
        return CalculateTextSimilarity(cleanBlock1, cleanBlock2);
    }

    private static string CleanMarkdownForComparison(string text)
    {
        // Markdownの書式設定記号を除去
        return text.Replace("#", "")
                  .Replace("*", "")
                  .Replace("_", "")
                  .Replace("`", "")
                  .Replace("|", "")
                  .Replace("-", "")
                  .Replace(">", "")
                  .Replace("•", "")
                  .Trim();
    }

}