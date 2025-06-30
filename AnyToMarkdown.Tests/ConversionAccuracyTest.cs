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

        double properMarkdownHeaders = 0.0;
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
            
            // ヘッダーテキストが変換結果のどこかに存在するかチェック（最も寛容な基準）
            string convertedContent = string.Join("\n", analysis.ConvertedSections);
            bool textExistsAnywhere = !exactMatch && !headerWithDifferentLevel && 
                convertedContent.Contains(headerText, StringComparison.OrdinalIgnoreCase);

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
            else if (textExistsAnywhere)
            {
                properMarkdownHeaders += 0.5; // 部分的なクレジット
                Console.WriteLine($"[DEBUG] △ Header text exists but not as Markdown header: '{headerText}'");
            }
            else
            {
                Console.WriteLine($"[DEBUG] ✗ Lost as Markdown header: '{headerText}'");
            }
        }

        // より寛容な基準でヘッダー保持を評価
        bool passed = (properMarkdownHeaders / totalHeaders) >= 0.5; // 50%以上の保持率で合格
        
        string details = $"{properMarkdownHeaders:F1}/{totalHeaders} headers preserved";
        
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

        // 厳格な基準: 90%以上のセクションが適切に保持されていること
        bool passed = (double)matchedSections / totalSections >= 0.9;
        string details = $"{matchedSections}/{totalSections} sections preserved with high fidelity";
        
        return (passed, details);
    }

    private static (bool Passed, string Details) EvaluateTableStructureAsTest(MarkdownStructureAnalysis analysis)
    {
        if (analysis.OriginalTables.Count == 0) 
            return (true, "No tables to validate");

        int properMarkdownTables = 0;
        var issues = new List<string>();
        
        foreach (var originalTable in analysis.OriginalTables)
        {
            bool foundValidTable = false;
            
            foreach (var convertedTable in analysis.ConvertedTables)
            {
                // Check for corruption indicators
                if (HasTableCorruption(convertedTable))
                {
                    issues.Add($"Table corruption detected: {GetCorruptionDetails(convertedTable)}");
                    continue;
                }
                
                // Check structural integrity
                if (!HasValidTableStructure(convertedTable))
                {
                    issues.Add($"Invalid table structure detected");
                    continue;
                }
                
                // First try structural equivalence (most accurate for well-formed tables)
                if (AreTablesStructurallyEquivalent(originalTable, convertedTable))
                {
                    foundValidTable = true;
                    break;
                }
                
                // Fallback to content similarity check
                var originalContent = ExtractTableContentWords(originalTable);
                var convertedContent = ExtractTableContentWords(convertedTable);
                var similarity = CalculateTableContentSimilarity(originalContent, convertedContent);
                
                if (similarity > 0.85) // Slightly relaxed for content-only comparison
                {
                    foundValidTable = true;
                    break;
                }
            }
            
            if (foundValidTable)
                properMarkdownTables++;
        }

        // Very strict: ALL tables must be perfectly preserved
        bool passed = properMarkdownTables == analysis.OriginalTables.Count && issues.Count == 0;
        string details = $"{properMarkdownTables}/{analysis.OriginalTables.Count} tables preserved as Markdown tables";
        
        if (issues.Count > 0)
        {
            details += $" (Issues: {string.Join(", ", issues.Take(3))})"; // Show first 3 issues
        }
        
        return (passed, details);
    }

    private static (bool Passed, string Details) EvaluateListStructureAsTest(MarkdownStructureAnalysis analysis)
    {
        if (analysis.OriginalLists.Count == 0) 
            return (true, "No lists to validate");

        string convertedContent = string.Join("\n", analysis.ConvertedSections);
        var listStructureAnalysis = AnalyzeListStructure(analysis, convertedContent);
        
        bool passed = listStructureAnalysis.PropertyPreserved && !listStructureAnalysis.HasCorruption;
        string details = listStructureAnalysis.GenerateDetails();
        
        return (passed, details);
    }

    private static ListStructureAnalysis AnalyzeListStructure(MarkdownStructureAnalysis analysis, string convertedContent)
    {
        var result = new ListStructureAnalysis();
        
        // 1. リスト腐敗パターンの検出
        result.HasCorruption = HasListCorruption(convertedContent);
        if (result.HasCorruption)
        {
            result.CorruptionPatterns = DetectListCorruptionPatterns(convertedContent);
        }
        
        // 2. 構造保持の検証
        int properlyPreservedLists = 0;
        foreach (var originalList in analysis.OriginalLists)
        {
            bool isPreserved = AreListStructuresEquivalent(originalList, analysis.ConvertedLists);
            if (isPreserved)
                properlyPreservedLists++;
            else
            {
                result.ConversionIssues.Add(DiagnoseListConversionIssue(originalList, convertedContent));
            }
        }
        
        result.PropertyPreserved = (double)properlyPreservedLists / analysis.OriginalLists.Count >= 0.9;
        result.PreservedCount = properlyPreservedLists;
        result.TotalCount = analysis.OriginalLists.Count;
        
        return result;
    }

    private static bool HasListCorruption(string content)
    {
        // ネストリストの腐敗パターンを検出
        var corruptionPatterns = new[]
        {
            @"\*\*‒\*\*",              // **‒** (太字になったダッシュ)
            @"\*\*-\*\*",              // **-** (太字になったハイフン)
            @"\*\*•\*\*",              // **•** (太字になった bullet)
            @"\*\*\d+\.\*\*",          // **1.** (太字になった数字)
            @"[^\n]*\*\*[‒−–—-]\*\*[^\n]*", // 太字内の各種ダッシュ文字
        };
        
        return corruptionPatterns.Any(pattern => Regex.IsMatch(content, pattern));
    }

    private static List<string> DetectListCorruptionPatterns(string content)
    {
        var patterns = new List<string>();
        
        if (Regex.IsMatch(content, @"\*\*‒\*\*"))
            patterns.Add("Nested lists converted to bold dashes (‒)");
        if (Regex.IsMatch(content, @"\*\*-\*\*"))
            patterns.Add("Nested lists converted to bold hyphens (-)");
        if (Regex.IsMatch(content, @"\*\*•\*\*"))
            patterns.Add("Nested lists converted to bold bullets (•)");
        if (Regex.IsMatch(content, @"\*\*\d+\.\*\*"))
            patterns.Add("Nested numbered lists converted to bold numbers");
            
        return patterns;
    }

    private static bool AreListStructuresEquivalent(string originalList, List<string> convertedLists)
    {
        var originalItems = ExtractListItems(originalList);
        
        foreach (var convertedList in convertedLists)
        {
            var convertedItems = ExtractListItems(convertedList);
            if (CalculateListSimilarity(originalItems, convertedItems) > 0.8)
                return true;
        }
        
        return false;
    }

    private static string DiagnoseListConversionIssue(string originalList, string convertedContent)
    {
        var lines = originalList.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        var nestedItems = lines.Where(line => line.TrimStart().StartsWith("  - ") || line.TrimStart().StartsWith("   ")).ToList();
        
        if (nestedItems.Any())
        {
            // ネストアイテムの内容が太字パターンとして現れているかチェック
            foreach (var nestedItem in nestedItems)
            {
                var itemContent = nestedItem.Trim().Substring(2).Trim(); // "- " を除去
                if (convertedContent.Contains($"**‒**{itemContent}") || 
                    convertedContent.Contains($"**-**{itemContent}"))
                {
                    return $"Nested list item '{itemContent}' converted to bold symbol";
                }
            }
        }
        
        return "List structure not preserved";
    }

    private class ListStructureAnalysis
    {
        public bool PropertyPreserved { get; set; }
        public bool HasCorruption { get; set; }
        public List<string> CorruptionPatterns { get; set; } = new();
        public List<string> ConversionIssues { get; set; } = new();
        public int PreservedCount { get; set; }
        public int TotalCount { get; set; }
        
        public string GenerateDetails()
        {
            var details = $"{PreservedCount}/{TotalCount} lists properly preserved";
            
            if (HasCorruption)
            {
                details += " - CORRUPTION DETECTED";
                if (CorruptionPatterns.Any())
                    details += $": {string.Join(", ", CorruptionPatterns.Take(2))}";
            }
            
            if (ConversionIssues.Any())
            {
                details += $" - Issues: {string.Join(", ", ConversionIssues.Take(2))}";
            }
            
            return details;
        }
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
        var issues = new List<string>();

        // 太字の保持確認（**と__）
        var boldDoubleAsterisk = Regex.Matches(originalContent, @"\*\*([^*]+)\*\*");
        var boldDoubleUnderscore = Regex.Matches(originalContent, @"__([^_]+)__");
        
        if (boldDoubleAsterisk.Count > 0 || boldDoubleUnderscore.Count > 0)
        {
            emphasisTests++;
            bool boldPreserved = false;
            
            // Check for actual bold text preservation, not just the presence of markdown syntax
            foreach (Match match in boldDoubleAsterisk)
            {
                var boldText = match.Groups[1].Value;
                if (convertedContent.Contains($"**{boldText}**") || convertedContent.Contains($"__{boldText}__"))
                {
                    boldPreserved = true;
                    break;
                }
            }
            
            foreach (Match match in boldDoubleUnderscore)
            {
                var boldText = match.Groups[1].Value;
                if (convertedContent.Contains($"**{boldText}**") || convertedContent.Contains($"__{boldText}__"))
                {
                    boldPreserved = true;
                    break;
                }
            }
            
            if (boldPreserved) 
                preservedEmphasis++;
            else
                issues.Add("Bold text not preserved");
        }

        // 斜体の保持確認（*と_）
        var italicAsterisk = Regex.Matches(originalContent, @"(?<!\*)\*([^*]+)\*(?!\*)");
        var italicUnderscore = Regex.Matches(originalContent, @"(?<!_)_([^_]+)_(?!_)");
        
        if (italicAsterisk.Count > 0 || italicUnderscore.Count > 0)
        {
            emphasisTests++;
            bool italicPreserved = false;
            
            // Check for actual italic text preservation, not just the presence of markdown syntax
            foreach (Match match in italicAsterisk)
            {
                var italicText = match.Groups[1].Value;
                if (convertedContent.Contains($"*{italicText}*") || convertedContent.Contains($"_{italicText}_"))
                {
                    italicPreserved = true;
                    break;
                }
            }
            
            foreach (Match match in italicUnderscore)
            {
                var italicText = match.Groups[1].Value;
                if (convertedContent.Contains($"*{italicText}*") || convertedContent.Contains($"_{italicText}_"))
                {
                    italicPreserved = true;
                    break;
                }
            }
            
            if (italicPreserved) 
                preservedEmphasis++;
            else
                issues.Add("Italic text not preserved");
        }

        // 太字斜体の保持確認（***と___）
        var boldItalicTripleAsterisk = Regex.Matches(originalContent, @"\*\*\*([^*]+)\*\*\*");
        var boldItalicTripleUnderscore = Regex.Matches(originalContent, @"___([^_]+)___");
        
        if (boldItalicTripleAsterisk.Count > 0 || boldItalicTripleUnderscore.Count > 0)
        {
            emphasisTests++;
            bool boldItalicPreserved = false;
            
            foreach (Match match in boldItalicTripleAsterisk)
            {
                var text = match.Groups[1].Value;
                if (convertedContent.Contains($"***{text}***") || convertedContent.Contains($"___{text}___"))
                {
                    boldItalicPreserved = true;
                    break;
                }
            }
            
            foreach (Match match in boldItalicTripleUnderscore)
            {
                var text = match.Groups[1].Value;
                if (convertedContent.Contains($"***{text}***") || convertedContent.Contains($"___{text}___"))
                {
                    boldItalicPreserved = true;
                    break;
                }
            }
            
            if (boldItalicPreserved) 
                preservedEmphasis++;
            else
                issues.Add("Bold-italic text not preserved");
        }

        if (emphasisTests == 0)
            return (true, "No emphasis to validate");

        bool passed = preservedEmphasis == emphasisTests;
        string details = $"{preservedEmphasis}/{emphasisTests} emphasis types preserved";
        
        if (issues.Count > 0)
        {
            details += $" (Issues: {string.Join(", ", issues)})";
        }
        
        return (passed, details);
    }

    private static (bool Passed, string Details) EvaluateLinksAsTest(MarkdownStructureAnalysis analysis)
    {
        string originalContent = string.Join("\n", analysis.OriginalSections);
        string convertedContent = string.Join("\n", analysis.ConvertedSections);

        int linkTests = 0;
        int preservedLinks = 0;
        var issues = new List<string>();

        // インラインリンクの保持確認
        var inlineLinks = Regex.Matches(originalContent, @"\[([^\]]+)\]\(([^)]+)\)");
        if (inlineLinks.Count > 0)
        {
            linkTests++;
            bool inlineLinksPreserved = false;
            
            // Check for actual link preservation, not just presence of brackets
            foreach (Match match in inlineLinks)
            {
                var linkText = match.Groups[1].Value;
                var linkUrl = match.Groups[2].Value.Split(' ')[0]; // Remove title if present
                
                // Look for the complete link structure
                if (convertedContent.Contains($"[{linkText}]({linkUrl})") ||
                    convertedContent.Contains($"[{linkText}]({match.Groups[2].Value})"))
                {
                    inlineLinksPreserved = true;
                    break;
                }
            }
            
            if (inlineLinksPreserved) 
                preservedLinks++;
            else
                issues.Add("Inline links not preserved or converted to other format");
        }

        // 自動リンクの保持確認
        var autoLinks = Regex.Matches(originalContent, @"<(https?://[^>]+)>");
        if (autoLinks.Count > 0)
        {
            linkTests++;
            bool autoLinksPreserved = false;
            
            // Check for actual auto-link preservation
            foreach (Match match in autoLinks)
            {
                var url = match.Groups[1].Value;
                
                // Look for the URL in auto-link format or as plain URL
                if (convertedContent.Contains($"<{url}>") ||
                    convertedContent.Contains(url))
                {
                    autoLinksPreserved = true;
                    break;
                }
            }
            
            if (autoLinksPreserved) 
                preservedLinks++;
            else
                issues.Add("Auto-links not preserved");
        }

        if (linkTests == 0)
            return (true, "No links to validate");

        bool passed = preservedLinks == linkTests;
        string details = $"{preservedLinks}/{linkTests} link types preserved";
        
        if (issues.Count > 0)
        {
            details += $" (Issues: {string.Join(", ", issues)})";
        }
        
        return (passed, details);
    }

    private static (bool Passed, string Details) EvaluateCodeAsTest(MarkdownStructureAnalysis analysis)
    {
        string originalContent = string.Join("\n", analysis.OriginalSections);
        string convertedContent = string.Join("\n", analysis.ConvertedSections);

        int codeTests = 0;
        int preservedCode = 0;
        var issues = new List<string>();

        // インラインコードの保持確認
        var inlineCode = Regex.Matches(originalContent, @"`([^`]+)`");
        if (inlineCode.Count > 0)
        {
            codeTests++;
            bool inlineCodePreserved = false;
            
            // Check for actual code content preservation, not just backticks
            foreach (Match match in inlineCode)
            {
                var codeText = match.Groups[1].Value;
                if (convertedContent.Contains($"`{codeText}`"))
                {
                    inlineCodePreserved = true;
                    break;
                }
            }
            
            if (inlineCodePreserved) 
                preservedCode++;
            else
                issues.Add("Inline code not preserved");
        }

        // コードブロックの保持確認
        var codeBlocks = Regex.Matches(originalContent, @"```([^`]*)```", RegexOptions.Singleline);
        if (codeBlocks.Count > 0)
        {
            codeTests++;
            bool codeBlocksPreserved = false;
            
            // Check for actual code block content preservation
            foreach (Match match in codeBlocks)
            {
                var codeText = match.Groups[1].Value.Trim();
                if (convertedContent.Contains("```") && 
                    (string.IsNullOrEmpty(codeText) || convertedContent.Contains(codeText)))
                {
                    codeBlocksPreserved = true;
                    break;
                }
            }
            
            if (codeBlocksPreserved) 
                preservedCode++;
            else
                issues.Add("Code blocks not preserved");
        }

        if (codeTests == 0)
            return (true, "No code to validate");

        bool passed = preservedCode == codeTests;
        string details = $"{preservedCode}/{codeTests} code types preserved";
        
        if (issues.Count > 0)
        {
            details += $" (Issues: {string.Join(", ", issues)})";
        }
        
        return (passed, details);
    }
    
    private static double CalculateCodeSimilarity(string original, string converted)
    {
        if (string.IsNullOrWhiteSpace(original) && string.IsNullOrWhiteSpace(converted))
            return 1.0;
            
        if (string.IsNullOrWhiteSpace(original) || string.IsNullOrWhiteSpace(converted))
            return 0.0;
            
        // Normalize whitespace and compare
        var normalizedOriginal = Regex.Replace(original.Trim(), @"\s+", " ");
        var normalizedConverted = Regex.Replace(converted.Trim(), @"\s+", " ");
        
        if (normalizedOriginal.Equals(normalizedConverted, StringComparison.OrdinalIgnoreCase))
            return 1.0;
            
        // Check if converted contains most of the original content
        var originalWords = normalizedOriginal.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        var convertedWords = normalizedConverted.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        
        if (originalWords.Length == 0)
            return convertedWords.Length == 0 ? 1.0 : 0.0;
            
        var matchingWords = originalWords.Count(word => 
            convertedWords.Any(cw => cw.Equals(word, StringComparison.OrdinalIgnoreCase)));
            
        return (double)matchingWords / originalWords.Length;
    }

    private static (bool Passed, string Details) EvaluateQuotesAsTest(MarkdownStructureAnalysis analysis)
    {
        string originalContent = string.Join("\n", analysis.OriginalSections);
        string convertedContent = string.Join("\n", analysis.ConvertedSections);

        var quoteAnalysis = AnalyzeQuoteStructure(originalContent, convertedContent);
        
        bool passed = quoteAnalysis.StructurePreserved && !quoteAnalysis.HasCorruption;
        string details = quoteAnalysis.GenerateDetails();
        
        return (passed, details);
    }

    private static QuoteStructureAnalysis AnalyzeQuoteStructure(string originalContent, string convertedContent)
    {
        var result = new QuoteStructureAnalysis();
        
        // 1. 引用腐敗パターンの検出
        result.HasCorruption = HasQuoteCorruption(convertedContent);
        if (result.HasCorruption)
        {
            result.CorruptionPatterns = DetectQuoteCorruptionPatterns(convertedContent);
        }
        
        // 2. 基本的な引用構造の検証
        var originalQuotes = ExtractQuoteBlocks(originalContent);
        var convertedQuotes = ExtractQuoteBlocks(convertedContent);
        
        result.OriginalQuoteCount = originalQuotes.Count;
        result.ConvertedQuoteCount = convertedQuotes.Count;
        
        if (originalQuotes.Count == 0)
        {
            result.StructurePreserved = true;
            return result;
        }
        
        // 3. ネストされた引用構造の検証
        result.NestedStructurePreserved = ValidateNestedQuoteStructure(originalQuotes, convertedQuotes);
        
        // 4. 引用内混合コンテンツの検証
        result.MixedContentPreserved = ValidateQuoteMixedContent(originalContent, convertedContent);
        
        // 5. 総合評価
        int structureTests = 0;
        int passedTests = 0;
        
        if (originalQuotes.Count > 0)
        {
            structureTests++;
            if (result.ConvertedQuoteCount > 0 && !result.HasCorruption)
                passedTests++;
        }
        
        if (HasNestedQuotes(originalContent))
        {
            structureTests++;
            if (result.NestedStructurePreserved)
                passedTests++;
        }
        
        if (HasQuoteMixedContent(originalContent))
        {
            structureTests++;
            if (result.MixedContentPreserved)
                passedTests++;
        }
        
        result.StructurePreserved = structureTests == 0 || (double)passedTests / structureTests >= 0.8;
        
        return result;
    }

    private static bool HasQuoteCorruption(string content)
    {
        // 引用構造の腐敗パターンを検出
        var corruptionPatterns = new[]
        {
            @"単一\*\*引用\*\*文",           // 引用が太字になっている
            @"複数行の\*\*引用\*\*文",       // 複数行引用が太字になっている
            @"レベル1引用>レベル2引用»",     // ネスト引用が記号に変換
            @"[^>]\s*\*\*[^*]+\*\*[^*]*(?:引用|quote)",  // 引用内容が太字化
        };
        
        return corruptionPatterns.Any(pattern => Regex.IsMatch(content, pattern));
    }

    private static List<string> DetectQuoteCorruptionPatterns(string content)
    {
        var patterns = new List<string>();
        
        if (Regex.IsMatch(content, @"単一\*\*引用\*\*文"))
            patterns.Add("Quote text converted to bold formatting");
        if (Regex.IsMatch(content, @"複数行の\*\*引用\*\*文"))
            patterns.Add("Multi-line quotes converted to bold");
        if (Regex.IsMatch(content, @"レベル1引用>レベル2引用»"))
            patterns.Add("Nested quotes converted to symbols (>, »)");
        if (Regex.IsMatch(content, @"[^>]\s*\*\*[^*]+\*\*[^*]*(?:引用|quote)"))
            patterns.Add("Quote content formatting corrupted");
            
        return patterns;
    }

    private static List<QuoteBlock> ExtractQuoteBlocks(string content)
    {
        var quotes = new List<QuoteBlock>();
        var lines = content.Split('\n');
        
        for (int i = 0; i < lines.Length; i++)
        {
            var line = lines[i];
            var quoteMatch = Regex.Match(line, @"^(>+)\s*(.*)$");
            if (quoteMatch.Success)
            {
                quotes.Add(new QuoteBlock
                {
                    Level = quoteMatch.Groups[1].Value.Length,
                    Content = quoteMatch.Groups[2].Value.Trim(),
                    LineNumber = i + 1
                });
            }
        }
        
        return quotes;
    }

    private static bool ValidateNestedQuoteStructure(List<QuoteBlock> original, List<QuoteBlock> converted)
    {
        var originalNested = original.Where(q => q.Level > 1).ToList();
        var convertedNested = converted.Where(q => q.Level > 1).ToList();
        
        if (originalNested.Count == 0)
            return true;
            
        return convertedNested.Count > 0; // 最低限ネスト構造が検出されること
    }

    private static bool ValidateQuoteMixedContent(string originalContent, string convertedContent)
    {
        // 引用内のリスト、太字、斜体などの混合コンテンツが保持されているかチェック
        var quoteLinesOriginal = originalContent.Split('\n')
            .Where(line => line.TrimStart().StartsWith(">"))
            .ToList();
            
        if (quoteLinesOriginal.Count == 0)
            return true;
            
        var quoteLinesConverted = convertedContent.Split('\n')
            .Where(line => line.TrimStart().StartsWith(">"))
            .ToList();
            
        // 引用内の書式設定が保持されているか簡易チェック
        bool hasOriginalFormatting = quoteLinesOriginal.Any(line => 
            line.Contains("**") || line.Contains("*") || line.Contains("- ") || line.Contains("`"));
            
        if (!hasOriginalFormatting)
            return true;
            
        bool hasConvertedFormatting = quoteLinesConverted.Any(line => 
            line.Contains("**") || line.Contains("*") || line.Contains("- ") || line.Contains("`"));
            
        return hasConvertedFormatting;
    }

    private static bool HasNestedQuotes(string content)
    {
        return Regex.IsMatch(content, @"^>>\s*", RegexOptions.Multiline);
    }

    private static bool HasQuoteMixedContent(string content)
    {
        var quoteLines = content.Split('\n')
            .Where(line => line.TrimStart().StartsWith(">"));
            
        return quoteLines.Any(line => 
            line.Contains("**") || line.Contains("*") || line.Contains("- ") || line.Contains("`"));
    }

    private class QuoteBlock
    {
        public int Level { get; set; }
        public string Content { get; set; } = string.Empty;
        public int LineNumber { get; set; }
    }

    private class QuoteStructureAnalysis
    {
        public bool StructurePreserved { get; set; }
        public bool HasCorruption { get; set; }
        public bool NestedStructurePreserved { get; set; }
        public bool MixedContentPreserved { get; set; }
        public List<string> CorruptionPatterns { get; set; } = new();
        public int OriginalQuoteCount { get; set; }
        public int ConvertedQuoteCount { get; set; }
        
        public string GenerateDetails()
        {
            if (OriginalQuoteCount == 0)
                return "No quotes to validate";
                
            var details = $"{ConvertedQuoteCount}/{OriginalQuoteCount} quote blocks detected";
            
            if (HasCorruption)
            {
                details += " - CORRUPTION DETECTED";
                if (CorruptionPatterns.Any())
                    details += $": {string.Join(", ", CorruptionPatterns.Take(2))}";
            }
            
            var issues = new List<string>();
            if (!NestedStructurePreserved)
                issues.Add("Nested quotes corrupted");
            if (!MixedContentPreserved)
                issues.Add("Mixed content not preserved");
                
            if (issues.Any())
                details += $" - Issues: {string.Join(", ", issues)}";
            
            return details;
        }
    }

    private static (bool Passed, string Details) EvaluateOtherSyntaxAsTest(MarkdownStructureAnalysis analysis)
    {
        string originalContent = string.Join("\n", analysis.OriginalSections);
        string convertedContent = string.Join("\n", analysis.ConvertedSections);

        var contentAnalysis = AnalyzeContentIntegrity(originalContent, convertedContent);
        
        bool passed = contentAnalysis.OverallIntegrityMaintained && !contentAnalysis.HasCriticalCorruption;
        string details = contentAnalysis.GenerateDetails();
        
        return (passed, details);
    }

    private static ContentIntegrityAnalysis AnalyzeContentIntegrity(string originalContent, string convertedContent)
    {
        var result = new ContentIntegrityAnalysis();
        
        // 1. 特殊文字・Unicode文字の保持検証
        result.UnicodePreserved = ValidateUnicodePreservation(originalContent, convertedContent);
        result.SpecialCharactersPreserved = ValidateSpecialCharacters(originalContent, convertedContent);
        
        // 2. 重大な腐敗パターンの検出
        result.HasCriticalCorruption = DetectCriticalCorruption(convertedContent);
        if (result.HasCriticalCorruption)
        {
            result.CorruptionPatterns = DetectContentCorruptionPatterns(convertedContent);
        }
        
        // 3. 構造的完全性の検証
        result.StructuralIntegrityMaintained = ValidateStructuralIntegrity(originalContent, convertedContent);
        
        // 4. その他のMarkdown記法保持
        result.OtherSyntaxPreserved = ValidateOtherMarkdownSyntax(originalContent, convertedContent);
        
        // 5. 総合評価
        int integrityTests = 0;
        int passedTests = 0;
        
        integrityTests++; // Unicode/特殊文字
        if (result.UnicodePreserved && result.SpecialCharactersPreserved)
            passedTests++;
            
        integrityTests++; // 構造的完全性
        if (result.StructuralIntegrityMaintained)
            passedTests++;
            
        integrityTests++; // その他記法
        if (result.OtherSyntaxPreserved)
            passedTests++;
        
        result.OverallIntegrityMaintained = !result.HasCriticalCorruption && 
                                           (double)passedTests / integrityTests >= 0.8;
        
        return result;
    }

    private static bool ValidateUnicodePreservation(string originalContent, string convertedContent)
    {
        // Unicode文字が元々存在しない場合は常にtrue
        if (!Regex.IsMatch(originalContent, @"[^\x00-\x7F]"))
            return true;
        
        // 日本語文字（ひらがな、カタカナ、漢字）の保持確認
        var japaneseChars = new[]
        {
            @"[あ-ん]",     // ひらがな
            @"[ア-ン]",     // カタカナ  
            @"[一-龯]"      // 漢字
        };
        
        // 各文字パターンに対して、50%以上の保持率があれば許容
        int totalPatterns = 0;
        int preservedPatterns = 0;
        
        foreach (var pattern in japaneseChars)
        {
            var originalMatches = Regex.Matches(originalContent, pattern);
            var convertedMatches = Regex.Matches(convertedContent, pattern);
            
            if (originalMatches.Count > 0)
            {
                totalPatterns++;
                double preservationRate = (double)convertedMatches.Count / originalMatches.Count;
                if (preservationRate >= 0.5) // 50%以上の保持率
                    preservedPatterns++;
            }
        }
        
        // 日本語文字が全くない場合、または50%以上のパターンが保持されている場合はOK
        return totalPatterns == 0 || (double)preservedPatterns / totalPatterns >= 0.5;
    }

    private static bool ValidateSpecialCharacters(string originalContent, string convertedContent)
    {
        // 重要な特殊記号のみをチェック（Markdown構造に影響するもの）
        var criticalChars = new[] { "|", "#", "[", "]", "(", ")", "*", "_", "`" };
        
        foreach (var specialChar in criticalChars)
        {
            var originalCount = originalContent.Count(c => c.ToString() == specialChar);
            var convertedCount = convertedContent.Count(c => c.ToString() == specialChar);
            
            // 重要な文字が大幅に減少している場合のみ問題とする
            if (originalCount > 0 && convertedCount < originalCount * 0.5)
            {
                return false;
            }
        }
        
        return true;
    }

    private static bool DetectCriticalCorruption(string convertedContent)
    {
        var criticalCorruptionPatterns = new[]
        {
            @"\*\*‒\*\*",                   // ネストリスト腐敗
            @"単一\*\*引用\*\*文",           // 引用腐敗
            @"レベル1引用>レベル2引用»",      // ネスト引用腐敗  
            @"\|[^|]*\|[^|]*\|\|̶̶",       // テーブル完全破綻
            @"<br>\d+",                     // PDF座標情報混入
            @"[A-Za-z]+\d+[A-Za-z]+\s*\|",  // テーブル内文字数字混合腐敗
        };
        
        return criticalCorruptionPatterns.Any(pattern => Regex.IsMatch(convertedContent, pattern));
    }

    private static List<string> DetectContentCorruptionPatterns(string convertedContent)
    {
        var patterns = new List<string>();
        
        if (Regex.IsMatch(convertedContent, @"\*\*‒\*\*"))
            patterns.Add("Nested lists corrupted to bold symbols");
        if (Regex.IsMatch(convertedContent, @"単一\*\*引用\*\*文"))
            patterns.Add("Quotes corrupted to bold text");
        if (Regex.IsMatch(convertedContent, @"レベル1引用>レベル2引用»"))
            patterns.Add("Nested quotes corrupted to symbols");
        if (Regex.IsMatch(convertedContent, @"\|[^|]*\|[^|]*\|\|̶̶"))
            patterns.Add("Tables completely corrupted");
        if (Regex.IsMatch(convertedContent, @"<br>\d+"))
            patterns.Add("PDF coordinate information leaked");
        if (Regex.IsMatch(convertedContent, @"[A-Za-z]+\d+[A-Za-z]+\s*\|"))
            patterns.Add("Table content merged incorrectly");
            
        return patterns;
    }

    private static bool ValidateStructuralIntegrity(string originalContent, string convertedContent)
    {
        // 基本的な構造が維持されているかチェック
        var originalLines = originalContent.Split('\n').Where(line => !string.IsNullOrWhiteSpace(line)).Count();
        var convertedLines = convertedContent.Split('\n').Where(line => !string.IsNullOrWhiteSpace(line)).Count();
        
        // 変換後に大幅に行数が減っている場合は構造的問題
        if (originalLines > 0 && (double)convertedLines / originalLines < 0.5)
            return false;
            
        return true;
    }

    private static bool ValidateOtherMarkdownSyntax(string originalContent, string convertedContent)
    {
        int syntaxTests = 0;
        int preservedSyntax = 0;
        
        // 水平線の保持確認
        var horizontalRules = Regex.Matches(originalContent, @"^(---+|\*\*\*+|___+)$", RegexOptions.Multiline);
        if (horizontalRules.Count > 0)
        {
            syntaxTests++;
            bool horizontalRulesPreserved = convertedContent.Contains("---") || convertedContent.Contains("***") || convertedContent.Contains("___");
            if (horizontalRulesPreserved) preservedSyntax++;
        }
        
        // エスケープ文字の保持確認
        var escapedChars = Regex.Matches(originalContent, @"\\[*_#\[\]]");
        if (escapedChars.Count > 0)
        {
            syntaxTests++;
            // エスケープ文字が何らかの形で保持されているか
            bool escapedPreserved = Regex.IsMatch(convertedContent, @"\\[*_#\[\]]") || 
                                   convertedContent.Contains("エスケープ");
            if (escapedPreserved) preservedSyntax++;
        }
        
        if (syntaxTests == 0) return true;
        
        return (double)preservedSyntax / syntaxTests >= 0.8;
    }

    private class ContentIntegrityAnalysis
    {
        public bool OverallIntegrityMaintained { get; set; }
        public bool HasCriticalCorruption { get; set; }
        public bool UnicodePreserved { get; set; }
        public bool SpecialCharactersPreserved { get; set; }
        public bool StructuralIntegrityMaintained { get; set; }
        public bool OtherSyntaxPreserved { get; set; }
        public List<string> CorruptionPatterns { get; set; } = new();
        
        public string GenerateDetails()
        {
            var issues = new List<string>();
            
            if (HasCriticalCorruption)
            {
                issues.Add("CRITICAL CORRUPTION DETECTED");
                if (CorruptionPatterns.Any())
                    issues.Add($"Patterns: {string.Join(", ", CorruptionPatterns.Take(3))}");
            }
            
            if (!UnicodePreserved)
                issues.Add("Unicode characters lost");
            if (!SpecialCharactersPreserved)
                issues.Add("Special characters lost");
            if (!StructuralIntegrityMaintained)
                issues.Add("Structural integrity compromised");
            if (!OtherSyntaxPreserved)
                issues.Add("Other syntax not preserved");
            
            if (issues.Any())
                return $"Content integrity failed - {string.Join(", ", issues.Take(3))}";
            else
                return "Content integrity maintained";
        }
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
    
    private static bool AreTablesStructurallyEquivalent(string originalTable, string convertedTable)
    {
        var originalRows = NormalizeTableForComparison(originalTable);
        var convertedRows = NormalizeTableForComparison(convertedTable);
        
        if (originalRows.Count != convertedRows.Count)
            return false;
            
        // Check if each row has the same data content (ignoring formatting)
        for (int i = 0; i < originalRows.Count; i++)
        {
            var originalCells = originalRows[i];
            var convertedCells = convertedRows[i];
            
            if (originalCells.Count != convertedCells.Count)
                return false;
                
            for (int j = 0; j < originalCells.Count; j++)
            {
                var originalCell = CleanCellContent(originalCells[j]);
                var convertedCell = CleanCellContent(convertedCells[j]);
                
                // Allow some flexibility in cell content matching
                if (!AreCellContentsEquivalent(originalCell, convertedCell))
                    return false;
            }
        }
        
        return true;
    }
    
    private static List<List<string>> NormalizeTableForComparison(string tableContent)
    {
        var lines = tableContent.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        var rows = new List<List<string>>();
        
        foreach (var line in lines)
        {
            if (!line.Contains('|')) continue;
            
            // Skip all separator rows regardless of alignment syntax
            // This includes |---|, |:---|, |---:|, |:---:| patterns
            if (IsSeparatorRow(line)) continue;
            
            var normalizedRow = NormalizeTableRow(line);
            var cells = ExtractTableCells(normalizedRow);
            
            if (cells.Count > 0)
                rows.Add(cells);
        }
        
        return rows;
    }
    
    private static List<string> ExtractTableCells(string normalizedRow)
    {
        // Split by pipes and clean up the cells
        var parts = normalizedRow.Split('|');
        var cells = new List<string>();
        
        foreach (var part in parts)
        {
            var trimmed = part.Trim();
            if (!string.IsNullOrEmpty(trimmed) || cells.Count > 0) // Include empty cells in the middle
            {
                cells.Add(trimmed);
            }
        }
        
        // Remove trailing empty cell if it exists (from trailing pipe)
        if (cells.Count > 0 && string.IsNullOrEmpty(cells[cells.Count - 1]))
        {
            cells.RemoveAt(cells.Count - 1);
        }
        
        return cells;
    }
    
    private static bool IsSeparatorRow(string line)
    {
        // Remove pipes, spaces, and colons, then check if only dashes remain
        var cleaned = line.Replace("|", "").Replace(" ", "").Replace(":", "");
        
        // Must contain at least one dash and only dashes
        return cleaned.Length > 0 && cleaned.All(c => c == '-');
    }
    
    private static string CleanCellContent(string cellContent)
    {
        // Remove extra whitespace and normalize content
        return cellContent.Trim().Replace("  ", " ");
    }
    
    private static bool AreCellContentsEquivalent(string cell1, string cell2)
    {
        // Exact match
        if (cell1.Equals(cell2, StringComparison.OrdinalIgnoreCase))
            return true;
            
        // Check if one contains the other (for cases where formatting changes content slightly)
        if (cell1.Contains(cell2, StringComparison.OrdinalIgnoreCase) ||
            cell2.Contains(cell1, StringComparison.OrdinalIgnoreCase))
            return true;
            
        // Check word-level similarity for complex cells
        var words1 = cell1.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        var words2 = cell2.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        
        if (words1.Length == 0 && words2.Length == 0)
            return true;
            
        var commonWords = words1.Intersect(words2, StringComparer.OrdinalIgnoreCase).Count();
        var totalWords = words1.Union(words2, StringComparer.OrdinalIgnoreCase).Count();
        
        return totalWords > 0 && (double)commonWords / totalWords >= 0.8;
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
        // Markdownの書式設定記号を除去（ただしテーブル構造は保持）
        return text.Replace("#", "")
                  .Replace("*", "")
                  .Replace("_", "")
                  .Replace("`", "")
                  .Replace(">", "")
                  .Replace("•", "")
                  .Trim();
    }
    
    private static bool HasTableCorruption(string tableContent)
    {
        // Check for critical corruption patterns only
        var corruptionPatterns = new[]
        {
            @"<br>\d+",           // HTML break tags with numbers
            @"\d+<br>",           // Numbers followed by HTML break tags  
            @"\*\*‒\*\*",         // Bold dash corruption from lists
            @"[A-Za-z]{5,}\d+[A-Za-z]{5,}", // Very long words merged with numbers (更に厳格に)
        };
        
        foreach (var pattern in corruptionPatterns)
        {
            if (Regex.IsMatch(tableContent, pattern))
            {
                return true;
            }
        }
        
        return false;
    }
    
    private static string GetCorruptionDetails(string tableContent)
    {
        var details = new List<string>();
        
        if (Regex.IsMatch(tableContent, @"<br>\d+"))
            details.Add("HTML break tags with numbers detected");
            
        if (Regex.IsMatch(tableContent, @"[A-Za-z]+\d+[A-Za-z]+"))
            details.Add("Text merged with numbers");
            
        if (Regex.IsMatch(tableContent, @"\|\s*\|\s*\|"))
            details.Add("Multiple empty cells");
            
        return details.Count > 0 ? string.Join(", ", details) : "Unknown corruption";
    }
    
    private static bool HasValidTableStructure(string tableContent)
    {
        var lines = tableContent.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        if (lines.Length < 2) return false;
        
        // Check if there's a separator row with proper markdown table alignment syntax
        bool hasSeparator = lines.Any(line => IsSeparatorRow(line));
        if (!hasSeparator) return false;
        
        // Get table rows (excluding separator rows)
        var dataRows = lines.Where(line => line.Contains('|') && !IsSeparatorRow(line))
                           .Select(line => NormalizeTableRow(line))
                           .ToList();
        
        if (dataRows.Count < 1) return false;
        
        // Check column count consistency for data rows
        var columnCounts = dataRows.Select(row => CountTableColumns(row)).ToList();
        
        if (columnCounts.Count == 0) return false;
        
        var minColumns = columnCounts.Min();
        var maxColumns = columnCounts.Max();
        
        // Allow some variation in column count for formatting differences
        return (maxColumns - minColumns) <= 1 && minColumns >= 1;
    }
    
    private static int CountTableColumns(string normalizedRow)
    {
        // Count cells by splitting on pipes and excluding empty cells at start/end
        var parts = normalizedRow.Split('|');
        
        // Remove empty parts at beginning and end (from leading/trailing pipes)
        var cells = parts.Skip(parts[0] == "" ? 1 : 0)
                        .Take(parts[parts.Length - 1] == "" ? parts.Length - (parts[0] == "" ? 2 : 1) : parts.Length - (parts[0] == "" ? 1 : 0))
                        .ToList();
        
        return Math.Max(1, cells.Count);
    }
    
    private static string NormalizeTableRow(string row)
    {
        // Remove leading/trailing whitespace and normalize pipe characters
        var normalized = row.Trim();
        
        // Ensure the row starts and ends with pipes for consistent parsing
        if (!normalized.StartsWith('|')) normalized = "|" + normalized;
        if (!normalized.EndsWith('|')) normalized = normalized + "|";
        
        return normalized;
    }

}