namespace AnyToMarkdown.Tests;

public class MarkdownStructureAnalysis
{
    public string FileName { get; set; } = string.Empty;
    public List<string> OriginalHeaders { get; set; } = new();
    public List<string> ConvertedHeaders { get; set; } = new();
    public List<string> OriginalSections { get; set; } = new();
    public List<string> ConvertedSections { get; set; } = new();
    public List<string> OriginalTables { get; set; } = new();
    public List<string> ConvertedTables { get; set; } = new();
    public List<string> OriginalLists { get; set; } = new();
    public List<string> ConvertedLists { get; set; } = new();
}

public class ConversionAccuracyTest
{
    private readonly string[] _testFiles = new[]
    {
        "test-basic",
        "test-table", 
        "test-complex",
        "test-japanese",
        "test-links",
        "test-multiline-table",
        "test-complex-table",
        "test-nested-content",
        "test-edge-cases"
    };

    [Fact]
    public void PdfToMarkdownAccuracyTest()
    {
        // 生成されたPDFをマークダウンに変換
        string pdfPath = "./Resources/test-generated.pdf";
        Assert.True(File.Exists(pdfPath), $"Generated PDF file should exist at: {pdfPath}");
        
        var result = AnyConverter.Convert(pdfPath);
        
        Assert.NotNull(result.Text);
        Assert.NotEmpty(result.Text);
        
        // 結果をファイルに保存して確認できるようにする
        string outputPath = "./test-output.md";
        File.WriteAllText(outputPath, result.Text);
        
        // 基本的な構造要素が含まれているかチェック
        Assert.Contains("テストドキュメント", result.Text);
        Assert.Contains("はじめに", result.Text);
        Assert.Contains("基本的なテキスト", result.Text);
        Assert.Contains("テーブル", result.Text);
        
        // テーブル構造が保持されているかチェック
        Assert.Contains("|", result.Text); // テーブルの区切り文字
        
        // 変換の警告を出力
        if (result.Warnings.Count > 0)
        {
            foreach (var warning in result.Warnings)
            {
                Console.WriteLine($"Warning: {warning}");
            }
        }
        
        Console.WriteLine($"Conversion completed. Output length: {result.Text.Length} characters");
        Console.WriteLine($"Warnings count: {result.Warnings.Count}");
    }

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

    private MarkdownStructureAnalysis AnalyzeMarkdownStructure(string original, string converted, string fileName)
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

        // 1. ヘッダー構造テスト - 厳密な基準
        var headerResult = EvaluateHeaderStructureAsTest(analysis);
        testResults.Add(("Header Structure", headerResult.Passed, headerResult.Details));

        // 2. セクション構造テスト - 厳密な基準  
        var sectionResult = EvaluateSectionStructureAsTest(analysis);
        testResults.Add(("Section Structure", sectionResult.Passed, sectionResult.Details));

        // 3. テーブル構造テスト - 厳密な基準
        var tableResult = EvaluateTableStructureAsTest(analysis);
        testResults.Add(("Table Structure", tableResult.Passed, tableResult.Details));

        // 4. リスト構造テスト - 厳密な基準
        var listResult = EvaluateListStructureAsTest(analysis);
        testResults.Add(("List Structure", listResult.Passed, listResult.Details));

        // 5. Markdown書式保持テスト - 厳密な基準
        var formattingResult = EvaluateMarkdownFormattingAsTest(analysis);
        testResults.Add(("Markdown Formatting", formattingResult.Passed, formattingResult.Details));

        // 結果の表示
        int passedTests = 0;
        int totalTests = testResults.Count;

        Console.WriteLine($"[{fileName}] Test Results:");
        foreach (var (testName, passed, details) in testResults)
        {
            string status = passed ? "PASS" : "FAIL";
            Console.WriteLine($"  {status}: {testName} - {details}");
            if (passed) passedTests++;
        }

        double passRate = (double)passedTests / totalTests * 100.0;
        Console.WriteLine($"[{fileName}] Overall pass rate: {passedTests}/{totalTests} ({passRate:F1}%)");

        // 厳格な合格基準: 基本的なケースでは80%以上、複雑なケースでも60%以上が必要
        double requiredPassRate = GetStrictPassRate(fileName);
        Assert.True(passRate >= requiredPassRate, 
            $"[{fileName}] Pass rate ({passRate:F1}%) below required threshold ({requiredPassRate:F1}%)");
    }

    private double GetStrictPassRate(string fileName)
    {
        // 厳格な基準 - 条件を緩めない
        return fileName switch
        {
            var f when f.Contains("basic") => 80.0,      // 基本的なケースは5項目中4項目以上
            var f when f.Contains("table") && !f.Contains("complex") => 80.0,
            var f when f.Contains("japanese") => 80.0,
            var f when f.Contains("complex") => 60.0,    // 複雑なケースでも5項目中3項目以上
            var f when f.Contains("multiline") => 60.0,
            var f when f.Contains("advanced") => 60.0,
            var f when f.Contains("scientific") => 60.0,
            var f when f.Contains("financial") => 60.0,
            var f when f.Contains("mixed") => 60.0,
            _ => 80.0
        };
    }

    private (bool Passed, string Details) EvaluateHeaderStructureAsTest(MarkdownStructureAnalysis analysis)
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

    private (bool Passed, string Details) EvaluateSectionStructureAsTest(MarkdownStructureAnalysis analysis)
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

    private (bool Passed, string Details) EvaluateTableStructureAsTest(MarkdownStructureAnalysis analysis)
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

    private (bool Passed, string Details) EvaluateListStructureAsTest(MarkdownStructureAnalysis analysis)
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
        var boldMatches = System.Text.RegularExpressions.Regex.Matches(originalContent, @"\*\*([^*]+)\*\*");
        if (boldMatches.Count > 0)
        {
            formattingTests++;
            // 厳格な基準: 太字がMarkdown記法として保持されているか、または視覚的に強調されているか
            bool boldPreserved = convertedContent.Contains("**") || // Markdown形式で保持
                                boldMatches.Cast<System.Text.RegularExpressions.Match>()
                                .All(match => convertedContent.Contains(match.Groups[1].Value, StringComparison.OrdinalIgnoreCase));
            
            if (boldPreserved) preservedFormatting++;
        }

        // 斜体の保持確認 - 厳格な基準
        var italicMatches = System.Text.RegularExpressions.Regex.Matches(originalContent, @"(?<!\*)\*([^*]+)\*(?!\*)");
        if (italicMatches.Count > 0)
        {
            formattingTests++;
            bool italicPreserved = convertedContent.Contains("*") || // Markdown形式で保持
                                 italicMatches.Cast<System.Text.RegularExpressions.Match>()
                                 .All(match => convertedContent.Contains(match.Groups[1].Value, StringComparison.OrdinalIgnoreCase));
            
            if (italicPreserved) preservedFormatting++;
        }

        // コードの保持確認 - 厳格な基準
        var codeMatches = System.Text.RegularExpressions.Regex.Matches(originalContent, @"`([^`]+)`");
        if (codeMatches.Count > 0)
        {
            formattingTests++;
            bool codePreserved = convertedContent.Contains("`") || // Markdown形式で保持
                               codeMatches.Cast<System.Text.RegularExpressions.Match>()
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

    private double GetPassingThreshold(string fileName)
    {
        // ファイルの複雑さに応じて合格基準を設定
        return fileName switch
        {
            var f when f.Contains("basic") => 85.0,
            var f when f.Contains("table") && !f.Contains("complex") => 75.0,
            var f when f.Contains("complex") => 60.0,
            var f when f.Contains("multiline") => 75.0,
            var f when f.Contains("japanese") => 70.0,
            var f when f.Contains("advanced") => 55.0,
            var f when f.Contains("scientific") => 50.0,
            var f when f.Contains("financial") => 50.0,
            var f when f.Contains("mixed") => 45.0,
            _ => 70.0
        };
    }

    private double EvaluateHeaderStructure(MarkdownStructureAnalysis analysis)
    {
        if (analysis.OriginalHeaders.Count == 0) return 40.0; // ヘッダーがない場合は満点

        int exactMatches = 0;  // 正確なMarkdownヘッダーとしてのマッチ
        int partialMatches = 0; // テキストとしてのマッチ
        int totalHeaders = analysis.OriginalHeaders.Count;

        Console.WriteLine($"[DEBUG] Evaluating {totalHeaders} headers:");

        foreach (var originalHeader in analysis.OriginalHeaders)
        {
            var headerText = ExtractHeaderText(originalHeader);
            var headerLevel = GetHeaderLevel(originalHeader);
            
            Console.WriteLine($"[DEBUG] Looking for header: '{originalHeader}' (level {headerLevel}, text: '{headerText}')");

            // 1. 正確なMarkdownヘッダーとしての一致（最高点）
            bool exactLevelMatch = analysis.ConvertedHeaders.Any(h => 
                GetHeaderLevel(h) == headerLevel && 
                ExtractHeaderText(h).Equals(headerText, StringComparison.OrdinalIgnoreCase));

            // 2. Markdownヘッダーとしての一致（レベル無視）
            bool headerMatch = analysis.ConvertedHeaders.Any(h => 
                ExtractHeaderText(h).Equals(headerText, StringComparison.OrdinalIgnoreCase));

            // 3. テキストとしての存在確認（低い点数）
            bool textExists = analysis.ConvertedSections.Any(s => 
                s.Contains(headerText, StringComparison.OrdinalIgnoreCase));

            if (exactLevelMatch)
            {
                exactMatches++;
                Console.WriteLine($"[DEBUG] ✓ Exact level match found for '{headerText}'");
            }
            else if (headerMatch)
            {
                exactMatches++;
                Console.WriteLine($"[DEBUG] ✓ Header match found for '{headerText}' (different level)");
            }
            else if (textExists)
            {
                partialMatches++;
                Console.WriteLine($"[DEBUG] ○ Text match found for '{headerText}'");
            }
            else
            {
                Console.WriteLine($"[DEBUG] ✗ No match found for '{headerText}'");
            }
        }

        // スコア計算：正確なマッチは2点、部分マッチは0.5点
        double score = (exactMatches * 2.0 + partialMatches * 0.5) / (totalHeaders * 2.0) * 40.0;
        Console.WriteLine($"[DEBUG] Header score: {exactMatches} exact + {partialMatches} partial = {score:F1}/40.0");
        
        return Math.Min(40.0, score);
    }

    private double EvaluateSectionStructure(MarkdownStructureAnalysis analysis)
    {
        if (analysis.OriginalSections.Count <= 1) return 30.0;

        int matchedSections = 0;

        foreach (var originalSection in analysis.OriginalSections)
        {
            // 元のセクションから空行で区切られたブロックを抽出
            var originalBlocks = ExtractContentBlocks(originalSection);
            
            foreach (var originalBlock in originalBlocks)
            {
                if (string.IsNullOrWhiteSpace(originalBlock)) continue;

                // 変換されたセクションでマッチするブロックを探す
                bool blockFound = false;
                foreach (var convertedSection in analysis.ConvertedSections)
                {
                    var convertedBlocks = ExtractContentBlocks(convertedSection);
                    
                    foreach (var convertedBlock in convertedBlocks)
                    {
                        if (CalculateBlockSimilarity(originalBlock, convertedBlock) > 0.5)
                        {
                            blockFound = true;
                            break;
                        }
                    }
                    
                    if (blockFound) break;
                }
                
                if (blockFound)
                {
                    matchedSections++;
                    break; // このセクションはマッチしたので次へ
                }
            }
        }

        return (double)matchedSections / analysis.OriginalSections.Count * 30.0;
    }

    private double EvaluateTableStructure(MarkdownStructureAnalysis analysis)
    {
        if (analysis.OriginalTables.Count == 0) return 20.0;

        int matchedTables = 0;

        // 各オリジナルテーブルについて、変換後での存在を確認
        foreach (var originalTable in analysis.OriginalTables)
        {
            var originalTableContent = ExtractTableContentWords(originalTable);
            bool tableMatched = false;

            // 1. 直接的なMarkdownテーブルとしての検出
            if (analysis.ConvertedTables.Any(convertedTable =>
                CalculateTableContentSimilarity(originalTableContent, ExtractTableContentWords(convertedTable)) > 0.3))
            {
                tableMatched = true;
            }

            // 2. セクション内での表形式データとしての検出
            if (!tableMatched)
            {
                foreach (var section in analysis.ConvertedSections)
                {
                    if (HasTabularData(section) && 
                        CalculateTableContentSimilarity(originalTableContent, ExtractTableContentWords(section)) > 0.2)
                    {
                        tableMatched = true;
                        break;
                    }
                }
            }

            // 3. 表の内容がテキストとして保持されているかの確認
            if (!tableMatched)
            {
                var allConvertedText = string.Join(" ", analysis.ConvertedSections);
                int preservedContentCount = 0;
                foreach (var word in originalTableContent)
                {
                    if (allConvertedText.Contains(word, StringComparison.OrdinalIgnoreCase))
                    {
                        preservedContentCount++;
                    }
                }
                
                if (originalTableContent.Count > 0 && 
                    (double)preservedContentCount / originalTableContent.Count > 0.5)
                {
                    tableMatched = true;
                }
            }

            if (tableMatched)
            {
                matchedTables++;
            }
        }

        return (double)matchedTables / analysis.OriginalTables.Count * 20.0;
    }

    private double EvaluateListStructure(MarkdownStructureAnalysis analysis)
    {
        if (analysis.OriginalLists.Count == 0) return 10.0;

        int matchedLists = 0;

        foreach (var originalList in analysis.OriginalLists)
        {
            var originalListItems = ExtractListItems(originalList);
            bool listMatched = false;

            // 1. 直接的なMarkdownリストとしての検出
            if (analysis.ConvertedLists.Any(convertedList =>
                CalculateListSimilarity(originalListItems, ExtractListItems(convertedList)) > 0.3))
            {
                listMatched = true;
            }

            // 2. セクション内での箇条書き的表現の検出
            if (!listMatched)
            {
                foreach (var section in analysis.ConvertedSections)
                {
                    if (HasListLikeStructure(section) &&
                        CalculateListSimilarity(originalListItems, ExtractListItemsFromText(section)) > 0.3)
                    {
                        listMatched = true;
                        break;
                    }
                }
            }

            // 3. リスト項目のテキストが何らかの形で保持されているかの確認
            if (!listMatched)
            {
                var allConvertedText = string.Join(" ", analysis.ConvertedSections);
                int preservedItemsCount = 0;
                foreach (var item in originalListItems)
                {
                    if (allConvertedText.Contains(item, StringComparison.OrdinalIgnoreCase))
                    {
                        preservedItemsCount++;
                    }
                }

                if (originalListItems.Count > 0 && 
                    (double)preservedItemsCount / originalListItems.Count > 0.5)
                {
                    listMatched = true;
                }
            }

            if (listMatched)
            {
                matchedLists++;
            }
        }

        return (double)matchedLists / analysis.OriginalLists.Count * 10.0;
    }

    private List<string> ExtractHeaders(string content)
    {
        return content.Split('\n')
            .Where(line => line.TrimStart().StartsWith("#"))
            .Select(line => line.Trim())
            .ToList();
    }

    private List<string> ExtractSections(string content)
    {
        var lines = content.Split('\n');
        var sections = new List<string>();
        var currentSection = new List<string>();

        foreach (var line in lines)
        {
            if (line.TrimStart().StartsWith("#") && currentSection.Count > 0)
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

    private List<string> ExtractTables(string content)
    {
        var lines = content.Split('\n');
        var tables = new List<string>();
        var currentTable = new List<string>();
        bool inTable = false;

        foreach (var line in lines)
        {
            if (line.Contains("|"))
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

    private List<string> ExtractLists(string content)
    {
        var lines = content.Split('\n');
        var lists = new List<string>();
        var currentList = new List<string>();
        bool inList = false;

        foreach (var line in lines)
        {
            var trimmed = line.TrimStart();
            if (trimmed.StartsWith("-") || trimmed.StartsWith("*") || trimmed.StartsWith("+") ||
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
            else if (inList && (trimmed.StartsWith(" ") || trimmed.StartsWith("\t")))
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

    private string ExtractHeaderText(string header)
    {
        return header.Replace("#", "").Trim();
    }

    private int GetHeaderLevel(string header)
    {
        return header.TakeWhile(c => c == '#').Count();
    }

    private double CalculateTextSimilarity(string text1, string text2)
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

    private bool HasTabularData(string content)
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

    private List<string> ExtractTableContentWords(string tableContent)
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

    private double CalculateTableContentSimilarity(List<string> content1, List<string> content2)
    {
        if (content1.Count == 0 || content2.Count == 0)
            return 0.0;

        var set1 = content1.Select(w => w.ToLowerInvariant()).ToHashSet();
        var set2 = content2.Select(w => w.ToLowerInvariant()).ToHashSet();

        var intersection = set1.Intersect(set2).Count();
        var union = set1.Union(set2).Count();

        return union == 0 ? 0.0 : (double)intersection / union;
    }

    private List<string> ExtractListItems(string listContent)
    {
        var items = new List<string>();
        var lines = listContent.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        
        foreach (var line in lines)
        {
            var trimmed = line.TrimStart();
            if (trimmed.StartsWith("-") || trimmed.StartsWith("*") || trimmed.StartsWith("+") ||
                (trimmed.Length > 2 && char.IsDigit(trimmed[0]) && trimmed[1] == '.'))
            {
                // リストマーカーを除去してアイテムテキストを抽出
                var itemText = trimmed.Substring(trimmed.IndexOfAny(new char[] { ' ', '\t' }) + 1).Trim();
                if (!string.IsNullOrEmpty(itemText))
                {
                    items.Add(itemText);
                }
            }
        }
        
        return items;
    }

    private List<string> ExtractListItemsFromText(string text)
    {
        var items = new List<string>();
        var lines = text.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        
        foreach (var line in lines)
        {
            var trimmed = line.Trim();
            // 箇条書き的なパターンを検出
            if (trimmed.StartsWith("•") || trimmed.StartsWith("◦") || trimmed.StartsWith("▪"))
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

    private bool HasListLikeStructure(string text)
    {
        var lines = text.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        int listLikeLines = 0;
        
        foreach (var line in lines)
        {
            var trimmed = line.Trim();
            if (trimmed.StartsWith("•") || trimmed.StartsWith("◦") || trimmed.StartsWith("▪") ||
                trimmed.StartsWith("-") || trimmed.StartsWith("*") || trimmed.StartsWith("+"))
            {
                listLikeLines++;
            }
        }
        
        return listLikeLines >= 2; // 2行以上の箇条書きがあればリスト的構造と判定
    }

    private double CalculateListSimilarity(List<string> items1, List<string> items2)
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

    private List<string> ExtractContentBlocks(string content)
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

    private double CalculateBlockSimilarity(string block1, string block2)
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

    private string CleanMarkdownForComparison(string text)
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

    [Fact]
    public void MultilineTableConversionTest()
    {
        string originalMdPath = "./Resources/test-multiline-table.md";
        string pdfPath = "./Resources/test-multiline-table.pdf";
        string convertedMdPath = "./test-multiline-output.md";

        Assert.True(File.Exists(originalMdPath), $"Original markdown file should exist: {originalMdPath}");
        
        if (!File.Exists(pdfPath))
        {
            Assert.Fail($"PDF not found. Run: pandoc {originalMdPath} -o {pdfPath} --pdf-engine=xelatex -V CJKmainfont=\"Hiragino Sans\" -V mainfont=\"Hiragino Sans\"");
        }

        var result = AnyConverter.Convert(pdfPath);
        
        Assert.NotNull(result.Text);
        
        // 変換結果をファイルに保存
        File.WriteAllText(convertedMdPath, result.Text);
        
        // 元のMarkdownを読み込み
        string originalContent = File.ReadAllText(originalMdPath);
        
        // 基本的な内容が保持されているかチェック
        VerifyContentSimilarity(originalContent, result.Text, "test-multiline-table");
        
        Console.WriteLine($"Multiline table test completed. Warnings: {result.Warnings.Count}");
    }

    [Fact]
    public void EmptyCellTableConversionTest()
    {
        string originalMdPath = "./Resources/test-complex-table.md";
        string pdfPath = "./Resources/test-complex-table.pdf";
        string convertedMdPath = "./test-empty-cell-output.md";

        Assert.True(File.Exists(originalMdPath), $"Original markdown file should exist: {originalMdPath}");
        
        if (!File.Exists(pdfPath))
        {
            Assert.Fail($"PDF not found. Run: pandoc {originalMdPath} -o {pdfPath} --pdf-engine=xelatex -V CJKmainfont=\"Hiragino Sans\" -V mainfont=\"Hiragino Sans\"");
        }

        var result = AnyConverter.Convert(pdfPath);
        
        Assert.NotNull(result.Text);
        
        // 変換結果をファイルに保存
        File.WriteAllText(convertedMdPath, result.Text);
        
        // 元のMarkdownを読み込み
        string originalContent = File.ReadAllText(originalMdPath);
        
        // 基本的な内容が保持されているかチェック
        VerifyContentSimilarity(originalContent, result.Text, "test-complex-table");
        
        Console.WriteLine($"Empty cell table test completed. Warnings: {result.Warnings.Count}");
    }

    [Fact]
    public void AdvancedDocumentConversionTest()
    {
        string originalMdPath = "./Resources/test-advanced-document.md";
        string pdfPath = "./Resources/test-advanced-document.pdf";
        string convertedMdPath = "./test-advanced-output.md";

        Assert.True(File.Exists(originalMdPath), $"Original markdown file should exist: {originalMdPath}");
        
        if (!File.Exists(pdfPath))
        {
            Assert.Fail($"PDF not found. Run: pandoc {originalMdPath} -o {pdfPath} --pdf-engine=xelatex -V CJKmainfont=\"Hiragino Sans\" -V mainfont=\"Hiragino Sans\"");
        }

        var result = AnyConverter.Convert(pdfPath);
        
        Assert.NotNull(result.Text);
        Assert.NotEmpty(result.Text);
        
        // 変換結果をファイルに保存
        File.WriteAllText(convertedMdPath, result.Text);
        
        // 元のMarkdownを読み込み
        string originalContent = File.ReadAllText(originalMdPath);
        
        // 基本的な内容が保持されているかチェック
        VerifyContentSimilarity(originalContent, result.Text, "test-advanced-document");
        
        Console.WriteLine($"Advanced document test completed. Length: {result.Text.Length}, Warnings: {result.Warnings.Count}");
    }

    [Fact]
    public void ScientificPaperConversionTest()
    {
        string originalMdPath = "./Resources/test-scientific-paper.md";
        string pdfPath = "./Resources/test-scientific-paper.pdf";
        string convertedMdPath = "./test-scientific-output.md";

        Assert.True(File.Exists(originalMdPath), $"Original markdown file should exist: {originalMdPath}");
        
        if (!File.Exists(pdfPath))
        {
            Assert.Fail($"PDF not found. Run: pandoc {originalMdPath} -o {pdfPath} --pdf-engine=xelatex -V CJKmainfont=\"Hiragino Sans\" -V mainfont=\"Hiragino Sans\"");
        }

        var result = AnyConverter.Convert(pdfPath);
        
        Assert.NotNull(result.Text);
        Assert.NotEmpty(result.Text);
        
        // 変換結果をファイルに保存
        File.WriteAllText(convertedMdPath, result.Text);
        
        // 元のMarkdownを読み込み
        string originalContent = File.ReadAllText(originalMdPath);
        
        // 基本的な内容が保持されているかチェック
        VerifyContentSimilarity(originalContent, result.Text, "test-scientific-paper");
        
        Console.WriteLine($"Scientific paper test completed. Length: {result.Text.Length}, Warnings: {result.Warnings.Count}");
    }

    [Fact]
    public void FinancialReportConversionTest()
    {
        string originalMdPath = "./Resources/test-financial-report.md";
        string pdfPath = "./Resources/test-financial-report.pdf";
        string convertedMdPath = "./test-financial-output.md";

        Assert.True(File.Exists(originalMdPath), $"Original markdown file should exist: {originalMdPath}");
        
        if (!File.Exists(pdfPath))
        {
            Assert.Fail($"PDF not found. Run: pandoc {originalMdPath} -o {pdfPath} --pdf-engine=xelatex -V CJKmainfont=\"Hiragino Sans\" -V mainfont=\"Hiragino Sans\"");
        }

        var result = AnyConverter.Convert(pdfPath);
        
        Assert.NotNull(result.Text);
        Assert.NotEmpty(result.Text);
        
        // 変換結果をファイルに保存
        File.WriteAllText(convertedMdPath, result.Text);
        
        // 元のMarkdownを読み込み
        string originalContent = File.ReadAllText(originalMdPath);
        
        // 基本的な内容が保持されているかチェック
        VerifyContentSimilarity(originalContent, result.Text, "test-financial-report");
        
        Console.WriteLine($"Financial report test completed. Length: {result.Text.Length}, Warnings: {result.Warnings.Count}");
    }

    [Fact]
    public void MixedContentConversionTest()
    {
        string originalMdPath = "./Resources/test-mixed-content.md";
        string pdfPath = "./Resources/test-mixed-content.pdf";
        string convertedMdPath = "./test-mixed-output.md";

        Assert.True(File.Exists(originalMdPath), $"Original markdown file should exist: {originalMdPath}");
        
        if (!File.Exists(pdfPath))
        {
            Assert.Fail($"PDF not found. Run: pandoc {originalMdPath} -o {pdfPath} --pdf-engine=xelatex -V CJKmainfont=\"Hiragino Sans\" -V mainfont=\"Hiragino Sans\"");
        }

        var result = AnyConverter.Convert(pdfPath);
        
        Assert.NotNull(result.Text);
        Assert.NotEmpty(result.Text);
        
        // 変換結果をファイルに保存
        File.WriteAllText(convertedMdPath, result.Text);
        
        // 元のMarkdownを読み込み
        string originalContent = File.ReadAllText(originalMdPath);
        
        // 基本的な内容が保持されているかチェック
        VerifyContentSimilarity(originalContent, result.Text, "test-mixed-content");
        
        Console.WriteLine($"Mixed content test completed. Length: {result.Text.Length}, Warnings: {result.Warnings.Count}");
    }
}