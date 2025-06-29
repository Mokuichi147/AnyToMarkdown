# Any To Markdown

任意ファイルをマークダウン形式に変換します。

## テスト状況

**現在のテスト結果**: 11個のテストが失敗、2個が合格 (合格率: 15.4%)

主な課題:
- PDF→Markdown変換の構造解析精度が不十分
- テーブル構造の保持率が77.8%で目標100%に未達
- 財務レポートや科学論文等の複雑な文書で22.2-33.3%の低い合格率
- 日本語文書の処理精度に改善が必要

## Supported document types

| type | text | image | table | bold | italic |
| --- | :-: | :-: | :-: | :-: | :-: |
| docx | ✅ | ✅ | ✅ | ✅ | ❌ |
| pdf (text) | ✅ | ✅ | 🔧 | ❌ | ❌ |

🔧 = 開発中・精度向上が必要

## How to use
```cs
using AnyToMarkdown;

using FileStream stream = File.OpenRead("<filePath>");
ConvertResult result = AnyConverter.Convert(stream);

// マークダウン形式のテキスト
Console.WriteLine(result.Text);
```

## How to build
```shell
dotnet pack
ls ./AnyToMarkdown/bin/Release
```