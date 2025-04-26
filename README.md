# Markdown Converter

ドキュメントファイルをマークダウン形式に変換します。


## How to use
```cs
using MarkdownConverter;

using FileStream stream = File.OpenRead("<filePath>.docx");
ConvertResult result = DocxConverter.Convert(stream);

Console.WriteLine(result.Text);
```

## How to build
```shell
dotnet pack
ls ./MarkdownConverter/bin/Release
```