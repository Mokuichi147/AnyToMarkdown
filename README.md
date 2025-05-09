# Any To Markdown

任意ファイルをマークダウン形式に変換します。


## Supported document types

| type | text | image | table | bold | italic |
| --- | :-: | :-: | :-: | :-: | :-: |
| docx | ✅ | ✅ | ✅ | ✅ | ❌ |
| pdf (text) | ✅ | ✅ | ✅ | ❌ | ❌ |


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