# ネストしたコンテンツテスト

## テーブル内リスト

| 機能 | 詳細 | 要件 |
|------|------|------|
| 認証 | - ログイン<br>- ログアウト<br>- パスワード変更 | 必須 |
| 管理 | 1. ユーザー管理<br>2. 権限管理<br>3. システム設定 | 必須 |
| レポート |  | オプション |

## コードブロックとテーブルの組み合わせ

```csharp
public class User
{
    public string Name { get; set; }
    public int Age { get; set; }
}
```

| メソッド | 説明 | 戻り値 |
|----------|------|--------|
| GetUser() | ユーザー情報を取得<br><br>```csharp<br>var user = GetUser();<br>``` | User |
| SaveUser() | ユーザー情報を保存 |  |