# ShareLinkCleaner365

SharePoint Online の共有リンクを管理・クリーンアップするためのPowerShellスクリプト集です。

## 機能

- 共有リンクの一括削除
- 匿名アクセスリンクの削除
- 期限切れリンクの削除
- 共有リンク状況のレポート作成
- リンクの有効性チェック

## 必要要件

- PowerShell 5.1以上
- Microsoft.Graph.Authentication モジュール
- SharePoint管理者権限

## インストール

1. このリポジトリをクローン：
```powershell
git clone https://github.com/yourusername/ShareLinkCleaner365.git
cd ShareLinkCleaner365
```

2. Microsoft Graph認証モジュールのインストール：
```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```

## 使用方法

### 共有リンクの一括削除

すべての共有リンクを削除します：
```powershell
.\src\Remove-SPSharingLinks.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/project"
```

オプション：
- `-WhatIf`: 削除をシミュレーションのみ実行
- `-Force`: 確認プロンプトをスキップ
- `-LogPath`: ログファイルの出力先を指定

### 匿名アクセスリンクの削除

匿名アクセス（Anyone）リンクのみを削除：
```powershell
.\src\Remove-SPAnonymousLinks.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/project"
```

### 期限切れリンクの削除

期限切れリンクを削除：
```powershell
.\src\Remove-SPExpiredLinks.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/project" -GracePeriodDays 7
```

オプション：
- `-GracePeriodDays`: 期限切れ後の猶予期間（日数）

### 共有リンク状況のレポート作成

現在の共有リンク状況をレポート：
```powershell
.\src\Get-SPSharingLinkReport.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/project"
```

オプション：
- `-Format`: 出力形式（CSV/JSON）
- `-IncludeDetails`: 詳細情報を含める
- `-GroupByType`: タイプごとにグループ化

### リンクの有効性チェック

共有リンクの有効性をテスト：
```powershell
.\src\Test-SPSharingLinkStatus.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/project"
```

オプション：
- `-OnlyInvalid`: 無効なリンクのみ表示
- `-IncludePermissions`: 権限情報を含める

## スクリプトの説明

1. `Remove-SPSharingLinks.ps1`
   - すべての共有リンクを対象に削除を実行する総合コマンド
   - サイト全体のクリーンアップに最適

2. `Remove-SPAnonymousLinks.ps1`
   - 匿名アクセス（Anyone）リンクに限定して削除
   - セキュリティ強化時の使用に最適

3. `Remove-SPExpiredLinks.ps1`
   - 有効期限切れのリンクを削除
   - 定期的なメンテナンス時に使用

4. `Get-SPSharingLinkReport.ps1`
   - 共有リンクの状況を一覧出力
   - CSV/JSONフォーマットでレポート作成可能

5. `Test-SPSharingLinkStatus.ps1`
   - 共有リンクの有効性をチェック
   - 問題のあるリンクを特定
   - 削除前の事前確認に最適

## 注意事項

- 実行前に必ず `-WhatIf` オプションでシミュレーション実行することを推奨
- 重要な操作には確認プロンプトが表示されます
- すべての操作はログファイルに記録されます
- 大規模なサイトでの実行時は処理に時間がかかる可能性があります

## ライセンス

MIT License

## 作者

Hisaho Nakata <nahisaho@microsoft.com>
