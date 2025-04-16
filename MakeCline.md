Microsoft 365 SharePoint の共有リンクを操作するプログラムをPowerShell, Microsoft Graph API で作成

次の5つのスクリプトを作成。

Remove-SPSharingLinks.ps1
　SharePoint のすべての共有リンク（匿名、期限切れ、無効なものなど）を対象に削除する総合コマンド。

Remove-SPAnonymousLinks.ps1
　匿名アクセス（Anyoneリンク）に限定して削除する、安全強化向け。

Remove-SPExpiredLinks.ps1
　有効期限切れのリンクをピンポイントで削除する用途に最適。

Get-SPSharingLinkReport.ps1
　各サイトまたはテナント内の共有リンク状況を一覧出力（CSV/JSONログ向けにも最適）。

Test-SPSharingLinkStatus.ps1
　共有リンクの有効性やアクセス状態をチェックする確認用コマンド。削除前のドライランに使える。