# 個人情報変換ちゃん macOS版 ビルド手順

この一式は、macOS上で「個人情報変換ちゃん」の `.app` と `.dmg` を生成します。

## 対応環境

- macOS 13以降を推奨
- Apple Silicon（M1以降）またはIntel Mac
- Python 3.11以降
- インターネット接続（初回の依存ライブラリ取得時のみ）

## ビルド

1. このフォルダをMacへコピーします。
2. ターミナルを開き、このフォルダへ移動します。
3. 次を実行します。

   ```bash
   bash build_macos.sh
   ```

4. 完了後、次の2つが作られます。

   - `dist/個人情報変換ちゃん.app`
   - `dist/個人情報変換ちゃん_Mac.dmg`

ビルドしたMacと同じCPUアーキテクチャ向けに生成されます。Apple Silicon Macでビルドすればarm64版、Intel Macでビルドすればx86_64版になります。

## 配布について

ビルドスクリプトが作るアプリはアドホック署名です。社内テストには使用できますが、一般配布でGatekeeperの警告を避けるには、Apple Developer ProgramのDeveloper ID証明書による署名とApple公証が必要です。

## 機能互換性

Windows版と同じ `.xlsx`、`.html`、`.epbmap` 形式を使用します。Windows版で作成した対応表もMac版で利用できます。
