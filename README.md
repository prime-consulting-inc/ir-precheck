# ir-precheck

適時開示・決算短信・決算説明会資料・成長可能性資料を提出する前に、ファイルに残りやすい形式面のミスを点検するAIスキルです。PDF・Word・PowerPoint・Excelを読み、指摘を「要修正」「要確認」「情報」に分けて日本語で返します。

作成：株式会社プライムコンサルティング

> このスキルは機械で拾える項目だけを見ます。指摘がないことは、開示してよいことの確認にはなりません。

## 点検する項目

| 区分 | 主な内容 |
|---|---|
| 文書のプロパティ | タイトルの「（案）」「ドラフト」「v3」「Microsoft Word -」、作成者の個人名・他社名、自社以外の証券コード |
| コメント・注釈 | PDFの注釈、Word・PowerPoint・Excelのコメント、Wordの変更履歴 |
| 墨消し・伏せ字 | 図形を重ねただけの黒塗り、適用していない墨消し、「A社」などの伏せ表記と同じ位置に残った文字、図形の下に残った画像 |
| 見えない文字 | 白地の白文字、極小文字、ページの外の文字、Wordの隠し文字 |
| 社内向けの表記 | 社外秘・取扱注意・Confidentialなどの表記と透かし、「差し替えてください」「数字は後で」などのレビュー文言 |
| PowerPoint | 発表者ノート、非表示スライド、スライドの外の図形（表・グループ内の文字も対象） |
| PDFの設定 | パスワード、印刷・コピー禁止、埋め込みファイル、非表示レイヤー、未埋め込みフォント |
| 日付 | 日付と曜日のずれ（年のある日付）、作成日時が開示日より後・90日以上前 |

黒塗りや伏せ字の下にある文字は、場所と文字数だけを出し、文字列そのものは出しません。

## 出力の例

```
# IR提出前チェック結果（ir-precheck 1.0.0）

- 総合判定: 要修正あり

## sample.pdf

### 要修正（6件）

- [メタデータ] タイトル（Title）: タイトル（Title）に公開したくない表現があります（内部稿・案、バージョン痕跡、Word変換名）
  - 詳細: Microsoft Word - 説明会開催のお知らせ（案）_v3
- [メタデータ] 作成者（Author）: 作成者が個人名またはログイン名に見えます
- [注釈] 2ページ / Text: 注釈またはコメントが残っています
- [黒塗りの残存] 1ページ: 黒い図形の下に文字データが残っています（図形を重ねただけの黒塗り）。選択やコピーで読める状態です
  - 詳細: 4文字分
- [レビュー文言] 2ページ: 社内向けの指示やメモ（数字は後で）が残っています
- [日付と曜日] 1ページ: 日付と曜日が一致しません（2026年11月20日は金曜日）
```

AIから使うと、この結果に場所と直し方（Word・PowerPoint・Acrobatのどの操作で直すか）を添えて返します。

## 入れ方

### Claude Code などのエージェント（npx skills）

```bash
npx skills add prime-consulting-inc/ir-precheck
# 更新
npx skills update ir-precheck
```

### Claude Code のプラグイン

```text
/plugin marketplace add prime-consulting-inc/ir-precheck
/plugin install ir-precheck@ir-precheck
```

更新は `claude plugin marketplace update ir-precheck` のあと `claude plugin update ir-precheck@ir-precheck` を実行し、Claude Code を再起動します。

### claude.ai（Web版・デスクトップ版）

[リリース](https://github.com/prime-consulting-inc/ir-precheck/releases/latest)に添付した `ir-precheck.zip` を、claude.aiのスキルの追加画面からアップロードします。この `ir-precheck.zip` にはスキルのフォルダだけが入っています。リポジトリのトップページから取得したZIPには開発用のファイルも含まれるため、こちらを使ってください。

### Codex・GitHub Copilot

SKILL.md は共通の形式（Agent Skills）で書いているので、Codex と GitHub Copilot（Copilot CLI、VS Code のエージェントモード、cloud agent）でも使えます。

```bash
# Codex（作業するリポジトリの .agents/skills/ に入ります）
npx skills add prime-consulting-inc/ir-precheck -a codex

# GitHub Copilot（GitHub CLI 2.90.0 以降）
gh skill install prime-consulting-inc/ir-precheck ir-precheck
```

手で置く場合は `skills/ir-precheck/` フォルダごと、Codex なら `.agents/skills/`、Copilot なら `.github/skills/` の下にコピーします。どちらもスクリプトを動かすため、その環境に Python とライブラリが要ります。cloud agent などネットワークが制限された環境では、`pip install` ができずに「未検査」になることがあります。

### 必要なもの

Python 3.9以上と、`skills/ir-precheck/requirements.txt` のライブラリ（PyMuPDF、python-docx、python-pptx、openpyxl）が要ります。ライブラリがないときは、そのファイルが「未検査」と表示されます。

## 使い方

AIに次のように頼みます。

```text
提出前チェックをして。会社名は〇〇株式会社、コード1234、2027年3月期、開示日は11月13日。
（PDFと元のPowerPointを添付）
```

会社名・コード・決算期・開示日を書かなかった場合は、AIが表紙から読み取って確認します。

### スクリプトだけで使う

```bash
python3 skills/ir-precheck/scripts/ir_precheck.py \
  --company "〇〇株式会社" --code 1234 --period "2027年3月期" --date 2026-11-13 \
  提出予定.pdf 説明会資料.pptx
```

| オプション | 内容 |
|---|---|
| `--company` `--code` | 作成者の他社名、自社以外の証券コードの判定に使います |
| `--period` | 指定した決算期が本文にあるかを見ます |
| `--date` | 開示日。作成日時との比較に使います |
| `--ng-words` | 追加で探す語（支援会社名、社内の案件名など）。カンマ区切りかファイル |
| `--redact-words` | 墨消ししたはずの語が本文から取り出せないかを見ます。自社のファイルにだけ使ってください |
| `--json` | JSONで出力します |
| `--clean-metadata 出力.pdf` | 文書のプロパティを消した公開用コピーを作ります。原本は変えません |

終了コードは、0が指摘なしまたは情報のみ、1が要確認あり、2が要修正あり、3が検査できなかったファイル（パスワード付きのOfficeファイル、未対応の形式など）または引数の誤りありです。

## 使ってよい範囲

- 自社が提出・掲載する前のファイルの点検に使ってください。
- 他社が公開済みの資料から、伏せた情報を探す用途には使わないでください。スキルの指示でも、この依頼には応じないようにしています。
- スクリプトは外部と通信しません。ただし、claude.ai などクラウドで動くAIでスキルを使う場合、ファイルはそのAIサービスに送られます。社外に出せない段階のファイルは、各社のAI利用の決まりに従ってください。

## 点検しないもの

グラフの棒の長さと数値の一致、他社ロゴ、中期経営計画との整合、画像になった文字、年のない日付（「11月20日（木）」）の曜日は見ていません。[`manual_checklist.md`](./skills/ir-precheck/references/manual_checklist.md) で目視確認してください。わかっている見逃しと誤検知は [docs/known-issues.md](./docs/known-issues.md) にまとめています。

## 更新

検知する項目を絞った状態で公開し、誤検知の報告や新しい事例が届くたびに小さく更新します。変更は [CHANGELOG.md](./CHANGELOG.md) に書きます。

誤検知・見逃しは [Issues](https://github.com/prime-consulting-inc/ir-precheck/issues) で受け付けます。報告には実物のファイルを添付せず、どの区分で何が起きたかと、再現できる最小の例（自作のファイル）を書いてください。

## ライセンス

MIT License
