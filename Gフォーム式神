# 🗺️ KML式神 (KML Shikigami) v24

**Gemini (AI) × Google Apps Script (GAS)** AIの力で、テキスト・画像・URLからGoogle Earth/MyMaps用の地図データ（KML）を全自動生成するシステム。

![Version](https://img.shields.io/badge/version-24.1-blueviolet)
[![License: CC BY 4.0](https://img.shields.io/badge/License-CC%20BY%204.0-lightgrey.svg)](https://creativecommons.org/licenses/by/4.0/deed.ja)
![Platform](https://img.shields.io/badge/platform-Google_Apps_Script-blue)

> 📖 **詳しい使い方は Note 記事をご覧ください:**
> [【完成】API不要でGoogleマイマップを全自動生成する「KML式神」を召喚してみた（Gemini × GAS）](https://note.com/_cop/n/n6c996826de5b)

---

## 📖 概要 (Overview)

「KML式神」は、Google Geminiのマルチモーダル解析能力と、GASによるWebアプリケーションを組み合わせた地図作成支援ツールです。

住所検索による位置ズレ（ハルシネーション）を極限まで排除するプロンプト設計と、人間が直感的に修正できる専用GUI（Leaflet.js搭載）により、**「AIに作らせて、人間が仕上げる」** という最強のワークフローを提供します。

### ✨ 主な機能
* **マルチモーダル入力:** テキストだけでなく、パンフレットの写真や、まとめサイトのURLからも地図を生成。
* **絶対座標特定 (Absolute Zero):** 住所検索を禁止し、POI検索とWGS84座標系を厳守させることで、山中へのピンズレを防ぐ。
* **リアルタイム同期GUI:** 地図上のピンをドラッグするだけで、裏側の座標データが即座に書き換わる直感的な修正画面。
* **ビジュアルアイコンピッカー:** Google Earth標準のアイコン（50種以上）を視覚的に選択・一括変換可能。

---

## 🛠️ セットアップ (Installation)

このシステムは Google Apps Script (GAS) 上で動作します。

### 1. Googleスプレッドシートの準備
1.  新規スプレッドシートを作成します。
2.  以下の3つのシートを作成し、1行目にヘッダーを設定します。

| シート名 | A1 | B1 | C1 | D1 | E1 | F1 | G1 |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| `generate_log` | map_id | 作成日時 | 地図タイトル | 説明文 | ステータス | KML URL | ダウンロードURL |
| `places_data` | map_id | スポット名 | 説明 | 緯度 (lat) | 経度 (lng) | アイコンタイプ | (空欄) |
| `config` | 項目名 | 値 (設定値) | (空欄) | (空欄) | (空欄) | (空欄) | (空欄) |

3.  `config` シートの A2に「GemURL」、B2にあなたのGeminiプロンプトの共有リンクを入力します。（任意）

### 2. スクリプトの導入
1.  スプレッドシートのメニューから **[拡張機能] > [Apps Script]** を開きます。
2.  **`Code.gs`**: リポジトリ内の `Code.gs` の内容をコピペします。
3.  **`index.html`**: スクリプトエディタでHTMLファイルを作成し、名前を `index` にして、リポジトリ内の `index.html` の内容をコピペします。

### 3. デプロイ
1.  GASエディタ右上の **[デプロイ] > [新しいデプロイ]** をクリック。
2.  **種類の選択**: 「ウェブアプリ」
3.  **次のユーザーとして実行**: 「自分」
4.  **アクセスできるユーザー**: 「全員」（または組織内）
5.  デプロイ後に発行される **ウェブアプリURL** にアクセスすれば完了です。

---

## 🧠 AIプロンプト (Gemini System Instruction)

Geminiに「KMLアーキテクト」としての役割を与えるための指示書です。
以下のファイルをGemini（またはGem）の System Instructions に設定してください。

* 📄 **[KML式神(Gem).md](./KML式神(Gem).md)**

---

## 🎮 使い方 (Usage)

1.  **GeminiでJSON生成:**
    * Geminiに「この画像の場所を地図にして」「〇〇県のうどん屋リストを作って」と指示し、JSONを出力させます。
2.  **アプリへ入力:**
    * Webアプリを開き、JSONを貼り付けると、即座に地図上にピンがプロットされます。
3.  **修正 (Human-in-the-loop):**
    * **位置:** 航空写真モードで確認し、ズレていればピンをドラッグして屋根の上に合わせます。
    * **見た目:** アイコンをクリックして色や形を変更します。
4.  **KML出力:**
    * 「KMLデータを構築」ボタンを押すと、Googleドライブに `.kml` ファイルが保存されます。
5.  **活用:**
    * GoogleマイマップやGoogle Earthにインポートして利用します。

---

## 🏗️ 技術スタック (Tech Stack)

* **Backend:** Google Apps Script (GAS)
* **Frontend:** HTML5, CSS3 (Cyber-Violet UI), JavaScript (Vanilla)
* **Map Library:** Leaflet.js
* **Map Tiles:**
    * OpenStreetMap
    * GSI Tiles (国土地理院・電子国土Web)
* **AI:** Google Gemini 1.5 Pro (Recommended)

---

## 📜 ライセンス (License)

この作品は **クリエイティブ・コモンズ 表示 4.0 国際 ライセンス (CC BY 4.0)** の下に提供されています。

[![License: CC BY 4.0](https://i.creativecommons.org/l/by/4.0/88x31.png)](https://creativecommons.org/licenses/by/4.0/deed.ja)

---
*Created by t-chi & The World's Best AI Engineer.*
