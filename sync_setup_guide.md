# 固形石鹸ダッシュボード：クラウド同期 ＆ 共有セットアップ手順書

本ドキュメントは、ご夫婦で安全にダッシュボードのデータ（お気に入り、購入履歴、使用期間、購入金額、感想メモ、5つの評価軸）を共同編集・同期し、かつスマホやiPadからいつでも簡単にアクセスできるようにするためのセットアップ手順書です。

---

## 🛠️ ステップ1：データの保存先（スプレッドシート ＆ GAS）の作成

### 1. 空のスプレッドシートを作成する
1. ブラウザで [Google スプレッドシート](https://sheets.google.com/) にアクセスし、Googleアカウントでログインします。
2. 左上の **「空白」**（＋ マーク）をクリックして、新しいスプレッドシートを作成します。
3. 左上の「無題のスプレッドシート」をクリックし、分かりやすい名前（例：`石鹸プロジェクト同期用`）に変更します。
4. シート名は初期状態の「シート1」のまま、**中身は何も入力せず空っぽの状態で放置**します。

### 2. Google Apps Script (GAS) を開く
1. 作成したスプレッドシートの上部メニューの **「拡張機能」 ＞ 「Apps Script」** をクリックします。
2. 開いたエディタ画面のコード入力欄に、最初から書かれている `function myFunction() { ... }` などの文字を**すべて選択して完全に削除**し、空にします。

### 3. コードの貼り付けとパスワード設定
1. 以下の **「GASコピペ用ソースコード」** をすべてコピーして、GASエディタに貼り付けます。
2. **【必須設定】** コード of 1行目にある `var SECURITY_PASSWORD = "ここにパスワードを入力";` の部分を、ご夫婦共通の**「秘密のパスワード（合言葉）」**に書き換えてください。
   - 例：`var SECURITY_PASSWORD = "soapSecretKey123";`
3. エディタ上部の **「保存」** アイコン（フロッピーディスクのマーク）をクリックして保存します。

#### 【GASコピペ用ソースコード】
```javascript
// 【必須設定】ご夫婦で共有する秘密のパスワードを設定してください
var SECURITY_PASSWORD = "ここにパスワードを入力";

function doGet(e) {
  var password = e.parameter.password;
  if (password !== SECURITY_PASSWORD) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Unauthorized: Invalid password" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var result = {};
  
  // 1行目はヘッダー [ID, お気に入り, 購入済み, 購入日, 感想メモ, 使用開始日, 使用終了日, 購入金額, 香り評価, 消臭効果評価, 泡立ち評価, 乾燥評価, 肌触り評価]
  // 2行目以降がデータ
  for (var i = 1; i < data.length; i++) {
    var id = data[i][0];
    if (!id) continue;
    result[id] = {
      favorite: data[i][1] === true || data[i][1] === "TRUE" || data[i][1] === 1,
      purchased: data[i][2] === true || data[i][2] === "TRUE" || data[i][2] === 1,
      date: data[i][3] ? formatDate(data[i][3]) : "",
      notes: data[i][4] || "",
      startDate: data[i][5] ? formatDate(data[i][5]) : "",
      endDate: data[i][6] ? formatDate(data[i][6]) : "",
      price: data[i][7] !== undefined && data[i][7] !== "" ? Number(data[i][7]) : "",
      ratingScent: data[i][8] !== undefined && data[i][8] !== "" ? Number(data[i][8]) : 0,
      ratingDeodorant: data[i][9] !== undefined && data[i][9] !== "" ? Number(data[i][9]) : 0,
      ratingFoaming: data[i][10] !== undefined && data[i][10] !== "" ? Number(data[i][10]) : 0,
      ratingDryness: data[i][11] !== undefined && data[i][11] !== "" ? Number(data[i][11]) : 0,
      ratingTexture: data[i][12] !== undefined && data[i][12] !== "" ? Number(data[i][12]) : 0
    };
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // スプレッドシートが空、またはヘッダーしかない場合は初期化
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    sheet.getRange(1, 1, 1, 13).setValues([["ID", "お気に入り", "購入済み", "購入日", "感想メモ", "使用開始日", "使用終了日", "購入金額", "香り評価", "消臭効果評価", "泡立ち評価", "乾燥評価", "肌触り評価"]]);
    // 50行分を初期作成
    var initialValues = [];
    for (var id = 1; id <= 50; id++) {
      initialValues.push([id, false, false, "", "", "", "", "", 0, 0, 0, 0, 0]);
    }
    sheet.getRange(2, 1, 50, 13).setValues(initialValues);
  }
  
  var postData = JSON.parse(e.postData.contents);
  var password = postData.password;
  if (password !== SECURITY_PASSWORD) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Unauthorized: Invalid password" }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // データをシートから読み込み、IDと行番号のマップを作る
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  var idToRow = {};
  for (var i = 1; i < values.length; i++) {
    idToRow[values[i][0]] = i + 1; // 1-based index for row
  }
  
  // 送られてきたデータを書き込む
  var favorites = postData.favorites || [];
  var purchased = postData.purchased || {};
  
  // 1〜50のIDについて更新用の配列を作成
  for (var id = 1; id <= 50; id++) {
    var isFav = favorites.indexOf(id) !== -1 || favorites.indexOf(String(id)) !== -1;
    var pInfo = purchased[id] || purchased[String(id)];
    var isBought = !!pInfo;
    var date = pInfo ? pInfo.date : "";
    var notes = pInfo ? pInfo.notes : "";
    var startDate = pInfo ? pInfo.startDate : "";
    var endDate = pInfo ? pInfo.endDate : "";
    var price = pInfo && pInfo.price !== undefined && pInfo.price !== "" ? pInfo.price : "";
    var rScent = pInfo && pInfo.ratingScent !== undefined ? pInfo.ratingScent : 0;
    var rDeod = pInfo && pInfo.ratingDeodorant !== undefined ? pInfo.ratingDeodorant : 0;
    var rFoam = pInfo && pInfo.ratingFoaming !== undefined ? pInfo.ratingFoaming : 0;
    var rDry = pInfo && pInfo.ratingDryness !== undefined ? pInfo.ratingDryness : 0;
    var rText = pInfo && pInfo.ratingTexture !== undefined ? pInfo.ratingTexture : 0;
    
    // スプレッドシートの行を特定して更新 (B列からM列の計12列分に拡大)
    var row = idToRow[id];
    if (row) {
      sheet.getRange(row, 2, 1, 12).setValues([[isFav, isBought, date, notes, startDate, endDate, price, rScent, rDeod, rFoam, rDry, rText]]);
    } else {
      // 行がなければ追加
      sheet.appendRow([id, isFav, isBought, date, notes, startDate, endDate, price, rScent, rDeod, rFoam, rDry, rText]);
    }
  }
  
  return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function formatDate(dateObj) {
  if (dateObj instanceof Date) {
    var year = dateObj.getFullYear();
    var month = ("0" + (dateObj.getMonth() + 1)).slice(-2);
    var day = ("0" + dateObj.getDate()).slice(-2);
    return year + "-" + month + "-" + day;
  }
  return String(dateObj);
}
```

### 4. ウェブアプリとしてデプロイ（インターネット公開）する
1. GAS画面の右上にある青い **「デプロイ」** ボタンをクリックし、 **「新しいデプロイ」** を選択します。
2. 左上の歯車アイコン（種類の選択）をクリックし、 **「ウェブアプリ」** を選択します。
3. 設定項目を以下のように指定します：
   - **次のユーザーとして実行**: **「自分」**
   - **アクセスできるユーザー**: **「全員」**（CORS通信エラーを回避し、スマホから通信するために必須）
4. 青い **「デプロイ」** ボタンをクリックします。
5. 初回承認を求められたら以下の手順で行います：
   - **「アクセスを承認」** をクリックし、ご自身のGoogleアカウントを選択します。
   - 警告画面が出たら、左下のグレーの文字 **「詳細」 ＞ 「（安全ではないページ）に移動」** をクリックします。
   - アクセス権確認画面で、右下の **「許可」** をクリックします。
6. デプロイ完了画面に表示される **「ウェブアプリのURL」**（最後に `/exec` で終わるURL）をコピーして控えておきます。

---

## 🌐 ステップ2：画面のURL化（GitHub Pages）とブックマーク

### 1. GitHubリポジトリへのアップロード
1. [GitHub](https://github.com/) にブラウザでログインし、新しくパブリックリポジトリ（例：`soap-project`）を作成します。
2. Macの **「ターミナル」** アプリを開き、以下のコマンドを順番に実行します：
   ```bash
   cd /Users/kk/Documents/MyVault/Project/SoapProject
   git init
   git add .
   git commit -m "Initial commit for Soap Dashboard with Sync"
   git branch -M main
   git remote add origin https://github.com/[あなたのユーザー名]/[リポジトリ名].git
   git push -u origin main
   ```

### 2. GitHub Pagesの設定
1. ブラウザで作成したGitHubのリポジトリページを開き、上部の **「Settings」**（歯車マーク）タブをクリックします。
2. 左のサイドメニューから **「Pages」** をクリックします。
3. **Build and deployment** セクションの **Branch** を「None」から **「main」** に変更します。
4. 右のフォルダ選択は「/ (root)」のまま **「Save」** ボタンをクリックします。
5. 1〜2分後にリロードすると、画面最上部に公開されたWebサイトURLが表示されます。
   - URL例：`https://[あなたのユーザー名].github.io/[リポジトリ名]/dashboard.html`

### 3. 各端末のSafariでブックマークする
1. PC、スマホ、iPadのSafariで上記URLを開きます。
2. Safariの **「共有」** アイコン（上矢印の箱マーク）から、**「お気に入りに追加」** または **「ホーム画面に追加」** を選択して登録します。

---

## ☁️ ステップ3：ダッシュボードでの連携設定

1. 各端末のSafariで開いたダッシュボード画面上部の **「☁️ スプレッドシート同期設定」** アコーディオンをクリックして開きます。
2. 以下の2項目を入力します：
   - **GASウェブアプリURL**: 【ステップ1-4】でコピーした `.../exec` URL
   - **同期用パスワード（合言葉）**: 【ステップ1-3】でGASコードの1行目に設定したパスワード
3. **「設定を保存して同期」** をクリックします。
4. 「Googleスプレッドシートとの同期設定が成功しました！」と表示されれば、完了です。
5. 以降、ダッシュボードでお気に入りトグルや感想入力を行うと、バックグラウンドで自動的にスプレッドシート側にも反映されます。

---

## 💸 仕様詳細：1日あたりのコスト（コスパ）計算について

- **計算式**:
  - `使用日数 = (使用終了日 - 使用開始日) + 1` (開始日と終了日が同じ日でも「1日間」と計算します)
  - `1日あたりの金額 = 購入金額 / 使用日数`
- **小数点表示**: 
  - 小数点第2位まで正確に算出・表示します（例: `💸 コスパ: 25.33 円/日`）。
- **デフォルト値**:
  - 「✓買った記録」を有効にした瞬間、入力の手間を省くため、その製品の「カタログ価格」が「購入金額（税込）」に初期値として自動コピーされます（手動で自由に変更可能です）。
- **コスパ順ソート**:
  - 並び替えメニューから「コスパが良い順（1日あたり安価）」または「コスパが悪い順（1日あたり高価）」で並び替えることができます。
  - 使用日数と金額が両方入力され、コストが算出されている石鹸が優先してソートされ、未算出のものは自動的に最下部に整理されます。
