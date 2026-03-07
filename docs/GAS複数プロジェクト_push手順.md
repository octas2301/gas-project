# 同じアカウントで複数GASプロジェクトを管理するときの push 手順

## 結論：**影響するのは「push を実行したフォルダ」のプロジェクトだけです**

- **clasp push** は、**今いるフォルダ（またはその配下）にある .clasp.json の scriptId** に対応する **1つのGASプロジェクト** にだけ反映します。
- 他のプロジェクト（別の scriptId）のコードは **一切書き換わりません**。
- 同じ Google アカウントで複数プロジェクトを持っていても、**どのフォルダで push するか** を間違えなければ、他プロジェクトに影響することはありません。

---

## ルール（覚えておくこと）

| やること | 実行するフォルダ | 更新されるプロジェクト |
|----------|------------------|--------------------------|
| メイン（Yahoo・コード.js）を反映 | `gas-project` | メインプロジェクト（.clasp.json の scriptId）のみ |
| 在庫管理アプリを反映 | `gas-project\inventory-app` | 在庫管理アプリのみ |
| 広告運用GASを反映 | 後述の「広告運用」のどちらか | 広告運用GASのみ |
| export-tax-list を反映 | `gas-project\export-tax-list` | export-tax-list のみ |

**必ず「反映したいプロジェクトのフォルダに cd してから」`clasp push` を実行してください。**

---

## フォルダとプロジェクトの対応（gas-project 内）

| フォルダ | 中身 | push するとき |
|----------|------|----------------|
| `gas-project` | コード.js, Yahoo.js（メイン） | `cd gas-project` → `clasp push` |
| `gas-project\inventory-app` | 在庫管理アプリ | `cd gas-project\inventory-app` → `clasp push` |
| `gas-project\広告運用GAS` | 広告運用GAS（編集用） | 下記「広告運用の場合」参照 |
| `gas-project\export-tax-list` |  export-tax-list | `cd gas-project\export-tax-list` → `clasp push` |

---

## 広告運用GASの場合（Desktop に clone がある構成）

現在の運用では、**編集は gas-project\広告運用GAS、反映（push）は Desktop\広告運用GAS_clone** で行っています。

- **gas-project の直下に .clasp.json がある**ため、`gas-project\広告運用GAS` で `clasp push` を実行すると、clasp が**親の .clasp.json（メインプロジェクト）を参照してしまう**ことがあり、広告運用ではなくメインプロジェクトに push されてしまいます。
- そのため **「コピー → Desktop の clone で push」** にすると、**親の .clasp.json の影響を受けず**、広告運用の scriptId だけに反映されます。

**広告運用だけを反映したいとき（他プロジェクトに影響させない）：**

```powershell
# 1. 最新を clone にコピー
Copy-Item -Path "C:\Users\takuy\Desktop\gas-project\広告運用GAS\*" -Destination "C:\Users\takuy\Desktop\広告運用GAS_clone\" -Recurse -Force

# 2. clone で push（ここで更新されるのは広告運用のプロジェクトだけ）
cd C:\Users\takuy\Desktop\広告運用GAS_clone
clasp push
```

- このとき更新されるのは **広告運用GAS_clone の .clasp.json に入っている scriptId のプロジェクトだけ**です。
- メイン（gas-project）や在庫管理アプリ（inventory-app）には一切反映されません。

---

## 在庫管理アプリ（inventory-app）の場合

- **inventory-app** は **gas-project のサブフォルダ**で、**自分専用の .clasp.json**（在庫管理用の scriptId）を持ちます。
- 親の gas-project の .clasp.json には **`"skipSubdirectories": true`** が入っているため、親で push しても inventory-app の中身は含まれません。
- **inventory-app だけを反映したいとき：**

```powershell
cd C:\Users\takuy\Desktop\gas-project\inventory-app
clasp push
```

- このとき更新されるのは **inventory-app の .clasp.json の scriptId（在庫管理アプリ）だけ**です。
- メイン・広告運用・export-tax-list には一切反映されません。

---

## 新しいサブプロジェクトを追加したとき（再発防止）

**gas-project のルートで `clasp push` すると、メインプロジェクト（AI出品ツール）にだけ反映されますが、.claspignore に書いていないフォルダはすべて push されてしまいます。**  
新しいGAS用フォルダ（例: 新しいアプリ用の my-app）を gas-project 内に作ったら、**必ずルートの `.claspignore` に `my-app/**` のように 1 行追加**すること。追加し忘れると、次にルートで push したときにメインプロジェクトにそのフォルダごと取り込まれてしまいます。

- 手順: ルートの **.claspignore** を開く → 「サブプロジェクトを除外」のブロックに **`フォルダ名/**`** を追加して保存。
- 詳細な対処（誤 push したときの戻し方）は [CLASP_PUSH_AND_MENU.md](CLASP_PUSH_AND_MENU.md) の「誤 push したときの対処」を参照。

---

## まとめ

| 質問 | 回答 |
|------|------|
| 同じアカウントで複数プロジェクトがあっても、他に影響せず push できる？ | **できます。** 更新されるのは「push を実行したフォルダの .clasp.json の scriptId」のプロジェクトだけです。 |
| 全て gas-project 内で管理しているが、広告運用だけ Desktop に clone がある | 広告運用を反映するときは **Desktop\広告運用GAS_clone で push**。在庫・メイン・export-tax-list は **それぞれのフォルダで push** すれば、そのプロジェクトだけが更新されます。 |
| 安全に push するには？ | **反映したいプロジェクトのフォルダに cd してから `clasp push` する。** どのプロジェクトが更新されるかは、そのフォルダの .clasp.json の scriptId で決まります。 |
| 新しいサブプロジェクトを追加したら？ | **ルートの .claspignore に `フォルダ名/**` を必ず追加する。** 忘れるとルートで push したときにメインに誤取り込みされる。 |
