# 広告運用GAS を PowerShell（clasp）で反映する手順

## 原因

- **広告運用GAS** フォルダは **gas-project の内側**（`gas-project\広告運用GAS`）にあります。
- clasp は **今いるフォルダより上** に `.clasp.json` があると、そちらを「プロジェクト」とみなすことがあります。
- そのため **`gas-project\広告運用GAS` で `clasp push`** を実行すると、**「Project file already exists」** などになり、広告運用用のスクリプトに反映されません。
- **Desktop の 広告運用GAS_clone** で push すると反映されますが、**中身が古い**と「Script is already up to date」のままになります（編集しているのは gas-project\広告運用GAS 側だから）。

---

## 対策：コピーしてから push する

**流れ**  
1. 編集は **gas-project\広告運用GAS** で行う（いつも通り）。  
2. 反映するときだけ **中身を 広告運用GAS_clone にコピー** する。  
3. **広告運用GAS_clone** で **clasp push** する。

こうすると、clasp は「親に .clasp.json のないフォルダ」で動くので、正しく広告運用のスクリプトに push できます。

---

## 手順（毎回）

PowerShell で次を **順に** 実行してください。

```powershell
# 1. gas-project\広告運用GAS の最新を、Desktop の clone フォルダにコピー
Copy-Item -Path "C:\Users\takuy\Desktop\gas-project\広告運用GAS\*" -Destination "C:\Users\takuy\Desktop\広告運用GAS_clone\" -Recurse -Force

# 2. clone フォルダに移動して push
cd C:\Users\takuy\Desktop\広告運用GAS_clone
clasp push
```

- **「Updated 2 files.» など** と出れば、広告運用の Apps Script に反映されています。
- **「Script is already up to date.»** のままなら、clone 側がすでに同じ内容だったか、コピーがうまくいっていない可能性があります。1. のコピーを再度実行してから 2. を試してください。

---

## 広告運用GAS_clone がない場合

Desktop に **広告運用GAS_clone** を消してしまった場合は、もう一度 clone してから上記を使います。

```powershell
cd C:\Users\takuy\Desktop
mkdir 広告運用GAS_clone -ErrorAction SilentlyContinue
cd 広告運用GAS_clone
clasp clone "1bxBU-4n5kdA5DFWkcRQFFfCYAWaMaNdECa2fSxjMcobsCrMqG5uHDrnw"
```

その後、**「手順（毎回）」** の 1. 2. で、gas-project の内容をコピーして push してください。

---

## まとめ

| やりたいこと | 実行する場所 | コマンド |
|--------------|----------------|----------|
| 広告運用の**編集** | gas-project\広告運用GAS（エディタで編集） | — |
| 広告運用の**反映（push）** | **Desktop\広告運用GAS_clone** | 先に gas-project\広告運用GAS をコピー → `clasp push` |

**反映できない原因**は「gas-project の内側で push している」ことです。**対策**は「一度 Desktop の clone にコピーして、そこで push する」です。
