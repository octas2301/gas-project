# 変更台帳（復元用メモ）

コード全文ではなく **対象・目的・戻し方** のみ記録する（[AGENT_HANDOVER.md](AGENT_HANDOVER.md) §8・§9.1）。

| 2026-11-15 | `vision_judge`／`prompt_fix_from_vision`／PLATING §3.1 | **C8言語化強化**: 具体差・keep・nextFix必須。抽象「別物禁止」だけの回し禁止。切戻し Git。clasp不要 | Git revert |
| 2026-11-15 | 外側フロー／PLATING §3.1 | **②一片FAIL→③深堀で補完**（②で止めない）。1枚直列は維持。切戻し Git。clasp不要 | Git revert |
| 2026-11-15 | C8／外側フロー／PLATING §3.1 | **1枚直列**: 1枚→Vision→不合格理由+合格側言語化→1枚再生成。3並行禁止。`passedFiles`明示。切戻し Git。clasp不要 | Git revert |
| 2026-11-15 | `vision_judge_piece_effect`／外側フロー／PLATING §3.1 | **C8合否=Visionで画像を見る**→FAILなら打ち消しlow再生成。数値のみ合否禁止。切戻し Git。clasp不要 | Git revert |
| 2026-11-15 | `PLATING` §3.1／`effect_verify_max_attempts` | **C8上限既定10**/スロット（旧3）。env `B5_EFFECT_VERIFY_MAX`。切戻し Git。clasp不要 | Git revert |
| 2026-11-15 | `PLATING` §3.1／`b5_generic_plating.run_effect_verify_loop`／外側フロー／PHASE | **C8効き確認ループ必須**: lowで禁止・危険文オフの効きを見てから先へ。スキップ禁止。上限 `B5_EFFECT_VERIFY_MAX`。切戻し Git。clasp不要 | Git revert |
| 2026-11-15 | `PLATING_AND_PIECE_GENERIC_LOGIC`／`b5_generic_plating.recommend_quality`／PHASE／HANDOVER | **合否§3を正・確定**: 商品非依存の型のみ。筋・血合では medium に上げない。一件禁止を共通に焼かない。切戻し Git。clasp不要 | Git revert |
| 2026-11-15 | `b5_generic_plating`／`b5_piece_truth_rules` | **SURFACE試験産物除外**: reinhard等を同送しない。切戻し Git。clasp不要 | Git revert |
| 2026-11-15 | `estimate_paste_match`／`estimate_cost_from_paste`／検品 build・apply_judge／見積検品要件 | **貼付→仕入を見送り＋名前部分一致可**。接合一致度を判断列・アラート色。clasp不要 | Git revert／`--no-name-partial` |
| 2026-11-15 | `estimate_inspection_build_rows.py` | **1商品1行**: 種を貼付名／単品名に集約。ASINタイトル分割をやめる。clasp不要 | Git revert |
| 2026-11-15 | `PurchaseResearch.js` D1 | **D1段階ログ**: cand後の exist／Keepa flat／GET／enrich／classify／書込と flat_read 進捗。最終 phase で切断点特定。要① clasp | Git revert（prD1PhaseLog_ 等） |
| 2026-11-15 | `PurchaseResearch.js` D1 cand | **門落ち込み固まり直し**: 調査カテゴリ外の落ちは倉庫読込前に除外。D1 RUNNINGログ追加。要① clasp | Git revert |
| 2026-11-15 | `PurchaseResearch.js` D1／再門／見積検品 build_rows／FLOW・検品要件 | **オリエンタル第1**: メーカー指定再門→D1門落ち込み（調査カテゴリONのみ・落ちはGETしない）→D2。検品は `--makers --include-skip --allow-no-cost`。要① clasp | Prop OFF／メニュー使わない／Git |
| 2026-11-15 | `b5_piece_truth_rules.write_piece_set_source`／`b5_piece_set_from_truth` | **activeSeed保持**: ゲート選定を一片生成で消さない。解決種を activeSeed に書く。切戻し Git。clasp不要 | Git revert |
| 2026-11-15 | `b5_piece_set_from_truth`／`b5_generic_plating`／`b5_piece_truth_rules`／`dry_sub`／F3／docs | **抜け配線**: 一片へSURFACE+面クロップ+核。マクロ除外。SURFACE+一片でSHAPE省略。F3 quality=recommend。切戻し `B5_GENERIC_PLATING_CORE=0`。clasp不要 | env 0／Git revert |
| 2026-11-15 | `b5_generic_plating.py`／`dry_sub`／`b5_main_base_compose`／`food_unit_canon`／F3／docs | **盛り共通核**: MAIN/SUB同じ型（危険文オフ・面クロップ・核プロンプト）。qualityは運用§3。切戻し `B5_GENERIC_PLATING_CORE=0`。clasp不要 | env 0／Git revert |
| 2026-11-15 | `PurchaseResearch.js` タジマヤ寸法 | **対象拡大**: FBA空に加え保管空・サイズ空も。既存FBAは消さない。メニュー名「FBA空・保管空」。要① clasp | 旧FBA空のみ条件へ戻す |
| 2026-11-15 | `estimate_list_fees_fill.py`／品番リスト BV・BW | **空の販手・保管を設定マスタ料率で補完**（カテゴリ空・寸法空は書かない。既存値非上書き）。販手1305／保管1155。寸法欠719はタジマヤ等。clasp不要 | 書いたBV/BWを消す |
| 2026-11-15 | ①/出品セット数量・品番リスト希望式／FLOW・見積検品要件 | **袋×ケース**（①＋出品同式）。希望＝販手+FBA+保管+目標利益込み（欠ける行は空）。Q3件修正。要①・出品 clasp | Q戻し／AQ旧式 `AO-AP-AN` |
| 2026-11-15 | 見積検品／`apply_judge`／`kentou_asin`／build_rows／要件 | **F判断列**・手数料不明は見送りせずアラート色・検討ASINの月販/順位をKeepa転記（空「-」固定を廃止）。clasp不要 | 判断列削除／可否を旧見送りへ |
| 2026-11-15 | 見積検品列／`estimate_inspection_cols_quote_diff`／build_rows／要件 | **G問屋見積・J差分（希望−見積・▲赤字・注意印）**。対象セットをNへ。石原MIX複数ASIN＋リンク復活。検討の別ASIN追記。clasp不要 | 列削除／N列戻し |
| 2026-11-15 | `estimate_inspection_build_rows.py`／見積検品 L | **対象セットにASINリンク**: 行末にASIN表示＋textFormatRuns。`--relink`で既存12行更新。石原MIX行は未（壊さない）。clasp不要 | L列を旧文に戻す |
| 2026-11-14 | `estimate_inspection_build_rows.py`／`estimate_inspection_main_images.py`／見積検品／FLOW | **他メーカー12行を見積検品へ回収**: 判断行＋画像／希望／提示。可=くらこん2。石原5は維持。clasp不要 | 追加12行削除 |
| 2026-11-14 | `estimate_cost_from_paste.py`／品番リスト AW／FLOW §3 | **貼付→1点仕入れ額（税込み）第1弾**: 見積候補×JAN接合×仕入空の12件。税込=floor(税抜×(1+税率))。上書きしない。clasp不要 | 当該12セルを空に戻す |
| 2026-11-14 | `b5_slot_lang.py`／`dry_sub`／`b5_main_base_compose`／F3 | **スロット言語化セット**: SURFACE Vision言語化保存／SHAPEはドナー形fields／S1合成の正／WRAP×一片相対サイズ。MAIN+SUB。切戻し `B5_SLOT_LANG=0`／`B5_S1_OMIT_OWN_PIECE=0`。clasp不要 | env／Git revert |
| 2026-11-14 | `b5_piece_deepen`／`b5_piece_truth_rules`／`dry_sub`／F3 | **S1添付＝仮想C**: SURFACE真の面のみ（face_crop埋め既定OFF）＋SHAPE＋IDENTITY＋袋。自SKU一片本体・一片言語化は同送しない。切戻し `B5_S1_OMIT_OWN_PIECE=0`／`B5_SURFACE_FACE_CROP=1`／`B5_SURFACE_SHAPE_SPLIT=0`。clasp不要 | env／Git revert |
| 2026-11-14 | `food_unit_canon.py`／F3 | **IDENTITY配置のみ**: fields優先同送。`sourcePromptBlock`の個品縁・表面節は除去。一片もfields優先。切戻し `B5_IDENTITY_PLATING_FIELDS_ONLY=0`。clasp不要 | env 0／Git revert |
| 2026-11-14 | `b5_piece_truth_rules`／`dry_sub`／F3 | **SURFACE本線同送（切り分け）**: 面アップ必須／無ければ一片面クロップ。SHAPEは形だけ（縁質感は借りない）。IDENTITYは別行で掃除。clasp不要 | `B5_SURFACE_SHAPE_SPLIT=0` |
| 2026-11-13 | ぶり `piece_set_source`／F3 `_tmp_piece_lock_prod_compare` | **F3ぶり着手**: HUMANへ同期で truth gate PASS。F3 SUBへ形状ドナー同送。人間検収待ち（MAIN2+SUB1）。clasp不要 | Git／virtfit退避 |
| 2026-11-13 | `b5_piece_truth_rules.py`／`b5_piece_deepen.py`／`dry_sub_overview_copy_test.py`／`b5_piece_gate.py`／B45 T4 | **形状ドナー自動**: 同ファミリー別 theme の piece_set を機械選択→深堀参考／SUB／1枚作り直しへ SHAPE_*。切戻し `B5_SURFACE_SHAPE_SPLIT=0`。clasp不要 | env 0／Git revert |
| 2026-11-13 | `b5_piece_deepen.py`／`sub_image_food_canon_gate.py`／B45 T4 | **深堀スタック自動**: タグONで参考・前回を自動構成。remakeは下段のみ・本線は前回のまま。下段採用API。切戻し `B5_PIECE_DEEPEN_UI=0`。clasp不要 | env 0／Git revert |
| 2026-11-13 | `b5_piece_deepen.py`／`sub_image_food_canon_gate.py`／`food_unit_canon.py`／B45 T4 | **深堀パネル第1**: 不合格タグで作業付近に大きめ手順。API `/api/piece-deepen`。役割分離文を lock 同送。切戻し `B5_PIECE_DEEPEN_UI=0`／`B5_SURFACE_SHAPE_SPLIT=0`。clasp不要 | env 0／Git revert |
| 2026-11-13 | 見積検品 F–I／要件 | **画像・価格を仕入可否直後へ移動**: F画像 G希望 H提示 I交渉。旧Q–T廃止。clasp不要 | 列順戻し |
| 2026-11-13 | 見積検品／`estimate_inspection_main_images.py`／FLOW／見積検品要件 | **提示価格二段**: 問屋提示。セット利益率≥30%なら上乗せ（着地29.9%）。200円下限優先。単品は提示=希望。列はF–I。clasp不要 | 列クリア／Git revert |
| 2026-11-13 | FLOW §5.3.2／見積検品要件／`見積検品` R | **希望価格の理由を正本化＋再算出**: セット `max(最安×15%,200円)`。200円＝発送梱包≈80円の上、15%＝定期5%＋広告8%。R列を品番リスト税抜から再取得。clasp不要 | docs戻し／再実行 |
| 2026-11-13 | ① `見積検品` Q–S／`estimate_inspection_main_images.py`／見積検品要件 | **判断表画像をKeepaフル正メインへ修正＋希望価格・値下げ交渉列**: 品番リスト画像は使わない。R=目標仕入額税抜、S=人記入。clasp不要 | 列Q–Sクリア／Git revert |
| 2026-11-13 | `purchase_research/PurchaseResearch.js`／FLOW §3 D | **タジマヤ二重実行ロック**: `PR_TAJIMAYA_DIM_LOCK_*`。busyは `skip=busy` で終了・続きは持っている実行だけ予約。要① clasp | Prop削除／Git revert |
| 2026-11-12 | `purchase_research/PurchaseResearch.js`／V1 §10／HUMAN_RUN | **タジマヤ失敗JAN印＋取得リトライ**: `PR_TAJIMAYA_MISS_JSON` で再チェック浪費を止める。HTTP等は6hクールダウン。続きは `prRunTajimayaDimFillFromTrigger`＋①ログに arm。要① clasp | Prop削除／Git revert |
| 2026-11-12 | `purchase_research/PurchaseResearch.js`／V1 §10／FLOW §3 D | **タジマヤ寸法補完**: 梱包三辺空＋JANのみ。ピース／ケース／格子→FBA。既存サイズ非上書き。メニュー＋D1。切戻し `PR_TAJIMAYA_DIM_FILL_ENABLED=false`。要① clasp | Prop false／Git revert |
| 2026-11-11 | `purchase_research/PurchaseResearch.js`／CYCLE §7.5／HUMAN_RUN／`_tmp_virtual_liveness_arm_order.py` | **沈黙復帰で洗い先武装**: `hard_reset` 後に洗い→C4。C4 例外は握りつぶす。旧順は C4 重走査で洗い未武装の無限沈黙（実測 11/06〜）。要① clasp | Git revert |
| 2026-11-10 | `sub_image_food_canon_gate.py` | **一片表示一本化**: ③は piece_set のみ（③b二重廃止）。HUMANは未反映貼付時だけ。clasp不要 | Git revert |
| 2026-11-10 | `b5_piece_gate.py`／`sub_image_food_canon_gate.py` | **一片作り直し・現行ロック切替**: 既定オンで現行写真を1枚目ロック／オフで種ベース。要望は差分。clasp不要 | Git revert |
| 2026-11-10 | `sub_image_food_canon_gate.py`／`food_unit_canon.py`／B45 | **IDENTITY盛り二段言語化**: ①元画像言語化＝本線正／②生成後＝生成指示。機械言語化API。clasp不要 | Git revert |
| 2026-11-11 | `competitor_jitta/CatalogEnum.js`／JITTA／README／`_tmp_virtual_queue_purge_cap.py` | **Catalog空転止め**: 未済≥上限なら API前に `batch_skip`（キューあり途中列挙のみ続行）。時計は間隔維持。仮想 PASS。要競合 clasp | Git revert |
| 2026-11-10 | `competitor_jitta/CatalogEnum.js`／JITTA／CYCLE／`_tmp_virtual_queue_purge_cap.py` | **幽霊規模帯禁止**: 追記0で band 非書・`catalog_pending_cap`／既存 via=comp＋キュー0を `catalog_phantom_heal`。仮想 PASS。要競合 clasp | Git revert |
| 2026-11-10 | `purchase_research/PurchaseResearch.js`／`competitor_jitta/CatalogEnum.js`／CYCLE §6.7.4bis／HUMAN §10.3non／FULL／`_tmp_virtual_queue_purge_cap.py` | **洗いキュー Keepa済即削除＋未済5万蛇口**: 完走時親パージ／メニューは終わり状態＋日数／新規列挙は未済≥50kで停止（途中社は続行）。選定は既存1回読み。仮想 ALL PASS。要①＋競合 clasp。初回一括は心拍停止→メニュー削除 | Prop `PR_WASH_QUEUE_PURGE_DONE=false`／`PR_WASH_QUEUE_PENDING_ROWS_MAX=0`／Git revert／backup `_local_backup/pre_queue_purge_cap_20261110_084822` |
| 2026-11-10 | `sub_image_food_canon_gate.py` | **見本下に言語化表示＋人間保存**: 一片／IDENTITY盛りの textarea・`/api/piece-lang-save`・`/api/identity-plating-lang-save`。すり合わせ用。clasp不要 | Git revert |
| 2026-11-10 | `food_unit_canon.py`／`dry_sub_overview_copy_test.py`／`_tmp_food_piece_plate_identity_lang_virt.py`／B45 | **IDENTITY盛り言語強化＋SUB移植ロック**: 観測項目追加・`identity_plating_lock_prompt_block`・SUB本線同送。ぶりで再言語化＋T06再生成。clasp不要 | Git revert／旧 `identity_plating_lang_prev_*` |
| 2026-11-10 | B45 §4.5.10bis-T2 | **不合格後言語化**: 個品／個包装ずれ→活性枚再読→ずれた面だけ作り直し。袋・文言・構図のみは不触。clasp不要 | docs revert |
| 2026-11-09 | `b5_piece_truth_rules.py`／`dry_b5_piece_truth_virt.py`／`food_unit_canon.py`／`sub_image_food_canon_gate.py`／B45 §4.5.10bis-T2 | **IDENTITY≠一片**: 同一バイト不合格・一片ロック場面は配置のみ＋料理具サイズ強制。切戻し `B5_PIECE_TRUTH_GATE=0`。clasp不要 | env 0／Git revert |
| 2026-11-08 | `b5_piece_gate.py`／`sub_image_food_canon_gate.py`／`b5_food_canon_status.py`／`food_unit_canon.py`／B45 §4.5.10bis-T3 | **見本ゲート一片レーン**: 有無・貼付・レ点→PIECE_TRUTH_HUMAN／piece_set。切戻し `B5_PIECE_GATE_UI=0`。clasp不要 | env 0／Git revert |
| 2026-11-08 | `b5_piece_truth_rules.py`／`dry_b5_piece_truth_virt.py`／`b5_piece_set_from_truth.py`／`b5_main_base_compose.py`／B45 §4.5.10bis-T2 | **実物由来F2**: human>wrap>virt・cube暴走禁止・source記録。切戻し `B5_PIECE_TRUTH_GATE=0`。clasp不要 | env 0／Git revert |
| 2026-11-08 | `food_unit_canon.py`／`b5_main_base_compose.py`／`b5_virt_abc.py`／B45 §4.5.10bis-T／CURRENT_PHASE／AGENT_HANDOVER | **薄い一片あり F0+F1**: 言語＝`max_n`一致／`has_piece_set`でIDENTITY盛り一本化／SUB単一片禁止文。切戻し `--piece-set-max 0`。clasp不要 | Git revert／flag 0 |
| 2026-11-03 | `b4_wrap_hand_hold.py`／`sub_image_food_canon_gate.py`／`b5_food_canon_status.py`／B45 §4.5.11 | **P3-W第2**: WRAP確定後手持ち候補UI。切戻し `B4_WRAP_ASSET_LANE=0`。clasp不要 | env 0／Git revert |
| 2026-11-03 | `b5_food_canon_status.py`／`コード.js`／`food_unit_canon.py`／gate／B45 §4.5.11 | **hasWrap鍵**: 親SKU→商品名ベース→JAN。export全件に `hasWrap` bool＋`byParentSku`。要出品 clasp | Git revert |
| 2026-11-03 | `b4_wrap_asset_rules.py`／`dry_b4_wrap_asset_virt.py`／`コード.js`／`b5_food_canon_status.py`／B45 §4.5.11 | **P3-W第1**: WRAPあり→themeId=21個包装訴求+1（最大11）。切戻し `B4_WRAP_APPEAL_THEME=false`。要出品 clasp | Prop false／Git revert |
| 2026-11-02 | `b4_sub_picture_rules.py`／`dry_b4_sub_picture_virt.py`／`dry_sub_overview_copy_test.py`／B45 §4.5.10 | **P3第1**: サブへ WRAP scrub・look・frontOk合否・袋判定寄せ。切戻し `B4_SUB_PICTURE_GATE=0` | env 0／Git revert |
| 2026-11-02 | B45 §4.5.9／CURRENT_PHASE／AGENT_HANDOVER／CHANGE_LEDGER | **P2第2人間検収OK（閉じた）**: かつお／まぐろ新規。次＝P3 | docs revert |
| 2026-11-02 | `コード.js`／`b4_theme_draft_rules.py`／`dry_b4_theme_draft_virt.py`／B45 §4.5.7 | **B-⑤下書き一致検証**: 親SKU・商品名ベース不一致はreuseしない。MAIN参考も親SKU優先。仮想16PASS。要出品 clasp | Prop／Git revert |
| 2026-11-02 | `コード.js`／B45 §4.5.9 | **P2第2GAS**: appealAxis＋つなぎ門。切戻し `B4_THEME_COHESION_GATE=false` | Prop false／Git revert |
| 2026-11-02 | `b4_theme_cohesion_rules.py`／`dry_b4_theme_cohesion_virt.py`／B45 §4.5.9 | **P2第2仮想**: T6/T7/T13。第1閉じた | 仮想のみ削除可 |
| 2026-11-02 | `b4_theme_text_rules.py`／`コード.js`／B45 §4.5.8 | **T4緩和**: 表記ゆれ同一視＋overview含め片寄り判定。仮想12/12 PASS | Git revert |
| 2026-11-02 | `コード.js`／B45 §4.5.8 | **P2第1スライスGAS**: `b4RefineProposalsWithTextGate_`。切戻し `B4_THEME_TEXT_GATE=false` | Prop false／Git revert |
| 2026-11-02 | `b4_theme_text_rules.py`／`dry_b4_theme_text_virt.py`／B45 §4.5.8 | **P2第1スライス仮想**: 門T2/T3/T4/T5/T8/T15。11/11 PASS | 仮想のみ削除可 |
| 2026-11-02 | `コード.js`／B45 §4.5.7 U1 | **再利用時スキップ**: MAIN走査・マスタ／B3を省略。`reusedCount`/`freshCount`を画面表示 | Git revert |
| 2026-11-02 | `コード.js`／`b4_theme_draft_rules.py`／`dry_b4_theme_draft_virt.py`／B45 §4.5.7 | **B-⑤ P1案UI**: 下書き再利用・レ点外し再作成・0.8s裏保存。切戻し `B4_THEME_DRAFT_REUSE=false` | Prop false／Git revert |
| 2026-11-02 | `B45_MAIN_SUB_ALIGN` §4.5／`CURRENT_PHASE`／SUB HUMAN_RUN／AGENT_HANDOVER | **MAIN→サブ寄せ一覧 P0**: 入れない以外は寄せ検討。U/T/G表＋P1〜P4 | docs revert |
| 2026-11-02 | `b5_main_base_compose.py`／`b5_virt_abc.py`／`b5_review_feedback_templates.py`／`コード.js` | **ミックス料理混入防止＋レビュー定型2本**: dishPartをcomposeへ残す。料理上／個包装を見本どおりボタン。要出品 clasp（定型GAS） | Git revert |
| 2026-11-01 | `b5_virt_abc.py`／`b5_food_look_read.py`／`コード.js` | **旧WRAP混ぜ文の除去＋手前判定を厳しく**: 保存案の『小皿に少し混ぜて出せ』をscrub。frontOk欠落は不合格。要出品 clasp（場面） | Git revert |
| 2026-11-01 | `コード.js`（B-④再作成） | **再作成後レ点ONの下書き残し**: 遅延保存がオフを上書きしない。再作成成功時は全案 selected=true で保存。要出品 clasp | Git revert |
| 2026-11-01 | `food_unit_canon.py`／`b5_virt_abc.py`／`b5_food_look_read.py`／`コード.js`／B5・ALIGN要件 | **WRAP置きA＋B**: 皿の横／後ろのみ（小皿混ぜ廃止）。カメラ最近傍は開封後の個品。要出品 clasp（場面） | Git revert |
| 2026-11-01 | `コード.js`（B-④正本） | **紐付け確定が無反応対策**: 成功後は先に画面切替／inline onclick＋準備完了表示／サーバ失敗をボタン下に出す。他親候補除外も継続。要出品 clasp | Git revert |
| 2026-11-01 | `コード.js`（B-④正本） | **スクリプト未準備の原因**: 訴求軸ラベル行の引用符ミスでダイアログJS全体が構文エラー→`b5ClickGoBind`未定義。要出品 clasp | Git revert |
| 2026-11-01 | `コード.js`／B5要件・HUMAN_RUN | **再作成＝訴求軸選択**: 候補ラジオ＋自由入力→ペルソナ〜場面をセット再生成。他案軸は除外。要出品 clasp | Git revert |
| 2026-11-01 | `コード.js`（`b5AlignSavedProposalsToChildCount_`）／B5要件 | **案不足は追加生成のみ**: 保存案は触らず、子レ点より少ない件数だけ足す。多いときは切るだけ。要出品 clasp | Git revert |
| 2026-11-01 | `b5_reuse_decide.py`／`.cursor/rules/b5-main-reuse-auto.mdc`／B5要件・HUMAN_RUN／`コード.js`指示 | **B-④再利用自動**: reuse／missing_only／full を Agent が判定。三択をチャットで聞かない。全部作り直しは人の明示のみ。要出品 clasp（指示文） | Git revert／規則オフ |
| 2026-11-01 | `competitor_jitta/CatalogEnum.js`／`JittaResolve.js`／JITTA／README | **Catalog落ちあと拾い強化**: 起動先頭で次予約／実態AI・ヒットから`catalog_rearm`／`updated`古ロック掃除／inflight重複抑止。仮想 rearm PASS。要競合 clasp | Prop OFF／時計停止／backup `pre_catalog_rearm_20261101_151621` |
| 2026-11-01 | `b5_bag_aspect.py`／`b5_main_base_compose.py`／`food_unit_canon.py`／`b5_virt_abc.py`／`dry_b5_bag_wrap_virt.py`／`コード.js`／B5要件 | **B5袋・個品ロック**: SCALEから袋切って測る／WRAP1〜2／料理上もIDENTITY／food_piece優先採用／printLightOk／窓内正本指示／案で場面をキャッチ場へ1回合わせ。要出品 clasp（案） | Git revert |
| 2026-11-01 | `competitor_jitta/CatalogEnum.js`／README | **Catalogキュー全読み撤去**: 1実行1回マップ＋追記継ぎ足し。失敗で`comp_enum`残さず`catalog_skip`。stuck heal。仮想 queue_fix PASS。要競合 clasp | 旧全読みに戻すなら Git／backup |
| 2026-10-31 | `competitor_jitta/CatalogEnum.js`／`JittaResolve.js`／`PurchaseResearch.js`／JITTA要件／README | **改訂P2 Catalog増員**: 競合列挙＋共有シート`Catalog親ロック`＋落ちあと再予約。未設定=`CJ_CATALOG_ENABLED`OFF。要競合＆① clasp・SP-API鍵 | Prop OFF／時計停止／`PR_SHARED_CATALOG_LOCK=false`／local_backup `pre_p2_catalog_20261031_224712` |
| 2026-11-01 | `b5_food_look_read.py`／`b5_main_base_compose.py`／`b5_virt_abc.py`／`food_unit_canon.py`／`dry_b5_bag_wrap_virt.py`／B5要件 | **WRAP look**: 個包装見本を機械読み（密着・空気・材質・縁等）→指示へ。案件固定の真空禁止にしない。clasp不要 | Git revert |
| 2026-10-31 | `food_unit_canon.py`／`b5_virt_abc.py` | **WRAP≠大袋**: WRAP写真の大袋・背景・手は無視。PACKAGEのみ大袋正。場面文の「個包装描くな」はWRAPあり時に除去。clasp不要 | Git revert |
| 2026-10-31 | `b5_bag_color.py`／`b5_bag_aspect.py`／`b5_main_base_compose.py`／`b5_virt_abc.py`／B5要件 | **包材印刷色**: 正本ベタ面を画素で彩度・明度・RGB数値化→指示。合否は袋判定の printHueOk（黒潰れ再生成）。細部は数値化しない。clasp不要 | Git revert |
| 2026-10-31 | `food_unit_canon.py`／`b5_main_base_compose.py`／`b5_virt_abc.py`／B5要件 | **袋解決＝JAN+oya枝**（ベンダー接頭辞違い吸収）／試し採用でOCR同群内は袋・個片・顔優先／個品明度はIDENTITY優先。clasp不要 | Git revert |
| 2026-10-31 | `コード.js` | **B-④案レ点初期ON**: 表示時は常にチェック。下書きの selected=false 残りを消す。要出品 clasp | Git revert |
| 2026-10-31 | `コード.js`／B5要件 | **ミックス料理ラジオ欠落対策**: 再作成・再利用後も dishCandidates を付け直し。候補空でも構成列ラジオ枠を出す。要出品 clasp | Git revert |
| 2026-10-31 | `b5_bag_aspect.py`／`b5_main_base_compose.py`／`food_unit_canon.py`／`b5_virt_abc.py`／`コード.js`／B5・B45 | **袋判定の結果書き出し**（両袋必須＋対比）／**WRAPありは皿横or少し混ぜて必須表示**。要出品 clasp（場面文） | Git revert |
| 2026-10-31 | `コード.js`／B5要件／HUMAN_RUN | **B-④引き継ぎ一本化**: catch／salesだけの overlay 廃止。正は下書き／compose丸ごと再利用。要出品 clasp | Git revert |
| 2026-10-31 | `コード.js` | **B-④再作成後レ点自動復帰**: 外した案の `selected=false` を作り直し後 true に戻す（指示漏れ防止）。要出品 clasp | Git revert |
| 2026-10-31 | `コード.js` | **B-④文字色**: 正本確定の成功時にドロップダウンを再読込してから案作成（切替待ち中の青選択を拾う）。要出品 clasp | Git revert |
| 2026-10-31 | `コード.js` | **B-④コピー紐付け＝スロット**: themeId＋#sellRank（管理番号）。ペルソナ単独廃止。要出品 clasp | Git revert |
| 2026-10-31 | `コード.js` | **B-④レ点なし再作成＝対象だけ**: remakeIndexes＋keep。料理プール省略。要出品 clasp | Git revert |
| 2026-10-31 | `コード.js` | **B-④コピー紐付け＝ペルソナ**: （撤回）スロットキーへ置換。要出品 clasp | Git revert |
| 2026-10-31 | `コード.js` | **B-④下書きA+B**: 作成直後Drive保存＋入力0.8秒デバウンス裏保存。要出品 clasp | Git revert |
| 2026-10-31 | `コード.js` | **B-④キャッチ≠セールス**: 梯子／Gemini同文を検知し差し替え（S2）。要出品 clasp | Git revert |
| 2026-10-31 | `コード.js` | **B-④料理ラジオ変更→場面文追従**（料理名差し替え・載せる構成。Geminiなし）。要出品 clasp | Git revert |
| 2026-10-31 | `コード.js`／B5要件 | **ミックス料理ラジオ＝構成列**＋選択で `dishPart*`。食べ方検索はマスタ→公式の不足時のみ（アレンジ／レシピ）。要出品 clasp | Git revert |
| 2026-10-31 | `コード.js`／B5要件 | **B-④案数＝子レ点＋下書き保持**: 画面・指示は子数に切る。下書きはマージで残す（子増後も前回案可）。要出品 clasp | Git revert |
| 2026-10-31 | `コード.js`／B5要件／HUMAN_RUN | **B-④前回案丸ごと再利用**: あれば生成スキップ（draft→compose）。戻るは保持。作り直し＝レ点なしのみ。要出品 clasp | Git revert |
| 2026-10-31 | `food_canon_manual_ingest.py`／gate／B45 | **手動見本最大6枚**（旧3）。競合削除は無し。clasp不要 | `MANUAL_MAX=3`／Git |
| 2026-10-31 | 競合JITTA要件／CYCLE／FLOW／README／HUMAN_RUN／AGENT_HANDOVER | **改訂P2要件ロック**: Catalog増員並走・①沈黙禁止。コード未。仮想 `_tmp_virtual_p2_revised.py` | docs revert |
| 2026-10-31 | `purchase_research/PurchaseResearch.js`／HUMAN_RUN／FLOW／AGENT_HANDOVER | **案P1 案E落ちあと拾い**: 心拍liveness＋洗い終わりで再武装。幽霊解消。実行中リース。finally。新番犬なし。要① clasp | `PR_ENUM_RESET_RUN_ENUM=false`／Git |
| 2026-10-31 | `food_unit_canon.py`／`sub_image_food_canon_gate.py`／`b5_main_base_compose.py`／`b5_virt_abc.py`／ALIGN | **WRAP最小**: 個包装見本（任意）確定・添付。無いときは創作禁止。clasp不要 | Git revert |
| 2026-10-31 | `コード.js`／`food_unit_canon.py`／`b5_virt_abc.py`／B5要件 | **ミックス①②③**: 構成別食べ方プール＋`dishPart*`（料理に載せる構成）＋個包装創作禁止。要出品 clasp | Git revert |
| 2026-10-31 | `food_unit_canon.py`／`b5_main_base_compose.py`／`sub_image_food_canon_launch.py`／`dry_sub_overview_copy_test.py`／`コード.js`／B45 | **ミックス個品紐付け＝マスタセット構成1〜5**（名前中黒廃止）。B-④ JSON に comboParts。①′は `--parent-sku` で組み合わせ。要出品 clasp | Git revert |
| 2026-10-30 | `広告運用GAS/UnadvertisedAsin.js`／HUMAN_RUN §4.1 | **3-②原因ログ**: `id_probe`／`bulk_resolve`（path・Hit・Near）。挙動は変えず観測のみ。要広告 clasp | Git revert |
| 2026-10-30 | `コード.js`／B5要件 | **B-④案再作成**: ボタン「レ点を外した案を再作成する」。レ点なしだけ差し替え（残す案はレ点したまま）。要出品 clasp | Git revert |
| 2026-10-30 | `コード.js`／B3確認docs | **B-③連続実行**: 楽天Yahoo/Amazon確認ダイアログを既定OFF。Amazon込み・Keepa全件なし。切戻し `B3_IMAGE_ASK_MALL_MODE=true`。要出品 clasp | Prop true／Git |
| 2026-10-30 | `コード.js`／B3確認docs | **B-③内モール横断自動**: 不足JAN（構成展開含む）を `menuTestCrossMallSetCountJudge(jans)` で自動実行。別メニュー不要。要出品 clasp | `askConfirm:true`／Git |
| 2026-10-29 | `コード.js` | **途中欠け復旧**: `_local_backup/gas/20261027_080002` を正に戻し、B-③ミックス連動のみ再適用。Git10/5復元版は `コード.js.before_restore_from_g27`。要出品 clasp | 同 before ファイル／Prop `B3_EXPAND_COMBO_PARENT_JANS=false` |
| 2026-10-29 | `コード.js`／B3確認docs | **B-③ミックス連動**: レ点時にセット構成1〜5の単品親JANも画像取得。要出品 clasp。切戻し `B3_EXPAND_COMBO_PARENT_JANS=false` | Prop false／Git |
| 2026-10-29 | `c1_packaged.py`／列マップ／FOOD列マップdocs | **ハイライト100473**: 「抜群」等を「良い」へ置換してからPACKAGED。oya3再PACKAGED済 | Git revert／phrases空 |
| 2026-10-29 | `c1_bulk_fill_by_name.py`／C1要件／T1 HUMAN_RUN | **dry_run xlsm 廃止**: B-T1 dry_run は `_DRYRUN.xlsm` を 03 に出さない（05のみ）。prod のみ PACKAGED | `write_dryrun_xlsm: true` に戻す／Git |
| 2026-10-29 | `purchase_research/PurchaseResearch.js`／HUMAN_RUN／FLOW | **N3 親分割並行**: 列挙＝メーカー名ロックのみ／洗い＝全体＋親。案E自己再予約＋洗い終了再武装。番犬増なし。要① clasp | `PR_WASH_PARENT_PARTITION=false`／Git |
| 2026-10-28 | `purchase_research/PurchaseResearch.js`／HUMAN_RUN | **案E enumリセット＋その場列挙**: 空振りをリセットしメニューで既存列挙。切れは自動再開。心拍分岐なし。要① clasp | `PR_ENUM_RESET_RUN_ENUM=false`／Git |
| 2026-10-28 | `purchase_research/PurchaseResearch.js`／CYCLE／FLOW／HUMAN_RUN | **title0_skip 再オープン**: メニューで列挙待ち＋Prop FAIL/SKIPクリア。要① clasp | メニュー使わない／Git |
| 2026-10-28 | `purchase_research/PurchaseResearch.js`／CYCLE／FLOW／HUMAN_RUN | **enum_empty_out 再オープン**: メニューで調査対象外→列挙待ち。要① clasp | メニュー使わない／Git |
| 2026-10-28 | `purchase_research/PurchaseResearch.js`／CYCLE／FLOW／HUMAN_RUN | **Catalog親辿り**: classifications の `parent` で調査ルート照合（食品が下位IDで落ちない）。要① clasp | Prop `PR_CATALOG_CLASSIF_WALK_PARENT=false`／Git |
| 2026-10-27 | `AmazonSpapiPut.js`／D_ENTRY HUMAN_RUN | **相乗りASIN**: 子行N列のみ（親フォールバック廃止）＋実行内ユニーク（先勝ちスキップ）。要出品 clasp | Git revert |
| 2026-10-27 | `purchase_research/PurchaseResearch.js`／HUMAN_RUN | **survey_diag**: Amazon大分類 vs ロジック許可を①ログへ（`allow=` `dropTop=`）。要① clasp | Prop `PR_SURVEY_DROP_DIAG=false`／Git |
| 2026-10-27 | `purchase_research/PurchaseResearch.js`／FULL／CYCLE／V1／HUMAN_RUN | **Catalog OR**: 検索ヒットを調査カテゴリ後にキュー（タイトル必須やめる）。切戻し `PR_CATALOG_TITLE_GATE=true`。要① clasp | Prop true／Git |
| 2026-10-27 | `YahooCaptionRich.js`／YAHOO_CAPTION | **セットリンク**: 親SKU兄弟＋袋/個パース。骨格単品JANリンクを残さない。要出品 clasp＋説明API | Git revert |
| 2026-10-27 | `YahooCaptionRich.js`／YAHOO_CAPTION | **案C**: 【商品について】枠廃止・注記残す・原料はcaptionのみ。要出品 clasp＋説明API | Git revert |
| 2026-10-27 | `YahooCaptionRich.js`／`build_yahoo_caption_themes.py`／YAHOO_CAPTION | **セット訴求は先頭1枠／原料重複ラベル除去／紹介種から原料除外**。要fit-only＋出品 clasp＋説明API | Git revert |
| 2026-10-27 | `コード.js`／BATCH_EXPORT_IMAGE_GATE | **D画像スキップ初期レ点**: 出品CKの楽天メイン画像1そろい時ON。再送は人間が外す。要出品 clasp | Git revert |
| 2026-10-27 | `YahooCaptionRich.js`／`build_yahoo_caption_themes.py`／YAHOO_CAPTION | **ミックス訴求（種数・品名）＋原料複数成分の改行**。要フィット再生成＋出品 clasp＋説明API | Git revert |
| 2026-10-27 | `YahooCaptionRich.js`／`build_yahoo_caption_themes.py`／YAHOO_CAPTION | **原料＝列`原料`・食品のみ／内容量行排除**。要出品 clasp＋説明API | Git revert |
| 2026-10-26 | `build_yahoo_caption_themes.py`／YAHOO_CAPTION | **Yahoo説明P6cミックス**: 構成品商品名→マスタ安全欄まとめ。カテゴリ非限定。要パック再生成＋出品 clasp＋説明API | Git revert |
| 2026-10-26 | `build_yahoo_caption_themes.py`／`YahooCaptionRich.js`／P6C HUMAN_RUN | **Yahoo説明P6c**: 同一JANを親SKU別に再組立（oya2/oya3）。旧oyaパック廃止。要 clasp＋lib UP＋Property ON＋説明API | Git revert／Property false |
| 2026-10-26 | `sub_image_mix_review_loop.py`／`sub_image_review_loop.py`／`b5_skip_generate_finish.py`／AUTOBIND HUMAN_RUN | **サブ出口=親1セット**: 採用分を子SKU名1件だけ出す。JAN／全子複製をやめ枠外候補を防ぐ。clasp不要 | Git revert |
| 2026-10-27 | `competitor_jitta/JittaResolve.js`／競合JITTA要件／README | **P1ブースト既定1000**（旧500）。要競合 clasp | Prop `CJ_DAILY_CAP_BOOST=500`／Git |
| 2026-10-26 | `purchase_research/PurchaseResearch.js`／CYCLE／HUMAN_RUN | **列挙原因の特定**: quiet title0 でも `why=title0`＋`pages/titles/surveyDrop`。要① clasp | Prop `PR_ENUM_DIAG_LOG=false`／Git |
| 2026-10-26 | `purchase_research/PurchaseResearch.js`／CYCLE／HUMAN_RUN | **列挙原因ログ**: ①ログ `enum_diag`／`enum_diag_sum`（`why=`）。条件見直し前の観測。要① clasp | Prop `PR_ENUM_DIAG_LOG=false`／Git |
| 2026-10-26 | `purchase_research/PurchaseResearch.js`／CYCLE／HUMAN_RUN | **案A強化**: 店続きより列挙優先。列挙後 sm=0 で skip_enum 剥がし。要① clasp | Prop `PR_CYCLE_ENUM_WHEN_SM0=false`／Git |
| 2026-10-25 | `purchase_research/PurchaseResearch.js`／CYCLE／HUMAN_RUN | **案A 小中=0で余力あり列挙**: `PR_CYCLE_ENUM_WHEN_SM0`。ログ `enum_when_sm0`。3回で店譲り。要① clasp | Prop `false`／Git revert |
| 2026-10-25 | `competitor_jitta/JittaResolve.js`／競合JITTA要件／CYCLE §4.5bis／README | **P1 未調査山ブースト**: U≥31で `CJ_DAILY_CAP_BOOST`（既定500）。ログ `U=` `mode=`。C係交代なし。要競合 clasp | Prop BOOST削除／Git revert |
| 2026-10-25 | `コード.js`／`dry_sub_overview_copy_test.py`／ALIGN§7.6bis | **サブ文字色**: 指示に `--parent-sku`。truth-hint からも補完。単品・ミックス同じ。要出品 clasp | Git revert |
| 2026-10-25 | `sub_image_mix_review_*.py` | **混在レビュー**: 同じJAN別商品でマニフェスト上書きしない。手順「いまやること」表示。clasp不要 | Git revert |
| 2026-10-25 | `コード.js`／ALIGN | **B-⑤正本プルダウン高速化**: 親SKU優先・最大12・サムネ表示分だけ（B-④と同型）。要出品 clasp | Git revert |
| 2026-10-25 | `コード.js`／B45 ALIGN §7.2.1 | **B-⑤テーマ1件再提案**: ▼指示欄＋再提案。他レ点維持。単品・ミックス同じ。要出品 clasp | Git revert |
| 2026-10-25 | `master_sets.py`／SET_MAIN_LAYOUT_RULES | **ミックス金丸**: 売る数量×構成数量×種。`バリエーション値`先頭（3袋/3種）の二重掛けを廃止。clasp不要 | Git revert |
| 2026-10-25 | `b5_main_base_review_loop.py`／`b5_main_base_compose.py` | **再作成①②**: ミックスは nameBase＋PART部品を渡す（単体 theme_new_run 落ち禁止）。案はファイル見出し固定・同テーマ先頭フォールバック禁止。焼成不一致は `__r2` に載せない。clasp不要 | Git revert |
| 2026-10-25 | `b5_main_base_review_launch.py`／rules／skill | **レビュー --fresh**: 開き直しは旧プロセス停止＋最新compose。古いURL渡し禁止。clasp不要 | Git revert |
| 2026-10-25 | `b5_main_base_review_loop.py` | **MAINレビュー文言**: 欄は焼成（catchBurn／sales）優先。再生成はファイル見出しで案特定＋新キャッチでファイル名。clasp不要 | Git revert |
| 2026-10-25 | `コード.js`／マスタHW | **バリエーション値**: 空欄時は GR 表示の直書き禁止→`=GRn`。2種ミックス該当行の直値を式に復旧。要出品 clasp | Git revert／HW式を戻す |
| 2026-10-25 | `b5_main_base_compose.py` | **試し採用キー**: 指定キャッチOCR合格を袋縦横より先に残す（袋印字の別文言が残るのを止める）。clasp不要 | Git revert |
| 2026-10-25 | `広告運用GAS/UnadvertisedAsin.js` | **再UL already exists削減**: 既存キャンペーン昇格強化、PA/除外の既存スキップ、括弧NFKC廃止。v1.13。要広告 clasp＋一括DL取り直し | Git revert |
| 2026-10-25 | `広告運用GAS/UnadvertisedAsin.js` | **ID抽出でAG取りこぼし修正**: 一括DLのAG行はキャンペーン名が空でもキャンペーンID＋AG名で取る。v1.12。要広告 clasp | Git revert |
| 2026-10-25 | `広告運用GAS/UnadvertisedAsin.js` | **バルクalready exists／×文字**: ID参照で既存昇格、AGをキャンペーンIDでも照合、新規名の`×`除去。v1.11。要広告 clasp | Git revert |
| 2026-10-25 | `広告運用GAS/UnadvertisedAsin.js` | **バルク親IDエラー修正**: 既存キャンペーンでAG数値ID無しのとき仮ID＋AG作成行（名前だけID禁止）。v1.10。要広告 clasp | Git revert |
| 2026-10-25 | `広告運用GAS/UnadvertisedAsin.js`／`コード.js`／HUMAN_RUN／第1スライス要件 | **ID参照用自動更新**: `📥一括キャンペーンDL`→一覧・3-②で書込。切戻し Prop `UNADVERTISED_ASIN_ID_AUTO_FROM_BULK=false`。要広告 clasp＋3-③でマニュアル再生成 | Prop false／Git revert |
| 2026-10-25 | `b5_main_base_review_loop.py` | **子SKUレ点**: 選択時にfamily奪取しない。1画像1子＋同じSKUだけ他カードから移す。clasp不要 | Git revert |
| 2026-10-25 | `food_unit_canon.py`／`b5_virt_abc.py`／ALIGN §7.15 | **ミックスMAIN**: 袋PART順と個品IDENTITY順を連動（小皿も同じ並び）。「左右自由／種類が分かればよい」を廃止。clasp不要 | Git revert |
| 2026-10-25 | `b5_main_base_auto_assign.py`／`b5_main_base_review_loop.py` | **MAINレビュー子割当**: 画像1子・見出し単位（themeId潰し廃止）。claimもfamilyのみ（上下まとめて外れ修正）。clasp不要 | Git revert |
| 2026-10-25 | `b5_main_base_auto_assign.py` | **review起動**: note文字列の括弧未閉じで SyntaxError。閉じ復旧。clasp不要 | Git revert |
| 2026-10-24 | `b5_virt_abc.py` | **MAIN compose**: `_prompt_p3_free` で `accent_text_block` 未代入→NameError。代入を復旧。clasp不要 | Git revert |
| 2026-10-24 | `コード.js`／ALIGN | **B-④①対象親**: JAN重複排除を廃止し親SKU単位（oya2／oya3同JANでも両行列）。要出品 clasp | Git revert。要 clasp |
| 2026-10-24 | `コード.js`／ALIGN §7.12.3 | **B-④①正本高速化**: シードは共有空のときだけ。一覧は01直下＋親SKU／保存rel。サムネ小・最大12。要出品 clasp | Git revert。要 clasp |
| 2026-10-24 | `コード.js` | **相乗りASIN無し**: soft skip＋次手順メッセージ（新規だけ／ASIN入れて相乗り）。要出品 clasp | Git revert。要 clasp |
| 2026-10-24 | `コード.js`／`AmazonApprovalExport.js`／`AmazonSpapiPut.js`／P0／D_ENTRY | **ALLOW_MASTER_QTY ゲート廃止**: マスタ在庫は Property 不要。開始前確認は残す。要出品 clasp | Git revert。要 clasp |
| 2026-10-24 | `mix_column_layout.py`／SET_MAIN §1.3 | **ミックス Octas**: 基準を表示後1袋×0.075に（列合算で過大だった）。clasp不要 | Git revert |
| 2026-10-24 | `mix_column_layout.py`／SET_MAIN §1.3 | **ミックス列積み Octas**: 単品と同じ不透明×0.075。幅×0.10廃止。clasp不要 | Git revert |
| 2026-10-24 | `set_main_image` mix分岐 | **ミックス each**: 1=正本N1／2–4=列積み／≥5=正本×単品。塗り順・SKU区切り一致。clasp不要 | Git revert |
| 2026-10-24 | `purchase_research/PurchaseResearch.js`／CYCLE／FLOW／HUMAN_RUN | **門中も自動レ点可**: 大特大禁止と分離。旧レ点止めは `PR_JITTA_BACKLOG_BLOCK_AUTO_CK=true`。要① clasp | Prop true／`PR_JITTA_BACKLOG_IDLE=false`／Git |
| 2026-10-24 | `competitor_jitta/JittaResolve.js` | **実態AIプロンプト案B**: 4キーのみ・出力768。業種URL棚reasonは取らない。要競合 clasp | `local_backup_thinB_*`／Git revert |
| 2026-10-24 | `purchase_research/PurchaseResearch.js`／CYCLE／FLOW／HUMAN_RUN | **門中全面idle廃止**: 大特大だけ禁止（`jitta_backlog_xl_block`）。自動レ点抑制は維持。要① clasp | `PR_JITTA_BACKLOG_IDLE=false`／Git revert |
| 2026-10-24 | `competitor_jitta/JittaResolve.js` | **実態AI対象集めの過剰読取を削除**: サマリ1回＋マスタ1回で未調査を優先。起動上限切れ対策。要競合 clasp | 手元 `JittaResolve.js.local_backup_20261024_161152`／Git revert |
| 2026-10-24 | `purchase_research/PurchaseResearch.js`／CYCLE／FLOW／HUMAN_RUN | **幽霊社=調査対象外＋未調査ゲート中は仕分け待ち**: `enum_empty_out`／`jitta_backlog_gate`。要① clasp | `PR_ENUM_EMPTY_OUT=false`／`PR_JITTA_BACKLOG_IDLE=false` |
| 2026-10-24 | `purchase_research/PurchaseResearch.js`／CYCLE／HUMAN_RUN／FLOW | **C0→大特大譲りを撤回**: 未設定=OFF。本線は小中補充。要① clasp | Prop 消す／false |
| 2026-10-24 | `purchase_research/PurchaseResearch.js`／CYCLE §4.4／HUMAN_RUN／FLOW | **（誤）余力ありC0→大特大**: 同日撤回。残さない | 撤回済み |
| 2026-10-24 | `purchase_research/PurchaseResearch.js`／CYCLE §4.1bis／FULL／HUMAN_RUN／FLOW | **Catalogで調査ルート以外を切る**: classifications取得→ルート照合→キュー投入前に落とす。欠測も落とす。ログ `surveyDrop`／`surveyMiss`。切戻し `PR_CATALOG_SURVEY_ROOT_GATE=false`。要① clasp | Prop false。Git revert |
| 2026-10-23 | `YAHOO_OAUTH_REAUTH_HUMAN_RUN`／rules／HANDOVER | **Yahoo戻り先の店正本**: `YAHOO_OAUTH_REDIRECT_URI`=`http://localhost/`。50件超は setProperty。エージェントは §0 先読み | docs／rule を戻す |
| 2026-10-23 | `コード.js`／CABINET_AUTO_FOLDER要件 | **R-Cabinet fileId後はsearchのみ再試行**: insert_only／search_only分離。Search失敗時 search_only_resume。容量系のみ縮小再insert。要出品 clasp | Prop retry false／Git revert。要 clasp |
| 2026-10-23 | `コード.js`／CABINET_AUTO_FOLDER要件 | **R-Cabinet UrlFetch切り分けを一括出品ログへ**: transient／recovered／final_fail／host_probe＋folderId・newUpStreak等。再試行自体は同日実装。要出品 clasp | Prop retry false／Git revert。要 clasp |
| 2026-10-23 | `コード.js`／CABINET_AUTO_FOLDER要件 | **R-Cabinet UrlFetch限定再試行**: insert／search の「使用できないアドレス」・429・5xx を最大3回（1s→2s）。attempt／ms／host_probe をログ。未設定=ON。要出品 clasp | Prop `RAKUTEN_CABINET_URLFETCH_RETRY_ENABLED=false`／Git revert。要 clasp |
| 2026-10-22 | `コード.js`／MATRIX_LAYOUT／SKU_AUTOBIND要件 | **マトリクスサブ全子同一**: `_subN`は親のみ。L2上書き共有＋L3後再配。子別Drive ID増殖止め。要出品 clasp | Prop LAYOUT false／Git revert。要 clasp |
| 2026-10-21 | `purchase_research/PurchaseResearch.js`／HUMAN_RUN／HANDOVER | **ASIN列走査を25000分割**: 索引・空行掃除・ずれ時。Prop `PR_KEEPA_FULL_ASIN_COL_CHUNK`（0=旧一括）。要① clasp | Prop=0／Git revert |
| 2026-10-21 | `purchase_research/PurchaseResearch.js`／HUMAN_RUN／HANDOVER | **ASIN列チャンク比較プローブ**: 99→4。幅5千〜4万の通否＋全行カバー totalMs。書込なし。要① clasp | メニュー削除／Git revert |
| 2026-10-21 | `purchase_research/PurchaseResearch.js`／HUMAN_RUN | **99メニュー階層化**: 1止める・診断／2洗い／3 D1D2／4 Keepaフル／5セラー／6メーカー／7サイクル／8バックアップ。画面はみ出し対策。要① clasp | Git revert |
| 2026-10-21 | `purchase_research/PurchaseResearch.js`／HUMAN_RUN／HANDOVER | **Keepaフル用量プローブ（読取のみ）**: 99→4 Keepaフル。段階 getValues の通否を①ログへ。書込なし。要① clasp | メニュー削除／Git revert |
| 2026-10-21 | `コード.js`／CABINET_AUTO_FOLDER要件 | **R-Cabinet フォルダ自動切替**: 上限／予防でルート直下に新フォルダ→B9更新→同じD続き。診断ログ。未設定=OFF。要出品 clasp＋`RAKUTEN_CABINET_AUTO_FOLDER_ENABLED=true` | Prop false。Git revert。要 clasp |
| 2026-10-21 | `docs/org/D_MENU_RAKUTEN_CABINET_AUTO_FOLDER_REQUIREMENTS.md` | **R-Cabinet フォルダ自動切替要件ロック**: ルート直下・同じD続き・停止時原因ログ。切戻し `RAKUTEN_CABINET_AUTO_FOLDER_ENABLED` | docs Git revert |
| 2026-10-21 | `コード.js` | **D診断ログ**: 一括出品ログへ予約／トリガー開始終了／親ごとの画像UP／心拍／失敗分類（URLFETCH等）。鍵は出さない。要出品 clasp | Git revert。要 clasp |
| 2026-10-20 | `glyph_assets.py`／`master_sets.py`／`rakuten_layer.py`／`compose_set_main.py`／レイアウト要件 | **金丸**: 一桁は1の高さで下辺を単位に揃える。ミックスは合計と「各」。単品はセット。切戻し `mixEach.enabled=false`。商品写真は未生成。clasp 不要 | `mixEach.enabled=false`。Git revert |
| 2026-10-19 | `コード.js` | **抜け2件を再掲**: AI同期は同名でもメーカー違いを混ぜない。E③前に09の未取込xlsmを棚registryへ。棚が無ければ案内して止める。戻し各 Prop false。要出品 clasp | Prop false。Git revert。要 clasp |
| 2026-10-19 | `コード.js`／CDE要件 | **C／D／Eメニュー再掲**: 9月18日の控えから入口だけ戻した。Cは楽天仕分け、Dは楽天・Yahoo、EはAmazon3項目。日次バックアップと右並びは残す。要出品 clasp | Git revert。要 clasp |
| 2026-10-19 | `コード.js`／MATRIX_LAYOUT要件 | **楽天マトリクス L1〜L3 再掲**: 関数がファイルに無かった。他セットのメインを右の空きへ、親は最少セットの子と同じ。Driveは増やさない。戻し `RAKUTEN_MATRIX_LAYOUT_EXPAND_ENABLED=false`。要出品 clasp | Prop false。Git revert。要 clasp |
| 2026-10-20 | `purchase_research/PurchaseResearch.js`／HUMAN_RUN | **Keepaフル読取・書込の生JSON分離**: マップ／flatは必要列のみ（chunk既定1）。追記は生JSON単独setValue。B4はサイズ上限で要修復＋再予約停止。要① clasp | Git revert／chunk Prop |
| 2026-10-20 | `purchase_research/PurchaseResearch.js`／HUMAN_RUN／HANDOVER | **Keepaフル案A（ASIN限定読取）**: マップ／flat／再門／追記前マップは要求ASINのみ。空キャッシュ握りつぶし禁止。段階ログ。切戻し `PR_KEEPA_FULL_ASIN_MAP=false`。要① clasp | Prop false。Git revert |
| 2026-10-20 | `purchase_research/PurchaseResearch.js`／HUMAN_RUN／FLOW | **B4 レ点メーカー実態→洗い**: 自動一時停止→レ点だけ実態AI→洗い→専用トリガーで飲み干し→自動再開。切戻し `PR_B4_ENABLED=false`。要① clasp | Prop false。Git revert |
| 2026-10-19 | `purchase_research/PurchaseResearch.js`／CYCLE §4.1bis／FLOW | **調査カテゴリ外メーカーは洗わない**: サマリ `Amazonルートカテゴリ` が①プロファイルONと無交差なら選定・続き・洗い本体を止め調査対象外。切戻し `PR_SURVEY_ROOT_GATE=false`。要① clasp | Prop false。Git revert |
| 2026-10-17 | `ProjectDailyBackup.js`／`purchase_research/ProjectDailyBackup.js`／HUMAN_RUN／要件 | **PJ別 Drive自動バックアップ第1版**: 出品＋①。`99.バックアップ`配下・昨日8時〜今日8時・破損は保存せずログ／任意メール。切戻し ENABLED=false。要 clasp（出品ルート＋①） | Prop false。トリガー削除。Git revert |
| 2026-10-17 | `docs/org/PROJECT_DAILY_BACKUP_REQUIREMENTS.md`／`project-daily-backup.mdc` | **PJ別 Drive自動バックアップ要件ロック**: `99.バックアップ`配下・昨日8時〜今日8時・破損は保存せず報告。第1版=出品→① | docs Git revert |
| 2026-10-17 | `purchase_research/PurchaseResearch.js`／HUMAN_RUN | **Keepaフル読み書き分割**: 新規行 insert（既定5）／索引書込（既定2000）／空行全列読み（既定40）。「セルの範囲を小さく」対策。切戻し各 Prop=`0`。要① clasp |
| 2026-10-17 | `コード.js`／`dry_b5_propose_virt.py`／B5要件 | **届いた3欄は消さない。キャッチに得たい結果を足さない。体の効能は例示で止める**: 控えは `_local_backup/b5_keep_story_20261017.txt`。戻し: Property `B5_KEEP_STORY=false`。要出品 clasp |
| 2026-10-17 | `コード.js`／`dry_b5_propose_virt.py`／B5要件 | **ペルソナの考える段階を最小にする**: 考える分が本文の上限を使ったため。戻し: Property `B5_THINK_OFF=false`。控えは `_local_backup/b5_think_off_20261017.txt`。要出品 clasp |
| 2026-10-17 | `コード.js`／`dry_b5_propose_virt.py`／B5要件 | **終わりの括弧が無くても完結した人は残す**: 分割は全部つなぐ。途中の最後は捨てる。控えは `_local_backup/b5_join_partial_20261017.txt`。戻し: Property `B5_JOIN_PARTIAL=false`。要出品 clasp |
| 2026-10-16 | `コード.js`／`dry_b5_propose_virt.py`／B5要件 | **3欄は流れで考えてから分ける**: 根拠の次に困り、得たい結果、訴求理由。流れの文は返さない。閉じる前の書き直しはしない。控えは `_local_backup/b5_flow_first_20261016.txt`。戻し: Property `B5_FLOW_FIRST=false`。要出品 clasp |
| 2026-10-16 | `コード.js`／`dry_b5_propose_virt.py` | **キャッチの字数理由と、状況と場面の結び付き理由をログに残す**: シートは B5理由。戻し: Git revert。要出品 clasp |
| 2026-10-16 | `コード.js`／`dry_b5_propose_virt.py` | **ペルソナの受け取りを実行の最後と▼ログに残す**: 先頭のログは切れる。戻し: Git revert。要出品 clasp |
| 2026-10-16 | `コード.js`／`dry_b5_propose_virt.py`／B5要件 | **キャッチは6〜12字の1本**: 【】は付けない。得たい結果の全文は前に足さない。12字超は1回書き直し、まだ超えたら画面に残して焼かない。戻し: Git revert。要出品 clasp |
| 2026-10-16 | `purchase_research/PurchaseResearch.js`／FLOW／CYCLE／HUMAN_RUN | **洗い自動リカバリー（§5.6／§6.11）**: skip直後次予約・ピック除外・stall番犬。切戻し `PR_WASH_AUTO_RECOVER=false`。要① clasp |
| 2026-10-16 | `B_PURCHASE_RESEARCH_FLOW.md` | **洗い停止原因＋自動リカバリー契約（§3 B／§5.6）**: 実態skip後の未予約／washN生存誤認／stall検知のみ。3層は同日コード投入。戻し: Git revert |
| 2026-10-16 | `コード.js`／`dry_b5_propose_virt.py`／B5要件 | **空の案は材料つきで1回やり直す**: 起点は材料の語。誰か結果が空なら3欄とも空。書き直しに材料・誰・起点を渡す。揃わなければ3欄とも空。戻し: Git revert。要出品 clasp |
| 2026-10-15 | `コード.js`／`dry_b5_propose_virt.py`／B5要件 | **買う人はネットの店で注文する人**: 食べる最中と玄関の手渡しは落とす。買う人が空なら状況を書かない。戻し: Git revert。要出品 clasp |
| 2026-10-15 | `コード.js`／`dry_b5_propose_virt.py`／`b5_virt_abc.py`／`b5_main_base_names.py`／B5要件 | **絵の指示は全商品共通。【】は外す**: 見本を前、包材は後ろ。袋を並べて主役にしない。キャッチとセールスのかぎ括弧を付けない。戻し: Git revert。要出品 clasp |
| 2026-10-14 | `コード.js`／`dry_b5_propose_virt.py`／B5要件 | **状況は足場のあと別段**: 人と得たい結果を先に書く。足場は材料にそのままある使い方だけ。状況はそのあと、その場の物まで導く。要件ごとの成否はログ `[b5Req_]`。贈り物でも選んだ料理は残す。戻し: Property `B5_CAUSE_FIX=false`、または Git revert。要出品 clasp |
| 2026-10-14 | `コード.js`／`dry_b5_propose_virt.py`／`b5_virt_abc.py`／B5要件 | **1枚目は状況の欄**: ペルソナに状況を新設。材料は全文。料理名は新しく作らず、既存の食べ方を1つ選んで画面で選択済み。場面はその場が絵で分かるところまで。贈り物は箱を横に置く。戻し: Property `B5_CAUSE_FIX=false`、または Git revert。要出品 clasp |
| 2026-10-13 | `コード.js`／`dry_b5_propose_virt.py`／B5要件 | **個品は商品名の括弧内。食事ならストックでも料理を止めるな**: 売り文句は個品に数えない。得たい結果が食事なら、区分がストックでも料理の候補を作る。戻し: Property `B5_CAUSE_FIX=false`、または Git revert。要出品 clasp |
| 2026-10-13 | `コード.js`／`dry_b5_propose_virt.py`／B5要件 | **料理は人が選ぶ。新しい料理は個品必須**: 元の料理は全部候補。新しい料理は個品を全部使えたときだけ足す。機械は勝ちを決めない。戻し: Property `B5_CAUSE_FIX=false`、または Git revert。要出品 clasp |
| 2026-10-12 | `コード.js`／`dry_b5_propose_virt.py`／`b5_main_base_compose.py`／`b5_virt_abc.py`／B5要件 | **返ったあとを見る。落ちた見出しは焼成しない**: 得たい結果の数、見出し、作り直しの主材料、贈り物の区分を返ったあと1回見る。まだ落ちた見出しは画面に残し、画像には載せない。戻し: Property `B5_CAUSE_FIX=false`、または Git revert。要出品 clasp |
| 2026-10-12 | `コード.js`／`dry_b5_propose_virt.py`／B5要件 | **料理なしは誰も見る。作り直しは1回。本文をログへ**: 贈り物とストックは誰を見て空。作り直しは主材料と作り方が両方あるときだけ。決まらなければ1回足す。キャッチ等の本文をログに残す。戻し: Git revert。要出品 clasp |
| 2026-10-12 | `コード.js`／`dry_b5_propose_virt.py`／B5要件 | **フィットの定義と、同じ1回での作り直し**: 隣に置けるだけでは残さない。合う候補が無ければ、同じ判定の中で作り直す。場面は選び直さない。戻し: Git revert。要出品 clasp |
| 2026-10-12 | `コード.js`／`dry_b5_propose_virt.py`／B5要件 | **得たい結果は1つ。料理は一枚。届けるはプレゼント**: 門の語は見出しだけ出さない。届ける案は料理を空にして場面に包みか手渡し。fitが無い食べ方案だけ、ヒットを親に作り方を変える。戻し: Git revert。要出品 clasp |
| 2026-10-12 | `コード.js`／`dry_b5_propose_virt.py`／B5要件 | **キャッチの原則**: 一枚で得たい結果が見える。文として完結しなくても、イメージが止まればよい。戻し: Git revert。要出品 clasp |
| 2026-10-12 | `コード.js`／`dry_b5_propose_virt.py`／B5要件 | **訴求を先に**: 困り・得たい結果・訴求理由は同じ依頼で揃える。キャッチとセールスはその三つから書く。書いたあとに材料の語は足さない。戻し: Git revert。要出品 clasp |
| 2026-10-12 | `コード.js`／`dry_b5_propose_virt.py`／B5要件 | **使い方の料理名**: 料理名そのものだけの使い方は空。文の途中にある料理名では消さない。戻し: Git revert。要出品 clasp |
| 2026-10-12 | `コード.js`／`dry_b5_propose_virt.py`／B5要件 | **買う人が先**: 年齢の幅・世帯・役目を先に書く。使い方は材料・販売文・公式ページから選び、足りない案だけ困りから書く。材料の3語では埋めない。使えなければ1回やり直し、2回目も空ならログ。画面のペルソナは買う人。固定の一覧は使わない。戻し: Git revert。要出品 clasp |
| 2026-10-12 | `コード.js`／`dry_b5_propose_virt.py`／B5要件 | **空の料理・文字切れ・同語・訴求の理由**: 置ける候補を先に使い、空の案には矛盾しない余りだけ回す。書いたあとの削りと語の抜き取りをやめる。同じ語が残ったらその案だけ作り直し、残ってもログだけ。訴求の理由の枠復唱は空。戻し: Git revert。要出品 clasp |
| 2026-10-12 | `コード.js`／`dry_b5_propose_virt.py`／B5要件 | **料理は案ごとに判定**: 訴求と商品名とその回の候補を1回見て、隣に置けるものだけ残す。料理名の固定一覧では決めない。商品名に料理名があるときは落とさない。戻し: Git revert。要出品 clasp |
| 2026-10-12 | `コード.js`／`dry_b5_propose_virt.py`／B5要件 | **キャッチは訴求から**: 文は材料の言い換えに限らない。材料に無い効能・累計・受賞・No.1は書かない。同じ語は同商品の他案で再使用しない。ペルソナ詳細に訴求の理由。おつまみには食事料理を付けない。戻し: Git revert。要出品 clasp |
| 2026-10-12 | `コード.js`／B5要件 | **確定済みは案へ直行**: 正本と個品見本が全親で確定済みなら、B-④を開いた時点で案を作り②へ進む。戻し: Property `B5_AUTO_TO_PROPOSE=false`。要出品 clasp |
| 2026-10-12 | `コード.js`／B5要件 | **ペルソナが空の調査ログ**: 応答が配列か、各枠が残ったか落ちたかを実行ログに出す。中身の直しは次のログのあと。戻し: Git revert。要出品 clasp |
| 2026-10-12 | `コード.js`／`dry_b5_propose_virt.py`／B5要件 | **B-④の4点**: 使い方の語を先に案を分ける。料理名は1回だけ、尽きたら空。公式URLはリンク。ペルソナ6枠は必ず埋め、他案と被らない。固定の年齢層は使わない。戻し: Git revert。要出品 clasp |
| 2026-10-12 | `コード.js`／`dry_b5_propose_virt.py`／B5要件 | **訴求はペルソナの順**: 使い方を1案1つ。商品名とメーカー名は用事にしない。誰に／何をのあとでキャッチ。前の案と同じ方向は禁止。戻し: Git revert。要出品 clasp |
| 2026-10-12 | `コード.js`／B5要件 | **食品列は見出しの「カテゴリー」と「食品」**: 改行があってもその列を読む。セルに「食品」があるとき食品。空の amazon カテゴリーへ落とさない。戻し: Git revert。要出品 clasp |
| 2026-10-11 | `コード.js`／`dry_b5_propose_virt.py`／B5要件 | **字数は画面＝焼く文。ペルソナは材料から。料理名はキャッチ外**: 固定の人物文を出さない。材料に無い枠は空。戻し: Script Property `B5_PERSONA_FROM_MATERIALS=false` で人物文の割当に戻る。要出品 clasp |
| 2026-10-11 | `コード.js`／`dry_b5_propose_virt.py`／B5要件 | **料理はマスタ先・訴求度は自己採点**: マスタの食べ方はURLなし可。足りなければ公式。門落ちの型差し替えをやめる。戻し: Git revert。要出品 clasp |
| 2026-10-11 | `コード.js`／B5要件 | **公式URLの落ち先をログ化**: | ▼ログ(システム用) に why（hit／empty_body／no_token／redirect_unresolved／weak_only／no_candidate）。検索は仮想と同じ2回。リダイレクトは本文とcanonicalで実URLを拾う。戻し: Git revert。要出品 clasp |
| 2026-10-11 | `コード.js`／B5要件§0・§2 | **virt仕様の穴**: キャッチ見出し＝用事（手軽／おつまみ連打止め）。公式URLはパス揺れ・検索URL展開・生存確認。ペルソナは固定6行と明記。戻し: Git revert。要出品 clasp |
| 2026-10-11 | `コード.js`／B5要件§0／HUMAN_RUN | **S15〜S19 GAS（virt揃え）**: ペルソナ詳細スロット・場面Gemini・公式URL料理プール・用事適合。戻し: Git revert。要出品 clasp |
| 2026-10-11 | `コード.js`／B5要件§0.1／HUMAN_RUN | **S15〜S17 GAS**: 画面＝焼成・生完結優先・セールス【】除去・画像用別欄廃止。戻し: Git revert。要出品 clasp |
| 2026-10-11 | `コード.js`／`dry_b5_propose_virt.py`／B5要件§0／HUMAN_RUN | **案A親キュー＋公式URLキャッシュ**: 1窓・親ごと完了→次・先読み・累積指示。公式URLはJANキャッシュ（TTL30日・弱URL除外）。戻し: Git revert。要出品 clasp |
| 2026-10-11 | `dry_b5_propose_virt.py`／B5要件§0・§0.1・§0.2／HUMAN_RUN | **S19＋公式ページ食べ方**: URL汎用解決・本文抽出・用事適合・公式再利用。一件URL焼かない。戻し: 当該差分revert。clasp不要（GAS未） |
| 2026-10-11 | `dry_b5_propose_virt.py`／B5要件§0・§0.1・§0.2／HUMAN_RUN | **S15〜S18 正本**: 画面＝焼成・生完結優先・セールス【】禁止・気分用事は料理任意・モール販売URL除外。戻し: 当該差分revert。clasp不要（GAS未） |
| 2026-10-11 | `dry_b5_propose_virt.py`／B5要件§0・§0.1・§0.2／HUMAN_RUN | **画面＝焼成＋字数フォールバック＋ソース階層**: 希望6〜12／4〜10（不成文のみ〜18／〜12）。転換禁止。料理＝材料→公式Google、競合既定オフ。戻し: 当該差分revert。clasp不要 |
| 2026-10-11 | `dry_b5_propose_virt.py`／B5要件§0・§0.1・§0.2／HUMAN_RUN | **S15途切れ禁止・S16添える料理URL・場面ばらけ（virt＋docs）**: Gemini検索で実在レシピ＋https。個品見本≠料理。包材は後ろ。GAS未。戻し: 当該差分revert。clasp不要 |
| 2026-10-10 | `コード.js`／`b5_main_base_compose.py`／`b5_virt_abc.py`／virt／B5要件 | **S10〜S14＋salesBurn GAS**: 入口ゲート廃止・同一用事・セールス画像用枠・composeが catchBurn＋salesBurn。戻し: Git revert。要出品 clasp |
| 2026-10-10 | `b5_copy_struct_gates.py`／`dry_b5_copy_struct_gates.py`／B5要件§0.1 | **S10〜S14 仮想ゲートPASS**: 同一用事／用事が先／場面／入口ゲート／大＋小。戻し: 当該差分revert。clasp不要 |
| 2026-10-10 | B5要件§0・§0.1・§0.2・§3.2／§3.5／§3.7／HUMAN_RUN | **役割定義ロック（docs）**: 入口ゲート廃止・同一用事・用事が先（N1以外）・大＋小焼成・S10〜S14。戻し: 当該 docs 差分。clasp不要 |
| 2026-10-10 | `dry_b5_propose_virt.py`／B5 HUMAN_RUN | **案画面仮想**: マスタ出品CK親→7案相当（子SKU数）・ペルソナ短文／詳細・catch／sales／catchBurn・訴求度作り直し予算3。戻し: 当該ファイル削除。clasp不要 |
| 2026-10-10 | `コード.js`／B5要件§0.1／HUMAN_RUN | **§0.1 S1〜S9 GAS配線**: ミックス両構成門・全案セールス・ペルソナ詳細UI＋整文プロンプト・意味単位catchBurn・場面編集・レ点のみ再作成・旧ペルソナ禁止文削除。戻し: Git revert。要出品 clasp |
| 2026-10-10 | B5要件§0.1／`b5_copy_struct_gates.py`／`dry_b5_copy_struct_gates.py` | **合格文・fixture汎用化**（構成・区別語。非魚ミックス例＋魚は検証例）。戻し: 当該差分revert。clasp不要 |
| 2026-10-10 | `b5_copy_struct_gates.py`／`dry_b5_copy_struct_gates.py`／B5要件§0.1 | **S1〜S9 仮想ゲート**（画像/GASなし）。ゲート識別の仮想PASS。GAS実装は別承認。戻し: 当該ファイル削除／§0.1からS5〜S9削除。clasp不要 |
| 2026-10-10 | B5要件 §0.1／HUMAN_RUN | **画像前・文案構造の合格文**（S1〜S4）。仮想PASS後にコード。課題は表へ随時追加。戻し: 当該節削除。clasp不要 |
| 2026-10-10 | `.claspignore`／`code-trunc-probe/size_guard.ps1` | **切れ対策の穴埋め**: `_tmp_*.js` を push 除外。見張りは縮小後に基準サイズ更新（証拠連打防止）。戻し: ignore行削除／旧ps1。clasp不要（ignoreは次回pushから） |
| 2026-10-10 | `コード.js`／B5要件 | **B-④再定義コード**: 件数＝レ点子（上限12）・ペルソナばらけ・catchBurn確認欄・戻る／レ点再作成・食品料理AI補完・ミックス両方特徴。焼成は catchBurn 優先。戻し: Git revert。要出品 clasp |
| 2026-10-10 | B5要件／HUMAN_RUN | **3者条件つき docs 反映**: 提案12〜36／焼成6〜18を確認画面の catchBurn 欄（AI短縮→人チェック）。案＝子SKU数・採用画像1〜4。自動訴求度3回と戻る／レ点再作成は別枠。料理は材料不足時LLM補完（外部Web APIなし）・本文可。汎用注記。コード別承認。戻し: 当該 docs 差分を戻す |
| 2026-10-10 | B5要件／HUMAN_RUN | **9/14以降合意の再定義（docsのみ）**: ペルソナあり・料理AI可で階段／10/6を明示上書き。戻る・レ点再作成・件数＝子SKU・分岐・個品主のみ・ミックス縦横比を §0 固定。3者との関係を階段節に記載。コードは別承認。戻し: 当該 docs 差分を戻す |
| 2026-10-10 | `コード.js`／B5要件 | **目的の階段（一部コード）**: 用事が先・味最後・訴求度画面・味だけT4なし・作り直し最大3回。戻し: Git revert。要出品 clasp |
| 2026-10-08 | FLOW §3 | **次の段階を I へ**。見積検品の見方はロック。G の自動取得は3者まで始めない。戻し: §3 の焦点を 2026-10-03 の文へ戻す。clasp不要 |
| 2026-10-08 | ① `見積検品`／見積検品要件 | **見方を要件化**。大カテゴリーで販売数1つ。入数はその味だけ。条件列は増分。見積リンクは2026年の添付がある1通。戻し: 要件ファイルと FLOW の当該行を消す。clasp不要 |
| 2026-10-06 | `コード.js` | **B-④セールストーク案**: 件数はレ点子SKU（2未満なら4、上限12）。ミックスは両方の特徴。場面を1〜3文で画面に出す。戻し: Git revert。要出品 clasp |
| 2026-10-06 | `コード.js`／`b5_main_base_compose.py`／`b5_virt_abc.py`／B5要件 | **B-④指示をロック文へ**。食品列とセット構成をJSONへ。シートに mix があるときは商品名の中黒で構成を足さない。戻し: Git revert。要出品 clasp |
| 2026-10-05 | `b5_main_base_compose.py`／`b5_virt_abc.py`／`food_unit_canon.py` | **組み合わせMAIN**: 中黒の商品名から単品2種を付け、確定文は残して「食べ方」と「両方の袋・個品」を場面へ足してから画像化する。戻し: Git revert。clasp不要 |
| 2026-10-04 | `コード.js`／ALIGN §7.12.3 | **B-④①で正本をDriveサムネ選択**（`b5TruthThumbPickerHtml_`）。B5本体をバックアップから復元＋サムネ。戻し: Git revert。要出品 clasp |
| 2026-10-04 | `package_truth.py`／`sub_image_food_canon_gate.py`／ALIGN §7.12.3 | **①′でも袋正本を画像選択**（01同置き場→parent_shared）。戻し: Git revert。clasp不要 |
| 2026-10-04 | ① `見積検品` | **商品キー=単品のみ**（MIXは構成単品へ紐づけ）。見送り行グレー。分析結果を結論・見通し・モール別に再定義。戻し: シートクリア。clasp不要 |
| 2026-10-04 | ① `見積検品` 石原全商品キー | **推奨に想定利益**、書式を旧1行形式へ。かつお/まぐろ/ぶり/MIX/チーズ系など7行。戻し: シートクリア。clasp不要 |
| 2026-10-04 | ① `見積検品` かつお50g行 | **セット別○×を表形式**、ASINに●点、月販をAmazon/楽天/Yahoo1列、仕入可否の右に分析結果列。戻し: 当該行削除。clasp不要 |
| 2026-10-04 | ① `見積検品` かつお50g行 | **改行区切り・ASINはtextFormatRunsでクリック可・推奨を本線/参考/条件時に分割**。月販はAmazon「●点以上」優先、Keepa50は参考と明記。戻し: 当該行削除。clasp不要 |
| 2026-10-03 | ①スプシ `見積検品` | **仕入可否提案表**（1商品1行）。かつお50g試し入れ。利益○=セット≥100円。戻し: シート削除。clasp不要 |
| 2026-10-03 | 仕入検討品番リスト BX–CJ（石原） | **経路別を横並び拡充**: FBA=最安・仕入・手数料・送料・保管・利益・状態／FBM=最安・仕入・手数料・自己配送・利益・状態。戻し: BX–CJ削除。clasp不要 |
| 2026-10-03 | 仕入検討品番リスト BX–CC 見出し色・表示形式 | **FBA塊オレンジ／FBM塊水色**。数値 `#,##0_);▲[Red]#,##0`。列順を FBA→FBM にまとめ。戻し: 書式クリア＋旧交互列順。clasp不要 |
| 2026-10-03 | 仕入検討品番リスト BX–CC／BC（石原・仕入あり） | **FBA/FBM最安＋経路別セット利益を表示**。倉庫Keepa（再GETなし）。BCは V→標* で送料マスタ。戻し: BX–CCと当該BCを消す。clasp不要 |
| 2026-10-03 | FLOW §5.3.2／PHASE 0re／HANDOVER | **FBA/FBM最安は見積付与後だけ追加取得**。洗い・門では取らない。戻し: 当該行を消す。clasp不要 |
| 2026-10-03 | 仕入検討品番リスト AW（石原） | **メール単価通し拡大**: チーズ190/220→615、小袋おだし50/35/40系→324。1kg・150/200g・チーズ+おだし混在・他社は空。戻し: 当該AWを消す。clasp不要 |
| 2026-10-03 | FLOW §3／§5.3.2／§5.3.2bis／PHASE 0re | **見積焦点**: メール単価→利益が先。nn左レイアウト・自己発送は仮。戻し: 当該段落を消す。clasp不要 |
| 2026-10-03 | 仕入検討品番リスト AW（チーズ11行） | **石原ボリューム単価へ訂正** 669→615（570税抜×1.08切下）。FLOW§5.3.2に2単価の優先を追記。戻し: AWを669へ。clasp不要 |
| 2026-10-03 | 仕入検討品番リスト BV–BW／AX・AY・BU／出品 `00_設定マスタ` | **利益に販売手数料・保管料を反映**。BV/BWへ円を書き込み（横断式なし）。AX=売−FBA−販売−保管（1点）。大口・割引プロモなし。販売手数料を2026帯＋保管単価行を設定マスタへ。戻し: BV–BW削除・AX/AY/BUを旧式へ・設定マスタの販売手数料行を旧料率へ。clasp不要 |
| 2026-10-01 | 仕入検討品番リスト BS–BU | **利益の状態を3列追加**。寸法と重さが揃う行だけ手数料を補完（4行）。寸法も重さもない478行は手数料不明のまま。戻し: 3列を消す。clasp不要 |
| 2026-09-29 | `photo_realism_rules.py`／`b5_bag_aspect.py`／`b5_virt_abc.py`／`food_unit_canon.py` | **包材は盛り付けの後ろ。窓以外は不透明。比は同じ面の器だけ。縁だけ柔らかく、影は下端に沿う** | Git revert。clasp不要 |
| 2026-09-28 | `b5_bag_aspect.py`／`b5_main_base_names.py`／`b5_food_look_read.py`／`b5_main_base_compose.py`／`sub_image_food_canon_gate.py` | **許容は最大3%。セールスは案ごとに別の訴求。面の縦横が入力にあるときだけ食器と比べる。1片は主見本と違うと描き直し。見本画面にマスタのサイズを出す** | `B5_BAG_ASPECT_TOL` は 0.03 以下。面の縦横は JSON の packageFaceHcm／Wcm。Git revert。clasp不要 |
| 2026-09-27 | `PurchaseResearch.js`／CYCLE §4.4 | **洗える小中が3社以上のあいだは④**。200以上でも店に入らない。店は3社未満かつ200以上の1回。列挙は余力なし | `PR_SHOP_TOKEN_GATE=false`／Git revert。要 clasp（① `purchase_research`） |
| 2026-09-27 | `PurchaseResearch.js` | **数字だけの8〜14桁はメーカー名にしない**。抽出・サマリ追記で落とす。既にある行は次の起床で消す | Git revert。要 clasp（① `purchase_research`） |
| 2026-09-27 | `PurchaseResearch.js`／CYCLE §4.4 | **余力なしでは名取りを始めない**。61以上は④。200以上だけ店1軒。完了店の抽出メーカーをサマリに足す。小中0の次は③ | `PR_SHOP_TOKEN_GATE=false`／Git revert。要 clasp（① `purchase_research`） |
| 2026-09-20 | `PurchaseResearch.js`／CYCLE §4.4 | **店作業中の印を続き後に残さない**: 名前取り続きで25分門閉じを廃止。次の起床は正本の分岐へ | Git revert。要 clasp（① `purchase_research`） |
| 2026-09-19 | `PurchaseResearch.js`／`competitor_jitta/JittaResolve.js`／CYCLE §5 | **当たり直しを競合GASへ**: ①はC2と小規模出現を待たない。競合の別時計が全店再走査。足し算禁止 | `PR_HIT_RECOUNT_OWNER=hb`／`CJ_HIT_ENABLED=false`／Git revert。要 clasp（①＋`competitor_jitta`） |
| 2026-09-27 | `b5_bag_aspect.py`／`b5_virt_abc.py`／`photo_realism_rules.py` | **包材の輪郭はファイルの縁。2つ以上は高さの対比も検査** | Git revert。clasp不要 |
| 2026-09-27 | `food_unit_canon.py`／`b5_food_look_read.py`／`photo_realism_rules.py`／`b5_virt_abc.py` | **個品の1片は主見本。副は盛り付け。袋の影は食品と同じ** | Git revert。clasp不要 |
| 2026-09-27 | `photo_realism_rules.py`／`b5_bag_aspect.py`／`b5_virt_abc.py`／`b5_main_base_compose.py` | **ピントは同じ距離。縦横は袋ファイルの縦÷横。外れたら差を書いて描き直し** | `B5_BAG_ASPECT_TOL`。Git revert。clasp不要 |
| 2026-09-27 | `b5_virt_abc.py`／`food_unit_canon.py`／`b5_main_base_compose.py` | **組み合わせ**: 親袋は大きさの見本。描く袋は単品。貼らない | Git revert。clasp不要 |
| 2026-09-20 | `b5_main_base_compose.py` ほか | **見出し検査・見本読み・種分け・写真感2回目** | Git revert。`--photo-pass none`。clasp不要 |
| 2026-09-20 | `b5_virt_abc.py`／`food_unit_canon.py` | **個品は見本優先**: 厚み・色・質感＝主見本。形・サイズ＝副見本。キャッチは文字だけ | Git revert。clasp不要 |
| 2026-09-20 | `b5_virt_abc.py`／`food_unit_canon.py` | **MAIN指示汎用化**: 袋①②・お茶請け廃止・見本に無い形・原料は背景のみ・品番を絵に入れない | Git revert。clasp不要 |
| 2026-09-20 | `b5_layer_stamp.py`／`b5_virt_abc.py`／`food_unit_canon.py` | **後載せ不合格で既定オフ**: 袋はAI配置・縦横比固定。個品は貼らない。短冊許可削除。厚切りは絵の題材から外す | `B5_STAMP_TRUTH_LAYERS=1` で旧実験／Git revert。clasp不要 |
| 2026-09-19 | `b5_layer_stamp.py`／`b5_main_base_compose.py`／ALIGN §7.14 | **MAIN袋・個品は正本画素後載せ**: AIは皿・背景のみ。複数は同一縮尺 | `B5_STAMP_TRUTH_LAYERS=0` または `--no-stamp-layers`／Git revert。clasp不要 |
| 2026-09-19 | `PurchaseResearch.js`／`cycle_heartbeat.py`／T-CHB-36〜39／CYCLE §4.4・§7 | **心拍は洗い優先**: 余り1軒を洗いの前に置かない。余力なし≤60／中61–199／あり≥200。レ点1日上限なし・1回1店。列挙1回の次は洗い。店印はC2より前。余力なしでは大を洗わない | `PR_TOKEN_AMPLE`／`PR_CYCLE_AUTO_CK_DAILY_CAP=8`／Git revert。要 clasp（① `purchase_research`） |
| 2026-09-19 | `コード.js`／`b5_virt_abc.py`／B5要件 | **B-④キャッチ**: 短いフック。手順禁止。空欄は根拠で戻す。N1は名前を切らない。成果は材料。件数は子SKU（2未満なら4） | Git revert。要出品 clasp |
| 2026-09-19 | `PurchaseResearch.js`／`schema.py`／セラーP/Q | **セラーに規模帯・名取り日をP/Qへ**: 既存の規模帯を左へ。名取り日は巡回日と同じ。見出し名で読むので列ずれしない | Git revert。要 clasp（①）。シート挿入は実施済 |
| 2026-09-19 | `コード.js`／ALIGN §7.15 | **B-④正本**: 紐付け確定で `処理済み/{親SKU}_単体` へコピー。Canva元は `01/_取り込み済み` へ移動。人が01を掃除しない | Git revert。要出品 clasp |
| 2026-09-19 | `コード.js`／B5要件 | **ミックス整文**: セット構成≥2で顧客体験（試す・気分・場面・人がいる場・贈る）。材料に無いお得・最強は禁止 | Git revert。要出品 clasp |
| 2026-09-18 | `コード.js`／B5要件 | **T1〜T4**: 「食べるおだし」固定禁止を廃止。繰り返さないのは商品名ベースとシリーズ名（括弧より前） | Git revert。要出品 clasp |
| 2026-09-18 | `コード.js` | **B-④セールストーク案**: 欠落していた `amazonAiLoadAiGenreLookup_`／`wnExact_` を追加。AI情報取得dataを読んで案材料にする | Git revert。要出品 clasp |
| 2026-09-18 | 食材見本ゲート／compose／ALIGN §7.15 | **組み合わせ袋の段**: 最上段＝セット袋。構成の袋は親SKU／商品名ベース。JANフォルダの袋を単品に流用しない | Git revert |
| 2026-09-18 | `コード.js` | **B-④ Drive全件なめ廃止**: 正本は01直下＋保存 rel のみ。色・正本は共有JSON。処理済み／過去ラン全走査しない | Git revert。要出品 clasp |
| 2026-09-18 | `コード.js` | **B-④親の一意キー**: JANではなく親SKU。同じJANの別親を落とさない。親SKU空のときだけJAN | Git revert。要出品 clasp |
| 2026-09-18 | `コード.js`／`Yahoo.js` | **B-④欠落本体**: 出品CKは `masterCheckboxIsTrue_`（true/TRUEのみ）。正本初期選択は保存済み親SKU／商品名ベースのみ。文字色は1〜5 | Git revert。要出品 clasp |
| 2026-09-19 | `PurchaseResearch.js`／`schema.py`／セラーP/Q | **セラーに規模帯・名取り日をP/Qへ**: 既存の規模帯を左へ。名取り日は巡回日と同じ。見出し名で読むので列ずれしない | Git revert。要 clasp（①）。シート挿入は実施済 | |
| 2026-09-18 | `PurchaseResearch.js`／`cycle_heartbeat.py`／T-CHB-34／CYCLE §4.4 | **洗える小中があるとき C0 しない**: 余り1軒と同じ。レ点も名前取りも無ければ残洗い | Git revert。要 clasp（① `purchase_research`） |
| 2026-09-18 | `PurchaseResearch.js` | **大特大の未済キューは10社**: 超過未済を書き戻し削除。列挙の再流入停止。小中・Keepa済は残す | `PR_WASH_QUEUE_XL_PENDING_MAKERS_MAX=0`／Git revert。要 clasp（① `purchase_research`） |
| 2026-09-18 | `PurchaseResearch.js`／`dry_cycle_seller_accum.py`／T-CSA-07／FLOW §5.4／CYCLE §3 | **C0 http 飛ばし**: 新品オファー200以外は済印＋次へ。①ログ `C0 skip_http` に番号と相手文。429は従来どおり停止 | `PR_C0_SKIP_HTTP=false`／Git revert。要 clasp（① `purchase_research`） |
| 2026-09-18 | `c1_packaged.py`／`c1_bulk_template_required.py`／T2 | **バルク列の約束**: ブラウズはプルダウン実値。親にサイズ・在庫・親番号を書かない | Git revert |
| 2026-09-18 | `.cursor/rules/dev-generic-cause-plain.mdc`／AGENTS.md | **開発ルール**: 汎用・原因先。分かったら根拠短記＋対策まで。短い返答 | ルールファイル削除／Git revert |
| 2026-09-18 | `packaged_soup_and_stew_column_map.json`／`shelf_registry.json`／T2 HUMAN_RUN | **B-T2 SOUP**: 指紋＋列マップを棚登録。かつお／まぐろ PACKAGED を 03 へ | Git revert。次回 E③ 手渡しは SOUP も B-T1 可 |
| 2026-09-18 | `export-tax-list/Code.gs`／再開手順 | **輸出許可書取り込み**: ロック／Invoice重複のみ／移動修正。メニュー「フォルダからデータ取り込み」（常用は第1階層、メンテは「その他」）。処理対象年で `令和N年[度]/輸出許可書` を解決。切れは resumeCurrentJob | Git revert 後 `export-tax-list` で clasp push || 2026-09-18 | `コード.js`／`AmazonSpapiPut.js`／PUT docs | **PUT max_items 既定 10→50**、ハード上限 50→200（Property 新規不要。30件相乗りを1回で通す） | `APPROVAL_AMAZON_SPAPI_PUT_MAX_ITEMS=10`／Git revert。要出品 clasp |
| 2026-09-18 | `コード.js`／`AmazonSpapiPut.js`／`AmazonApprovalExport.js`／P0／PHASE | **ALLOW_MASTER_QTY 未設定=ON**: Property に true を足さない。明示 false だけ停止 | 明示 `false`／Git revert。要出品 clasp |
| 2026-09-17 | `コード.js`／D新規棚ゲート | **E③ 09取込**: 棚なしは専用案内。09のxlsmをbrowseIndexへ。2回目E③は選択画面（自動実行しない） | `AMAZON_SHELF_INGEST_09_ENABLED=false`／Git revert。要出品 clasp |
| 2026-09-17 | `コード.js`／要件 §6.1／TITLE_WORKAREA | **AI-DESC-FISH-MIX**: 同期は一意ASIN→一意商品名ベース＝AI商品名→一意JAN（メーカー両方非空なら一致）。ブランドはJOINキーにしない。生成プロンプトは1SKU固定。未設定=ON | `AI_DESC_FISH_MIX_JOIN_ENABLED=false`／Git revert。要出品 clasp |
| 2026-09-17 | `AmazonCategoryPt.js`／P4b承認／HUMAN_RUN | **貼付◎で分類多数決**。無いときはログ◎→マスタ他店→JAN複数→商品名。同点は空欄にしない。自ASINは単品除外 | `APPROVAL_AMAZON_P4B_PASTE_CIRCLE_VOTE_ENABLED=false`／Git revert。要 clasp |
| 2026-09-17 | `コード.js`／`AmazonApprovalExport.js` | **Script Properties 上限**: B5／PT指紋をDrive退避。Z 15-㉓〜㉕。秘密はシートに書かない | Git revert。要 clasp |
| 2026-09-16 | SESSION／HANDOVER §8.2／CURRENT_PHASE | **Yahoo部分更新を駐車**: getItem＋現行値同送は未着手。実装しない | docs revert |
| 2026-09-16 | `AmazonDriveImageExport.js`／`コード.js`／U4 HUMAN_RUN | **U4ドレイン**: 未充足全件。切れはE②開始から22分。残りはトリガー再開（上限12波）。MAX_SKUSは収集に使わない | **Git**: revert。要 clasp。旧 `AMAZON_U4_SLICE_MS=270000` は削除 |
| 2026-09-16 | `.cursor/rules/full-auto-no-repeat-menu.mdc`／`full-auto-no-human-type.mdc`／AGENTS.md | **全自動**: 同じメニューの件数分割再実行を提案禁止。切れは機械再開 | docs／rules revert |
| 2026-09-16 | `AmazonDriveImageExport.js`／U4 HUMAN_RUN | **U4対象収集**: 上限20は未充足SKUから。済み先頭20の後ろが永久漏れしない | **Git**: revert。要 clasp |
| 2026-09-16 | `コード.js`／`AmazonImageMatrixExport.js`／CDE要件 | **E①**: 楽天未整理非読。一致=楽天欄保持／不一致=重なりSKUだけ引継ぎ。切戻し `AMAZON_E1_KEEP_RAKUTEN_ENABLED=false` | **Git**: revert。要 clasp |
| 2026-09-16 | `コード.js` onOpen E下層 | **E①②③** 接頭辞（文言本体そのまま） | **Git**: revert。要 clasp |
| 2026-09-16 | `tools/set_main_image/transparent_bg.py`／HUMAN_RUN | **マット②**: 白不透明は外周リングのみ。白袋ラベルの誤NG回避 | **Git**: revert |
| 2026-09-16 | `build_yahoo_caption_themes.py`／`YahooCaptionRich.js`／要件 §0.1 | **紹介から内容量数字を除外**（単位不問）。大容量は残す。種に入れず機械で削る | Git revert。要出品 clasp |
| 2026-09-16 | `build_yahoo_caption_themes.py`／`YahooCaptionRich.js`／要件 §0.1 | **P6c 紹介組立**: catch＋salesTalk／subs＋魚種・内容量・主原料。写真用overviewと商品説明文AIは使わない。その後文字数合わせ | Property OFF／Git revert。要出品 clasp |
| 2026-09-16 | `build_yahoo_caption_themes.py`／`YahooCaptionRich.js`／要件 §0.1 | **P6c 文字数合わせ**: PC・スマホ同じ仕様。既存 overview／subs だけ短縮（`overviewSp`）。超える欄は GAS が落として P5 に戻す | Property OFF／Git revert。要出品 clasp |
| 2026-09-16 | `build_yahoo_caption_themes.py`／`YahooCaptionRich.js`（P6C_PACK）／要件 §0.1 | **P6c pool＝マスタ楽天サブ画像1〜10**: 入っている枠が当選。楽天CDN直載せなし | Property `YAHOO_CAPTION_RICH_P6C_THEMES_ENABLED=false`／Git revert。要出品 clasp |
| 2026-09-15 | `PurchaseResearch.js`（`prCycleOneShopYieldEnabled_`／`prPickWashContinueTarget_`／`prNextHeartbeatAction_`／`prRunAlightHit_`）／`cycle_heartbeat.py`／CYCLE §4.4 | **余り時は店1軒→すぐ洗い**: キュー未済を続きにしないので残り2補充が届く。Keepa中は止めない。予約60超かつ洗える小中>0なら1軒（小・中の店を先）。直後は洗い必須。特大件数は増やさない。ログ `one_shop_yield` | `PR_CYCLE_ONE_SHOP_YIELD=false`。**Git**: `git revert`。要 clasp（①） |
| 2026-09-15 | `upload_lib_images_api.py`／`Yahoo.js`（`uploadLib`） | **P6b**: 公式 uploadLibImage で maguro/buri 6枚UP成功・公開200。Pro手UP不要 | スクリプト再実行／Git revert。要 clasp（uploadLib） |
| 2026-09-15 | `YahooCaptionRich.js`／`prep_odashi_lib_bodies.py`／P6b HUMAN_RUN | **P6b**: 魚種libファイル自動差し込み（UP検知）。素材は theme_new_full。Pro手UP6枚 | Git revert。要 clasp＋Pro UP |
| 2026-09-15 | `YahooCaptionRich.js`／P6a HUMAN_RUN | **P6a欠画FIX**: 楽天CDNをcaptionに載せない。Yahooホスト／katsuo libのみ。要再API反映 | Git revert。要出品 clasp |
| 2026-09-15 | `YahooCaptionRich.js`／`Yahoo.js`／P6a HUMAN_RUN | **P6a**: 本文画像を楽天サブ優先で自動差し込み。Property `YAHOO_CAPTION_RICH_P6_IMG_ENABLED`（未設定=ON） | Prop `false`／Git revert。要出品 clasp |
| 2026-09-15 | 要件 §6.1 | **AI-DESC-FISH-MIX メモ**: 単品SKUなのにAI商品説明が複合魚種。Yahoo explanation 誤表示。改善は別タイミング | docs revert |
| 2026-09-15 | `YahooCaptionRich.js`／`Yahoo.js`／P5 HUMAN_RUN | **P5**: おだし共通骨格へ魚種・内容量・セット（兄弟SKU）自動埋込。テンプレ列空でもJAN/名で解決。チーズかつおは除外 | Git revert。Property OFF。要出品 clasp |
| 2026-09-15 | `docs/org/YAHOO_CAPTION_TEMPLATE_LOCK_REQUIREMENTS.md` 他 | **P5再定義（docsのみ）**: 手製第2テンプレ→共通骨格＋変数自動埋込。コード未着手 | docs revert |
| 2026-09-14 | `YahooCaptionRich.js` | **説明API反映で他欄消失FIX**: Property ON時は `updateItem`（全マッピング＋テンプレ注入）。最小ペイロードだと headline/abstract/explanation が空上書き | Git revert。要 clasp。消えたSKUはメニュー再実行 |
| 2026-09-14 | `YahooCaptionRich.js` | **P3 ed-00006対策**: 反映中ロック時に待機リトライ＋件間sleep。要出品 clasp | Git revert |
| 2026-09-14 | `Yahoo.js`（`updateItem` テンプレ注入）／要件 | **案A'**: 省略LOCK廃止。Property ON＋テンプレ解決時は caption/FS に生成HTMLを同送（Yahoo「省略＝デフォルト上書き」対応） | Property `false`。Git revert。要出品 clasp |
| 2026-09-14 | `YahooCaptionRich.js` | **API反映FIX**: editItem に name/path/product_category/price を同送（装飾のみだと it-01011 等で400） | Git revert。要出品 clasp |
| 2026-09-14 | `YahooCaptionRich.js`／`コード.js`（Yahooメニュー）／要件・HUMAN_RUN | **Yahoo説明案A**: 手貼り本線廃止。テンプレ組立→メニュー「Yahoo説明をAPI反映」で editItem（caption/FS）。通常出品はLOCK。再生成 `_gen_yahoo_caption_rich_js.py` | メニュー非使用／Git revert。要出品 clasp |
| 2026-09-14 | `Yahoo.js`（`yahooCaptionRichLockIsActive_`） | **Yahoo説明LOCK簡略**: Property ON だけで出品CK対象全件の caption／FS を送らない。親ごとの LOCK レ点不要 | Property `false`。要出品 clasp |
| 2026-09-14 | `Yahoo.js`（`yahooCaptionRichLock*`／`updateItem`） | **Yahoo説明LOCK（P4）**: Property `YAHOO_CAPTION_RICH_LOCK_ENABLED`＋親 `▼マスタ(Yahoo説明LOCK)` 時、editItem から caption／additional1/2／スマホFSを送らない。explanationは送る | Property `false`／LOCK外し／Git revert。要出品 clasp |
| 2026-09-13 | `PurchaseResearch.js`（`prSmNeedShopRestock_`／`prNextHeartbeatAction_`／`prAlightShouldShopBeforeLarge_`／`prCycleDeferLargeContinueForShop_`）／`cycle_heartbeat.py`／CYCLE §4.4・§7 | **小中は残り2で店補充・1日40店の絶対上限を外す**: 3以上ある間は店に商品 Keepa を使わない。2以下は jitta→店→残洗い（0になってからだと洗えない1社で止まる）。40は1実行1店＋予約60と重複。ログ `sm_restock`／`restock_wash`／`sm_pool=` | `PR_SM_RESTOCK_BELOW=1`（0になってから店）。**Git**: `git revert`。要 clasp（①） |
| 2026-09-13 | `PurchaseResearch.js`（`prShopHasTokenRoom_`／`prAlightShouldShopBeforeLarge_`／`prCycleAutoCkLeft_`）／CYCLE §4.4 | **店の日次10を「トークンが予約60より多い間は続ける」に**: 小中が空でも満タン放置しない。洗いは約10件分を残す。ログ `tokenRoom=` `left=` | `PR_SHOP_WHILE_TOKEN=false`。**Git**: `git revert`。要 clasp（①） |
| 2026-09-13 | `PurchaseResearch.js`（`prEnsureMakerSumBandGrayCf_`） | **メーカーサマリの 大／特大 行を薄いグレー**（条件付き書式 `#eeeeee`。規模帯列が 大 or 特大）。心拍と規模帯書込で付け直す | 条件付き書式を手で消す／Git revert。要 clasp（①） |
| 2026-09-13 | `tools/yahoo_page_mock/paste/*` | **かつお1袋 caption仮貼付用HTML作成**（反映済CTA URL・セット2袋〜全件MAIN）。戻しは BACKUP_BEFORE_PASTE.html。人間がProで貼付 | バックアップ再貼付 |
| 2026-09-13 | `tools/yahoo_page_mock/*`／Drive `04.クーポン画像` | **導線修正**: セットボタン→セットコーナー→MAIN全件。CTA画像をDriveへ保存 | Git revert |
| 2026-09-13 | `tools/yahoo_page_mock/*`／`assets/yahoo_coupons/README.md` | **セット誘導をメイン画像全件リンクに変更**。専用セットCTA廃止。売れ筋バナーのみ残す。SETS_BLOCK自動生成は後段 | Git revert |
| 2026-09-13 | `tools/yahoo_page_mock/*`／`assets/yahoo_coupons/cta_gourmet_*.jpg` | **CTAレイアウトA確定**: 全幅縦積み。文言「お得なセットを見る」「当店の売れ筋を見る」（ストアTOP）。単位なし。左右並びBは見送り | Git revert |
| 2026-09-13 | `tools/yahoo_page_mock/*`／`assets/yahoo_coupons/cta_gourmet_bestsellers.*` | **商品説明CTAをセット＋当店の売れ筋（ストアTOP）に変更**。クーポン埋め込みは廃止。モック表示更新 | Git revert |
| 2026-09-13 | `PromoRemind.js`／`コード.js`（Yahooメニュー）／`tools/yahoo_page_mock/README_DEPLOY.md` | **販促クーポン期限リマインド**: シート「▼販促施策リマインド」に2四半期分シード（Q3〜10/1 0時・Q4〜1/1 0時）。日次メール（B17）。発行APIは無し。要 clasp（出品） | Prop `PROMO_REMIND_ENABLED=false`／トリガー削除／Git revert。要 clasp |
| 2026-09-13 | `PurchaseResearch.js`（`PR_WASH_ST_OUT_`／`PR_WASH_ST_FIX_`／`prWashStSplitEnabled_`／`prClassifyWashEnd_`／`prWashShouldRunC4_`／`prMigrateWashHoldSplit_`／`prMigrateWashHoldSplitDecide_`／`prSetMakerWashStatus_`／`prSanitizeEmptyResearchStuck_`／`prCycleClearFalseComplete_`／`prCycleRewashFalseCompletes_`／`prDispatchHeartbeat_` 入口／洗い完走 C4）／CYCLE §6.10・HUMAN §10.3tredec | **洗い状態の終わりを3区分にする**: 保留を捨て、完了（測り終わり。通過0可・90日で戻る）／調査対象外（title0・非仕入先。戻さない）／要修復（①候補0行。人が見る）。完了＋通過0を自動で列挙に戻さない。C4 は完了かつ通過≥1。移行は初回心拍1回（完了には空の `最終調査日` へ今日を入れる）。人がシートを触らない。node `_tmp_wash_st_split_check.js`。切戻し `PR_WASH_ST_SPLIT=false` | Property `PR_WASH_ST_SPLIT=false`（旧: 完了＋通過0を保留）。再移行は `PR_WASH_ST_SPLIT_MIGRATED` を消す。**Git**: `git revert`。要 clasp（①） |
| 2026-09-13 | `PurchaseResearch.js`（`prMakerSummaryRowJittaRetryable_`／`prJittaRetryNonHighEnabled_`／`prJittaMarkRetryExhausted_`／`prJittaRetryExhaustEnabled_`／`prJittaHbDailyMax_`・`prJittaHbDailyCount_`・`prJittaHbDailyBump_`・`prJittaHbDailyHasRoom_`／`prCountMakerJittaUninvestigated_`／`prCollectJittaResolveTargets_`／`prApplyJittaMirrorToSummary_`／`prResolveMakerJitta_` の cache-heal と AI 両経路／`prCycleHeartbeatDecide_` の `jittaPend`／`prDispatchHeartbeat_` の jitta 分岐）／CYCLE §6.9・HUMAN §10.3duodec | **新メーカーが「要確認」で行き止まり、洗える材料が増えなかった**: 「AI済み」の定義が2つあった。`prMakerNeedsJittaResolve_` は `prMakerJittaSummaryAiConfirmed_`（信頼度 === `high`）で見るのに、`prMakerSummaryRowUninvestigated_` は「信頼度が空でなければ対象外」で見る。`prCollectJittaResolveTargets_` が両方を AND で見るため、**信頼度 medium／low の 要確認・未調査 28社**（要確認 medium 10／low 16／未調査 medium 2）が「再調査が必要なのに対象に数えない」＝誰にも引き取られない。プールは `prMakerJittaWashEligible_`（製造メーカー／入口ブランドのみ）なので永久に洗えない。実測: `action=jitta_resolve` の最終実行は **2026-09-04 7:47**、以後9日間 `jittaPend=0` で1度も動いていない（9/13 の心拍は alight 32回・enum_unknown 16回・wash_med 4回）。**本体の害は `prApplyJittaMirrorToSummary_` も同じ判定で書き込みをスキップしていたこと**で、メーカーマスタに `実態候補=製造メーカー` があってもサマリに映らない社ができていた（実測 **株式会社マッセル／株式会社K&K／高橋商店／笑喜亭　タカヒラ** の4社。AI もトークンも不要でミラーするだけで洗える。11:20 に大を呼び込んだのがマッセル／K&K＝§6.8 と同根）。**「日次上限だけで空転を止める」案は誤り**だった: キャッシュ経路（`prMakerMasterJittaCached_`）は AI を呼ばないので、回数を制限しても同じ値を書き戻すだけで1社も解決しない（28社中7社がキャッシュ、21社はマスタに行なし＝AI を呼ぶ）。修正は再調査専用の判定を新設し、**範囲は信頼度が high でなく かつ 実態が 未調査／要確認／空 の行だけ**（人が入れた 製造メーカー／入口ブランド／非仕入先 は触らない。レガシーメニュー `prMenuMigrateJittaUnchecked_` は旧判定のまま）。打ち切りは**結果が出たときだけ**メーカーマスタ `実態再採取しない` を立てる（`no_resolve`＝429・タイムアウトでは立てないので次の回でやり直す。復活は人がフラグを外す）。安全弁として1日 `PR_JITTA_HB_DAILY_MAX`（既定8）回で jitta レーンを降りる（フラグ書き込みが失敗しても心拍が `smN===0 && jittaPend>0` で居座らない。居座ると 2026-09-13 03:00〜04:20 の1.5時間停止と同じ壊れ方）。node 実コード検証 **24/24 PASS**（実測の組み合わせ5件／触らない行4件／切戻し2件／打ち切り8件／安全弁4件／日付リセット1件）。ログの目印 `jitta_retry_exhausted`・`jitta_hb_daily_cap`・`heartbeat jitta` の ` round=n/8` | Property `PR_JITTA_RETRY_NON_HIGH=false`（旧判定）／`PR_JITTA_RETRY_EXHAUST=false`（打ち切らない）。**Git**: `git revert`。要 clasp（①） |
| 2026-09-13 | `PurchaseResearch.js`（`prAlightShouldShopBeforeLarge_`／`prAlightShopRoundsUntilCapEnabled_`／`PR_ALIGHT_SHOP_ROUNDS_UNTIL_CAP_PROP_`／`prWashContinueDeferInfo_` の defer 文言）／CYCLE §6.8・HUMAN §10.3undec | **大・特大の発火条件が逆だった**: 「セラーリサーチしても新メーカーが出なくなったら大」が正なのに、`prAlightShouldShopBeforeLarge_` 末尾の `lastNew === 0 && rounds < roundMax` は**新メーカーが1社でも出たら店ラウンドを閉じて大へ行く**動きだった。新メーカーはサマリに 実態=未調査／要確認 で入るため `prMakerJittaWashEligible_`（製造メーカー／入口ブランドのみ）を通らず選定プールに入らない＝**「見つかったが洗えない」のに店が閉じる**。実測 2026-09-13 11:20:58: `done=6`（目標5超）・`lastNew=2`（10:44のA-lightで2社）→ false → 未調査の店が21軒あるのに `cont=キーコーヒー via=queue allow_xl sm=0 unk=0 large=28 band=大`（見積306件≒487トークン≒回復1.6時間）が週枠を消費して開始。その2社は **株式会社マッセル／株式会社K&K** で実態=未調査・規模帯=空・見積=空＝洗えない。11:42 に `new_makers=0` が出た 12:01 からは正しく `defer_shop_before_large` に戻っていた。修正後は「見つかった数」で閉じず日次上限まで店を回す（呼び出しは `smN===0` の文脈のみ）。**空転しない**: `done>=maxD`（既定10）と `rounds>=roundMax`（既定10）はどちらも日次カウンタで必ず止まる。§4.4 の小→中→大→特大 の順は変えない。node 実コード検証 **10/10 PASS**（11:20再現／切戻しで旧動作／日次上限で止まる／ラウンド上限で止まる／目標未達は続ける／`lastNew=-1` 旧バグ非再発／A-light続きは常に続ける／`FALSE` 大文字／`true` 明示）。ログの目印 `shop_rounds_until_cap`（①ログ の `deferred_skip`） | Property `PR_ALIGHT_SHOP_ROUNDS_UNTIL_CAP=false`（新メーカー数で閉じる旧動作）。**Git**: `git revert`。要 clasp（①） |
| 2026-09-13 | `PurchaseResearch.js`（`prCycleShopBeforeLargeAction_` をキャッシュ用ラッパ化／本体を `prCycleShopBeforeLargeActionRaw_` に分離／`prShopBeforeLargeCacheEnabled_`／`prShopBeforeLargeCacheClear_`／`prGasStart_`／`prDispatchHeartbeat_` 入口／A-light 直後／`prPickWashContinueTarget_` の `deferredDetail_`）／CYCLE §6.6bis・HUMAN §10.3dec | **店の件数読みが pick を9分にして A-light を殺していた**: 小中が0になると `prWashContinueDeferInfo_` が `defer_xl_sm_remain` で即 return せず `prCycleShopBeforeLargeAction_({})` に入る。引数なしだと中で `prSellerUncheckedCount_`（別スプシ `セラー` を openById して全読み・キャッシュなし）＋`prSellerAutoCkCandidateCount_`（同じ `セラー` をもう1回）＋`prCandUntaggedGateCount_`（`①候補` 26,009行）が毎回走り、これがキュー未済ループから社ごとに呼ばれる。未済は大28＋特大77＋中1＝約106社なので**1回の pick で外部読みが200回超**。実測 2026-09-13: 11:52（`sm=0`）を境に pick が 96〜129秒→**470〜627秒**。pick＋pool で実行枠25〜30分を使い切り、**A-light（所要21分）が 12:09・12:51 とも途中で殺され**（実行時間上限なのでエラーが残らない）、`shops_ck_on=21` が59分間1つも減らず、洗いは 11:19〜13:08 で**0件**・`liveLeft=300` 満タンのまま毎時300トークン（約189件ぶん）を捨てていた。**洗いキュー7.5万行は主犯ではない**（キュー読みは1回・1〜2秒。§6.7 は別問題）。引数なし呼びを実行内1回にし、心拍は `unchecked`/`autoCand`/`untagged` を実数で渡すので非キャッシュ＝判定精度は落とさない。捨てるのは `prGasStart_`・心拍入口・**A-light 直後**（`巡回日`・`採取元` が変わる）。§4.4 の優先順は変えない。node 実コード検証 **13/13 PASS**（106回呼び→外部読み1回／心拍呼びは非キャッシュ／捨てたら読み直し／`false` で旧動作／0軒で null／再入なし）。ログの目印 `shopCacheHit=`（①ログ の `deferred_skip`）・`shop_before_large_cache` | Property `PR_SHOP_BEFORE_LARGE_CACHE=false`（毎回読み直す旧動作）。**Git**: `git revert`。要 clasp（①） |
| 2026-09-13 | `PurchaseResearch.js`（`prCycleHeartbeatDecide_` の `shop_after_sm`／`shop_after_large` 分岐／`prCycleShopAfterSmOnceEnabled_`／`prCycleShopAfterSmOnceTag_`）／CYCLE §6.6・HUMAN §10.3sept | **完走後の店レーンを本当に1回だけにする**: §6.3bis 案B の意図は「Keepa完走直後に店レーンを1回挟む」だが、`prCycleClearShopAfterSm_` が自動レ点・A-light・C0 の3つとも0になるまで呼ばれないため、店の未調査が残る間ずっと `shop_after_sm` を返し続けていた。その間は心拍が `wash_small` / `enum_unknown` に到達せず、小・中が1社でも残ると `defer_xl_sm_remain` で大26社・特大76社も止まり **Keepa レーンが丸ごと停止**。実測 2026-09-13 03:00〜04:20 は `liveLeft=300`（満タン）のまま洗い0件・1.5時間・約200件ぶんの取り逃し。`pick` も毎回109〜197秒かけて `via=deferred_all parent=`（やることなし）を返していた。1つ返す直前に印を消す。店レーンの本体は `smN === 0` の `prCycleShopBeforeLargeAction_` 側なので消化は止まらない。node 実コード検証 6/6 PASS。ログの目印 `shop_after_sm_once`／detail の ` once=Y` | Property `PR_CYCLE_SHOP_AFTER_SM_ONCE=false`（尽きるまで居座る旧動作）。**Git**: `git revert` |
| 2026-09-11 | `PurchaseResearch.js`（`prCatalogSearchPage_`／`prCatalogRankCount_`／`prCatalogRankProbeEnabled_`／`prCatalogRankCutEnabled_`／`prCatalogNoRankTag_`／`prCatalogEnumerateQueryToQueue_`／`prEnumerateMakerToQueue_`／`prEnumerateSplitWordsToQueue_`）／CYCLE §6.5・HUMAN §10.3sex | **列挙で Amazon の順位有無を記録する（切るのは次段）**: 門の落ち 17,209件のうち `stats空` 4,769（27.7%）＋`順位欠測` 1,487（8.6%）＝**33.8%** は「Keepa に90日の価格も順位も無い」＝判定できずに落ちた分で、1件1トークン払って捨てている。SP-API 実測（400件・読取のみ・`tools/purchase_research_path3/_tmp_salesranks_check.py`）で **門通過→Amazon順位なし 1%／`stats空`→99%／`順位欠測`→96%**。順位の「値」は使えない（Amazon はサブカテゴリ順位・通過の中央値156位に対し `順位>60000` の93%も6万位以内で分離不能）。使えるのは「順位が付いていない」だけ。`includedData=summaries,salesRanks` にして `rankN` を持ち、列挙時に `noRank` を数えて `列挙完了 … noRank=41/120(34%)` としてログに残す。**既定は記録だけでキューには従来どおり入れる**。CUT を ON にすると洗う件数が約34%減り、門通過率 7.0%→10.6%・同じトークンで通過が約1.5倍（14.1→約21件/時）・節約 約6,256トークン＝Keepa約21時間分。失う通過は約1%（約13件）。ログの目印 `noRank=`／`enum main … noRank=` | Property `PR_CATALOG_RANK_PROBE=false`（salesRanks を取らない旧動作）。切りを止めるのは `PR_CATALOG_RANK_CUT` を未設定に戻す。**Git**: `git revert` |
| 2026-09-11 | `PurchaseResearch.js`（`prCycleRunEnumBatch_`／`prCycleEnumStockN_`／`PR_CYCLE_ENUM_STOCK_N_DEFAULT_=3`）／CYCLE §6.4oct・HUMAN §10.3quin | **在庫を作る: まとめ列挙を「洗える社が3社たまるまで」続ける**: 旧は1社見つかった時点で洗いへ譲るため在庫が常にゼロで、1社洗い終わるたびに列挙からやり直し。実測（09-11 00:49〜08:05）で社の切替が中央値 **22.4分・窓の75%**（予約1.5分＋選定2分＋心拍6分＋予約1分＋選定2分）。3社ためれば続く2回の切替は心拍を丸ごと飛ばせる。Catalog は1周の社数が増えるぶん周回数が減るので総消費はほぼ不変。予算切れは従来どおり `prCycleEnumBatchHasBudget_`。ログの目印 `enum_batch_stock`／`stop=queue_found=… stock=3/3` | Property `PR_CYCLE_ENUM_STOCK_N=1`（1社で譲る旧動作）。**Git**: `git revert` |
| 2026-09-11 | `PurchaseResearch.js`（`prWashQueuePendingMap_`／`prWashQueuePendingCount_`／`prWashQueuePendingMapEnabled_`／`prWashQueueCountCacheClear_`／`prWashQueueAppendAsins_`／`prWashQueueMarkKeepaDoneByAsins_`／`prGasStart_`）／CYCLE §6.4oct・HUMAN §10.3quin | **選定の高速化: 未処理件数を1回読みの表にする**: `prWashQueuePendingCount_` が社ごとに `prWashQueueSnapshot_`（キュー全26,238行）を呼ぶため、候補8社を試す選定で8回スキャンしていた。実測 pick 中央値 106〜115秒（続き `via=sum` は15秒）。①候補の `prCandPassSnapshot_` と同じ「1回読んで全社ぶん数える」方式に統一。スライス中の社は既存スナップショットを優先し、追記・Keepa済マーク時は表を差分更新（作り直すと全読みに戻る）。ログの目印 `queue_pending_map rows=… elapsedMs=` | Property `PR_WASH_QUEUE_PENDING_MAP=false`。**Git**: `git revert` |
| 2026-09-10 | `PurchaseResearch.js`（`prRunMakerWashFullBody_` の `needEnum` 分岐／`prWashStSyncEnabled_`）／CYCLE §6.4sex・HUMAN §10.3quat | **効率1: サマリに書いた洗い状態を実行内の `st` にも反映**: キューに未処理を入れて「Keepa中」を書いた直後も手元の `st` が古いままだったため、`isKeepaCont` が false になり洗える相手なのに `catalog_first_other` で1サイクル譲っていた。譲り先はこのレーンでは処理されないので、次サイクルで同じ社を洗い直すだけの空振り。2026-09-10 07:00〜18:25 の実測で25回・約44分 | Property `PR_WASH_ST_SYNC=false`。**Git**: `git revert` |
| 2026-09-10 | `PurchaseResearch.js`（`prScheduleWashFullTrigger_`／`prWashYieldToHeartbeat_`／`prRunWashFullFromTrigger` の shop_busy／`prHbYieldArmNoShortenEnabled_`／`prWashTriggerTagIsHbYield_`／`PR_HB_YIELD_ARM_MIN_=5`）／CYCLE §6.4sex・HUMAN §10.3quat | **効率2: 心拍譲りの予約を心拍の実測所要に合わせる**: 心拍（規模確定）は実測約4分かかるのに予約は3分で、さらに満タン短縮（`PR_WASH_ARM_FULL_CAP`）が掛かって2分になっていた。心拍の途中で洗いが発火し `shop_busy` で空振り→3分後に再予約。実測36回・約108分。`hb_yield*` / `shop_busy_wait` タグは短縮対象外にし、予約を5分にする（トークン待ちではなく前の仕事の待ちなので短縮は誤り） | Property `PR_HB_YIELD_ARM_NO_SHORTEN=false`（3分＋短縮ありの旧挙動）。**Git**: `git revert` |
| 2026-09-10 | `PurchaseResearch.js`（`prCandGatePassCountForMaker_`／`prCandPassSnapshot_`／`prCandPassSnapClear_`／`prCandPassByQueueEnabled_`／`prGasStart_`／`prUpsertCandidate_`）／FULL §4.3・HUMAN §10.3ter | **対策A: 門通過を自社キューの ASIN でも数える**: カタログ照合が部分一致なので社名が包含関係にある社（415社中53組。中村食品 ⊂ 中村食品産業）が互いの商品を取り込み、①候補は 1ASIN=1行の後勝ちなので先に洗った社が合格を奪われて「通過0」に見えた。実例=`B096T9YHWD`（商品名は中村食品産業だがメーカー列は中村食品）。メーカー列で0件のときだけ自社キューを照合。ログの目印 `cand_pass_by_queue` | Property `PR_CAND_PASS_BY_QUEUE=false`。**Git**: `git revert` |
| 2026-09-10 | `PurchaseResearch.js`（`prCycleClearFalseComplete_`／`prFalseCompletePendingOnlyEnabled_`）／CYCLE §6.4・HUMAN §10.3ter | **対策B: 偽完了救済で Keepa中 に戻す条件を「未処理>0」だけに**: 旧条件 `queueN > 0 \|\| pending > 0` は Keepa済も含む全件を見ていたため、未処理0の社が Keepa中 に戻り「洗えない続き相手」になった。全件洗い済みで候補0は再列挙せず保留（Catalog も消費しない）。ログの目印 `empty_research queue=… all_done` | Property `PR_FALSE_COMPLETE_PENDING_ONLY=false`。**Git**: `git revert` |
| 2026-09-10 | `PurchaseResearch.js`（`prRunMakerWashFullBody_` の Keepa 直前／`prWashPhantomStrike_`／`prWashPhantomClear_`／`prWashPhantomClearIf_`／`prWashPhantomFixEnabled_`）／CYCLE §6.4・HUMAN §10.3ter | **対策C: 幽霊の続き相手を自動解消**: 洗い状態=Keepa中 なのに未処理0 の社は洗いレーンに毎回選ばれるが洗えず、他レーンへ譲って自分を再予約するだけなので心拍の列挙レーンに永久に順番が来ない（2026-09-09 04:01→09-10 05:44 の25時間・549回・切片0）。連続2回で完了へ書き戻す（候補0なら `prSetMakerWashStatus_` が保留へ振り分け）。**手作業でのシート書き換えを不要にする**。ログの目印 `phantom_keepa` / `phantom_keepa_resolved` | Property `PR_WASH_PHANTOM_FIX=false`。**Git**: `git revert` |
| 2026-09-10 | `PurchaseResearch.js`（`prEffJudgeAfterArm_`／`prEffStallGuardEnabled_`／`prEffMinSinceLastSlice_`）／HUMAN §10.3ter | **対策C（監視）: 切片が出ていない空転を 効率合否 で FAIL に**: 間隔合否 `prEffJudgeAfterSlice_` は切片が出たときしか動かないため、1件も洗えない25時間が全て PASS だった。最終切片から60分以上なら `no_slice_stall` FAIL。予約自体の合否より優先 | Property `PR_EFF_STALL_GUARD=false`。**Git**: `git revert` |
| 2026-09-09 | `Yahoo.js`（`YahooApiClient._post`）／`HANDOVER.md` §4.5.1 | **Yahoo出品の一時通信失敗を限定再試行**: 実測は `editItem` 成功後の `setStock` 1件だけ「使用できないアドレス」、前後29件は正常。通信例外／HTTP 429・5xxだけ最大3回（1秒→2秒待機）。恒常4xxは即失敗し、ログへendpoint・attempt・応答要約を残す。要出品 clasp | `Yahoo.js` の `_post` を単発 `UrlFetchApp.fetch` に戻す／Git revert。要 clasp |
| 2026-09-09 | `PurchaseResearch.js`（`prMakerLegalVariants_`／`prMakerNeedleLegalEnabled_`／`prMakerWashNeedles_`／`prEnumerateMakerToQueue_`／`prRunMakerWash_`）／FULL §7ter・HUMAN §10.3bis | **①社名の法人格ゆれを Catalog 照合語に足す**: `prCatalogTitleHits_` は商品名＋ブランドの `indexOf` 完全一致なので、マスタが `大平水産(株)` だと Amazon の商品名（`大平水産` / `大平水産株式会社`）が全件落ち、タイトル含有0になっていた（Keepa製品ファインダーでも `大平水産株式会社`=209件・`大平水産`=一致なし・`(株)` 表記はどちらにも無い）。表示名から `大平水産`／`大平水産株式会社`／`株式会社大平水産` を自動生成して照合語に追加（表示名は先頭のまま＝既存の突合を壊さない）。検索キーワードは**0件のときだけ**法人格なしで1回だけ言い換え（`legal_retry`）＝現に動いている社の列挙量を変えない。落とした語が3文字未満（`PR_MAKER_NEEDLE_MIN_LEN_`）なら追加しない＝誤ヒット防止。模擬20ケース PASS。要 clasp（①） | Prop `PR_MAKER_NEEDLE_LEGAL=false`（表示名だけで照合する旧動作）。Git revert。要 clasp |
| 2026-09-09 | `PurchaseResearch.js`（`prCycleMakerBandPools_`／`prCycleTitle0PoolSkipEnabled_`）／HUMAN §10.3bis | **②タイトル含有0を1回出した社を選定プールから外す**: 中身ゼロの社は推定件数が空のまま `unk` に居座り、`prWashContinueGuard_` の `defer_xl_unk_remain`（`smN>0` **または** `unkN>0` で大・特大を後回し）で伊藤園・ロッテ等の洗いを止め続けていた。旧は3回連続してはじめて保留。`prTitleZeroMap_` の回数>0で選定・列挙候補の両方から除外し `pool_title0_skip` に出す。成功時は `prTitleZeroClear_` で自動復帰。模擬4ケース PASS。要 clasp（①） | Prop `PR_CYCLE_TITLE0_POOL_SKIP=false`（3回連続まで残す旧動作）。Script Properties の `PR_WASH_TITLE0_FAIL_JSON` 削除で個別復帰。Git revert。要 clasp |
| 2026-09-09 | `PurchaseResearch.js`（`prWashFullKeepaSlice_` の `no_product` 分岐／`prKeepaMissMap_`／`prKeepaMissSave_`／`prKeepaMissDoneEnabled_`／`prKeepaMissDoneAfter_`）／HUMAN §10.3bis | **③Keepa が返さない ASIN を見切って Keepa済にする**: 旧は `no_product` を `doneAsins` に入れずキューへ残すため、同じ ASIN を4分ごとに取り直す無限ループになった（アリアケジャパン 19:33〜20:52 の1時間20分で27回・`wrote=0 left=2 two_stage=2/0`・毎回トークン消費・効率合否 FAIL 20件）。ASIN→回数を `PR_KEEPA_MISS_JSON` に持ち、2回（`PR_KEEPA_MISS_DONE_AFTER`）で Keepa済にして落とす。取れたら回数はリセット＝一時的な取りこぼしを永久化しない。Property は500件上限で古い順に間引く。書き込みはスライス末尾1回。模擬6ケース PASS（27回→2スライスで `left=0`）。要 clasp（①） | Prop `PR_KEEPA_MISS_DONE=false`（キューに残す旧動作）／`PR_KEEPA_MISS_DONE_AFTER` で回数調整。Git revert。要 clasp |
| 2026-09-09 | `PurchaseResearch.js`（`prCycleRunEnumBatch_`／`prCycleEnumBatchStopQueueEnabled_`）／CYCLE §6.3・HUMAN §10.3bis | **④まとめ列挙の打ち切りを「①洗いキューが実際に増えた社」に**: 旧は規模帯が小・中なら打ち切りだが、中身ゼロ（タイトル含有0）の社も小と判定されるため**実測31回中29回が1社目で打ち切られ 1.5社/回**しか進まず、まとめ列挙（`PR_CYCLE_ENUM_BATCH_N=8`）が効いていなかった。さらに材料ゼロのまま洗いへ譲るため空振り選定の往復15分を捨てていた。`queueN>0` を条件に追加し `stop=queue_found=`（旧 `stop=sm_found=`）。中身のある小中では従来どおり止める＝小中優先は崩さない。模擬5ケース PASS（同じ並びで n=1→n=4）。要 clasp（①） | Prop `PR_CYCLE_ENUM_BATCH_STOP_QUEUE=false`（規模帯だけで打ち切る旧動作）。Git revert。要 clasp |
| 2026-09-08 | `PurchaseResearch.js`（`prCycleRunEnumBatch_`／`prCycleCatalogPendingList_`／`prCycleEnumBatchN_`／`prCycleEnumBatchHasBudget_`／`prDispatchHeartbeat_` の `enum_unknown` 分岐）／CYCLE §6.3・§7.1 | **規模帯の確定（Catalog列挙）を1実行でまとめる**: 心拍1回=1社→最大8社（`PR_CYCLE_ENUM_BATCH_N`）。列挙そのものは実測16〜48秒なのに、1社ごとに 洗い→店→心拍 の受け渡しをやり直すため**1社あたり7分**かかっていた（選定1.6分＋yield2分＋店busy2分＋往復1分）。未判定244社の消化が約28時間→**約7時間**。打ち切りは 小中が1社出た（`stop=sm_found`）／予算残5分（`stop=budget`）／対象切れ。1社の例外は飛ばして継続（title0 は保留処理のみ）。`prCycleFirstCatalogPending_` は一覧版の cap=1 に寄せて条件の二重管理を排除。模擬6ケースで打ち切り確認。要 clasp（①） | Prop `PR_CYCLE_ENUM_BATCH_N=1`（心拍1回=1社の旧動作）。Git revert。要 clasp |
| 2026-09-08 | `PurchaseResearch.js`（`prMakerBandFromSum_`／`prMakerBandRowMap_`／`prMakerBandMapClear_`／`prWriteMakerSumBand_`）／FULL §7bis 10 | **規模帯引きの全社ループ内シート読みを撤去**: `prMakerBandFromSum_` を「メーカー→{帯,進捗,見積}」表1回読みに。旧は1社ごとに `prFindMakerSumRow_`（サマリ458行全読み）＋ヘッダ＋当該行の3回読みで、選定はキューに出る社ごとにこれを呼んでいた（模擬120社で **480回→2回**・判定結果は一致）。**目的=洗える社が無いときの選定 2〜11.9分→数秒**（1社見つかる選定は19秒。空振り24回・夜間10.3hの61%が Keepa作業ゼロだった真因）。規模帯書き換え時は `prMakerBandMapClear_()`。要 clasp（①） | Prop `PR_MAKER_BAND_MAP=false`（1社ずつ読み直す旧動作）。Git revert。要 clasp |
| 2026-09-08 | `PurchaseResearch.js`（`prCycleAliveReason_`／`prCycleInflightStaleLimitMs_`／`prCycleInflightStaleFastEnabled_`／`prCycleEnsureLiveness_` のログ）／HUMAN_RUN 強制終了耐性 | **番犬の実行中印を35分基準に**: `inflight` の腐り判定だけトークン回復基準（66分）から30分壁基準（35分＝洗い本体の `PR_WASH_INFLIGHT_STALE_MS_` と同値）へ。GAS は約30分で殺すので35分超の実行中印は生存しえない。**沈黙死の検知が70分→40分**（15分周期のため実測ケースで 01:12→00:42）。予約が残ると `reason=wash_trigger` で生存扱いになるため、この判定が唯一の検知口だった（実測 09/08 00:02 発火→`pick_begin` 沈黙→179分 Keepa作業ゼロ・唯一の FAIL `gap_hole`）。`expect_overdue` は不変。`liveness_reset` に `deadPhase=`／`overdueMin=` を追加。要 clasp（①） | Prop `PR_CYCLE_INFLIGHT_STALE_FAST=false`（旧66分）。Git revert。要 clasp |
| 2026-09-07 | `PurchaseResearch.js`（`prEffJudgeAfterArm_`／`prEffCkptBudgetRuleEnabled_`／`prEffCkptNormalAfterMin_`）／HUMAN_RUN §10.2・FULL §10 | **ckpt の効率合否を予算基準に**: `*_ckpt` は「残り予算+2分」で貼るのが正（`prCheckpointWashFullTrigger_`。同一実行と重ならない安全ネットなので短縮しない。満タン2分上限も ckpt 除外）。判定側だけ固定15分の古い前提で、正常な afterMin 15〜22分を WARN `ckpt_long_net`／FAIL `ckpt_full_bucket_long_wait` にしていた（**実害の空き時間はゼロ＝誤検知でログが埋まるだけ**。実測7件すべて PASS へ）。予算+2+3（既定25分）以下は `ckpt_safety_net` PASS、超過**かつ満タン**だけ `ckpt_long_despite_tokens` WARN。ckpt 以外の予約・overlap の判定は不変。要 clasp（①） | Prop `PR_EFF_CKPT_BUDGET_RULE=false`（旧 固定15分）。Git revert。要 clasp |
| 2026-09-07 | `PurchaseResearch.js`（`prListExistingAsinRows_`／`prListDiffUseFormulaEnabled_`）／FLOW §5.2 | **品番リストの毎回書き直しを停止**: 差分判定で数式セルを数式として比べる。`getValues` は `=HYPERLINK(...)` を「メイン」で返すため、画像URLがある行が**常に差分あり**になり、洗いスライスごとに削除→行5再挿入していた（実測 `refresh=4 skip_same=1` を80秒おき。**人が手入力した列が毎回消えていた**）。要 clasp（①） | Prop `PR_LIST_DIFF_USE_FORMULA=false`。Git revert。要 clasp |
| 2026-09-07 | `PurchaseResearch.js`（`prKeepaWashGateProducts_`／`prKeepaProductStats90Fetch_`／`prKeepaHasOffersInfo_`／`prFullNeedsD1Enrich_`／`prEnrichKeepaForPassAsins_`／`prWashFullKeepaSlice_`）／FULL §5.1・5.2 | **Keepa GET を2段ゲートに**: 門は `stats` だけで判定できる（出品者数は `stats.current[11]`。offers は門に使っていない）ので、1段目は offers なし＝1トークン/件。通過（実測3%）だけ offers=20 で取り直す。**目的=1件あたり約5→約1.2トークン**。落ちも薄い行で Keepaフルに残す（GETなし再門の資産）。通過で2段目が429なら Keepa済にせず次スライスへ。D1補強に「通過なのに BuyBox 2列とも空」を追加（実測 通過782件中5件・20件/回上限）。要 clasp（①） | Prop `PR_KEEPA_TWO_STAGE_GATE=false`（従来の一発 offers=20）／`PR_D1_ENRICH_BLANK_BB=false`。Git revert。要 clasp |
| 2026-09-07 | `PurchaseResearch.js`（`prCleanupKeepaFullBlankRows_`／`prKeepaFullAsinRowIndex_`）／FULL §7bis 8・9 | **Keepaフルの1行ずつ読みを撤去**: 空行掃除は ASIN 列1回読み→候補行だけ全列。索引づくりも ASIN 列だけ。**目的=Keepa取得のあるスライス22〜29分→数分**（実データ8,386行に対し毎回8,385回の個別読み取り。1ASINでも22.9分）。要 clasp（①） | Prop `PR_KEEPA_FULL_BLANK_SCAN_BATCH=false`。Git revert。要 clasp |
| 2026-09-07 | `PurchaseResearch.js`（`prWashQueueSnapshot_`／`prWashQueueAppendAsins_`／`prWashQueueDedupeAsins_`／`prWashQueuePendingKeepa_`／`prWashQueueMarkKeepaDoneByAsins_`／`prScheduleWashFullTrigger_`）／FULL §7bis | **洗い本体の全走査撤去＋満タン2分上限の分離**: 1社の処理中はキューを1回だけ読む（列挙の追記は継ぎ足し）。Keepa済は連続行まとめ書き。満タン時の待ち2分上限を `PR_WASH_ARM_FULL_CAP` として効率ループから独立（ckpt は対象外）。**目的=1スライス22〜29分→数分・空白待ちの解消**。要 clasp（①） | Prop `PR_WASH_QUEUE_SNAPSHOT=false`／`PR_WASH_ARM_FULL_CAP=false`。Git revert。要 clasp |
| 2026-09-07 | `PurchaseResearch.js`（`prCyclePoolCounts_`／`prWashQueueUniqCount_`／`prMakerCatalogEstimate_`／`prFinalizeEnumOnlySize_`／`prRunMakerWashFullBody_`）／FULL §7bis | **選定の全走査撤去**: キュー件数は「メーカー→件数」表を1回作って共有。選定プールは1実行1回。**目的=選定24分→数分**（pick 12.8分＋pick後11.4分で予算20分を食い潰し Keepa 到達0だった）。要 clasp（①） | Prop `PR_WASH_QUEUE_COUNT_MAP=false`／`PR_CYCLE_POOL_REUSE=false`。Git revert。要 clasp |
| 2026-09-06 | `competitor_jitta/JittaResolve.js`／`PurchaseResearch.js`／`maker_alias.py`／B_COMPETITOR_JITTA | **法人格ゆれ Phase A**: 同一キーへ実態ミラー・マスタ heal・二重AI防止。行削除なし。要 clasp（競合＋①） | Git revert。要 clasp |
| 2026-09-07 | `appsscript.json` | **`script.scriptapp` 復元**（D一括の `newTrigger` 権限。作業ツリーから欠落→clasp push でGAS側も欠落）。要 clasp＋**再承認** | スコープ削除で戻る。要 clasp |
| 2026-09-07 | `コード.js`／CDE要件 | **C／D／E／Bメニュー復元**（truncate後Git復元で旧IAに戻った分）。bak先頭＋B5完結関数。要 clasp・シート再読込 | bak差戻し／Git revert。要 clasp |
| 2026-09-06 | `コード.js`／マトリクス表示 | **Driveプレビューを thumbnail `sz=w1000`**（`rakutenDrivePreviewImageFormula_`）。実ファイル解像度・Dアップは不変。要 clasp | ヘルパーを `uc?export=view` に戻す／Git revert。要 clasp |
| 2026-09-06 | `コード.js`／SUB_PACK_ASC要件 | **サブ昇順詰め**: N昇順→同Nはファイル名昇順で列1〜10。番号なしは余り枠／候補。`RAKUTEN_MATRIX_SUB_PACK_ASC_ENABLED`。要 clasp | Prop false。Git revert。要 clasp |
| 2026-09-05 | `コード.js`／WORK_MERGE_SCOPE要件 | **C①02+一時合算**＋**②レ点外JANスキップ**（先頭行青枠投入廃止）。Prop `RAKUTEN_MATRIX_MERGE_WORK_FOLDER`／`RAKUTEN_MATRIX_SCOPE_FILTER_ENABLED`。要 clasp | Prop false。Git revert。要 clasp |
| 2026-09-05 | `コード.js`／SKIP_EXISTING要件 | **C取得上限100**＋**マスタ既存URL枠はアップ省略**（`RAKUTEN_UPLOAD_SKIP_EXISTING_MASTER_URL`）。要 clasp | Prop false。PROCESS_LIMIT戻し。Git revert。要 clasp |
| 2026-09-05 | `コード.js`／MATRIX_LAYOUT要件／LEDGER | **楽天マトリクス L1〜L3**: メイン2〜10他セット昇順・サブ全子共有・親=最少セット。JANサブ→親。要 clasp。戻し `RAKUTEN_MATRIX_LAYOUT_EXPAND_ENABLED=false` | Prop false。Git revert。要 clasp |
| 2026-09-05 | `コード.js`／CDE要件／CURRENT／ALIGN／HANDOVER | **C／D／E分割**: C=楽天マッチング。D=楽天Yahooのみ。E=Amazon画像＋出品。D補助廃止→Z。要 clasp | Git revert。要 clasp |
| 2026-09-05 | `コード.js`／B6要件／ALIGN §1.2／SET_MAIN HUMAN／CURRENT／HANDOVER | **B-⑥ Amazon MAIN**: メニュー本線。CからAmazon量産削除・補助化。要 clasp | Git revert。要 clasp |
| 2026-09-04 | `コード.js`／ALIGN §1.1／B3確認／CURRENT／HANDOVER | **B→B-③ 案1**: メニュー増なし。起動確認＋`B_INTEGRATED_INCLUDE_B3_IMAGES`。全パック後に約1分で非対話B-③（Amazon自動）。要 clasp | Prop false。Git revert。要 clasp |
| 2026-09-04 | `コード.js`／`b5_parent_shared.py`／ALIGN | **正本・個品・文字色の完全連携**: B-④で正本初期選択＋個品ラベル。親SKU共有＋Drive `_b5_parent_shared.json`。要 clasp | Git revert。要 clasp |
| 2026-09-04 | ALIGN §1／§3／§1.1／B3確認／SET_MAIN | **文字色＝親共有**（MAIN／サブ順不同）。**B一括＋B-③＝オプション**（未設定OFF・実装後段） | docs のみ。実装は別タスク |
| 2026-09-04 | `c_set_main_rakuten_prompt`／`b5_main_base_review_loop`／ALIGN／SET_MAIN HUMAN_RUN | **弱い統合・親絞り**: 02後C指示＝今の親だけ（`--parent-sku`）。任意実行ボタン。`byChildSku`集計修正。全親は scope=all | Git revert |
| 2026-09-05 | `competitor_jitta/`／`PurchaseResearch.js`／CYCLE／HUMAN_RUN | **競合DB実態AI第一段**: トリガー可・① `PR_JITTA_OWNER=comp` で二重防止。切戻し OWNER=hb／競合トリガー停止。要 clasp（競合フォルダ＋①） | Prop／トリガー。Git revert |
| 2026-09-04 | `PurchaseResearch.js`／`cycle_heartbeat`／CYCLE§4.5bis／HUMAN_RUN | **実態AI分離**: 心拍 `jitta_resolve`（cap8）・A-light末尾デフォルトOFF・未調査≥30で自動レ点抑制。仮想 T-CHB-24〜27（27/27）。切戻し `PR_JITTA_HB_LANE_ENABLED=false`。要 clasp | Git revert。Prop false。要 clasp |
| 2026-09-03 | `PurchaseResearch.js`／watchdog dry／HUMAN_RUN | **店レーン譲り**: 洗い→心拍で inflight clear＋3分安全網＋shop busy。alight 完了で busy 解除＋1分。切戻し `PR_SHOP_LANE_CLEAR_INFLIGHT=false`。仮想 T-CTW-23/24。要 clasp | Git revert。Prop false。要 clasp |
| 2026-09-03 | `PurchaseResearch.js`／watchdog dry／HUMAN_RUN | **P-eff-1/2 効率ループ整流**: 実行中 ckpt 廃止・満タン `afterMin≤2`（`eff_loop_cap`）・WAIT_TOKEN 前 inflight clear・overlap age≥8 即 takeover／短再試行2分。切戻し `PR_WASH_EFF_LOOP_ENABLED=false`。要 clasp | Git revert。Prop false。要 clasp |
| 2026-09-03 | `PurchaseResearch.js`／`eff_judge_report.py`／HUMAN_RUN | **効率合否ログ**: ①ログ `効率合否`（slice/arm/overlap/stale・cause・fixHint）。切戻し `PR_EFF_JUDGE_ENABLED=false`。要 clasp | Git revert。要 clasp |
| 2026-09-02 | `PurchaseResearch.js`／`maker_alias.py`／`cycle_c3_select.py`／`d2_filter.py` | **実態ラベル分離**: 新規行は `未調査`（AI前）。`要確認` は AI後の人判定待ち。`prCollectJittaResolveTargets_` は信頼度空のみ。移行メニュー 99→16。要 clasp | Git revert。要 clasp |
| 2026-09-01 | `PurchaseResearch.js`／`cycle_trigger_watchdog.py`／fixtures T-CTW-13〜18 | **洗い ckpt 自己衝突対策**: ckpt/overlap で `wait_shortened` 無効・overlap 必ず再予約（stale 寄せ）・inflight 中は `trigger_missed` 抑制。切戻し `PR_WASH_CKPT_NO_SHORTEN=false`／`PR_WASH_OVERLAP_STALE_AWARE=false`。仮想18/18。要 clasp＋実機1回 | Git revert。要 clasp |
| 2026-09-01 | `b5_main_base_review_loop`／`b5_virt_abc` measure_grid20 | **100マスガイド**: 文字/装飾マスを薄赤、金丸予約4マス黄。16/24/25 目安 | Git revert |
| 2026-09-01 | `b5_virt_abc`／`b5_main_base_compose`／templates | **20マス厳守**: キャッチサイズ縮小・目標16・retry上限25 | Git revert |
| 2026-09-01 | `sub_image_food_canon_gate`／`sub_image_food_canon_launch`／`コード.js` beginB5AfterTruthBind | **MAIN用見本ゲート2ボタン**（変更なしで閉じる／選び直したらここで保存）。`--context main`。要 clasp | Git revert |
| 2026-09-01 | `b5_main_base_names`／compose `--only-theme`／ALIGN | **salesTalk途切れ修正**（硬切り禁止・だけで後半優先） | Git revert |
| 2026-09-01 | `コード.js`／ALIGN §7.14 | **B-④指示に見本ゲート必須**（finishB5AfterTruthBind）。要 clasp | Git revert。要 clasp（出品） |
| 2026-09-01 | `food_unit_canon`／`b5_main_base_compose`／`b5_virt_abc`／ALIGN §7.14 | **P-align-3**: MAIN が theme_new_run 人間確定 CANON を参照 | Git revert |
| 2026-09-01 | ALIGN §7.14／§1／§5／MAIN要件／CURRENT_PHASE | **個品見本 MAIN↔サブ共有**要件ロック。構図は別。P-align-3 実装未 | docs のみ。実装は別タスク |
| 2026-09-01 | `b5_main_base_names`／`b5_virt_abc`／compose／ALIGN §7.6ter | **MAIN catch大＋sales小**: 階層プロンプト。20マス厳守 | Git revert |
| 2026-08-31 | `b5_virt_abc`／`b5_main_base_compose`／ALIGN §7.6ter | **MAIN Phase1**: scene/catch分離・catchImage短文化・STYLE_TYPO色競合解消 | Git revert |
| 2026-08-31 | `b5_virt_abc`／`b5_main_base_compose`／ALIGN §7.6bis | **MAIN textColor**: `_prompt_p3_free` に accent_text_block 配線 | Git revert |
| 2026-08-31 | `コード.js`／`b5_set_base_map`／compose／dry_sub／ALIGN §7.6bis | **textColor 親ごと**: Drive/マップ自動＋`--text-color-for`。ハードコード1廃止 | Git revert。要 clasp（出品） |
| 2026-08-31 | `PurchaseResearch.js`／CYCLE§6.4quat／HUMAN_RUN | **D2 overflow**: onlyBlank・cap1000・停止メール（復帰手順）・overflow gate・メニュー28・D2後自動C4 | Git revert。要 clasp |
| 2026-08-31 | `mix_review`／`_find_truth`／gate `_truth_for` | **「未選定または失敗」誤表示修正**（人間選定を正しく表示）。正本解決のまぐろ既定を廃止・JANヒント優先 | Git revert |
| 2026-08-31 | `コード.js`／ALIGN §7.13 | **生成不要を正本紐付けへ**: ON時テーマ案Geminiスキップ→指示直行。要 clasp | Git revert。要 clasp（出品） |
| 2026-08-30 | `コード.js`／`b5_skip_generate_finish`／gate／dry_sub／compose／ALIGN §7.13 | **画像の生成不要**: B-⑤ CK→skip指示・既存 export＋レビュー完走。要 clasp | Git revert。要 clasp（出品） |
| 2026-08-30 | `food_unit_canon`／`food_canon_gate`／`dry_sub`／ALIGN §7.12.6 | **IDENTITY／SIZZLE UI**＋副のみエラー＋テーマ mode 参照 | Git revert |
| 2026-08-30 | `コード.js`／ALIGN §7.12.3／RAKUTEN_COURSE HUMAN_RUN | **B-⑤指示先頭＝見本ゲート必須**（確定前に dry_sub/compose 禁止） | Git revert。要 clasp（出品） |
| 2026-08-30 | ALIGN §7.12／`food_canon_manual_ingest`／`food_unit_canon`／gate | **§7.12**: コピペ＋ゼロ案内＋同構図デデュプ（aHash／ランク多様） | Git revert |
| 2026-08-30 | `food_unit_canon`／`sub_image_food_canon_gate*`／ALIGN §7.11 | **見本ゲート**: 候補5〜6・主1＋副0〜2→生成→同画面レビュー | Git revert |
| 2026-08-30 | `food_unit_canon`／`sub_image_mix_*`／`dry_sub_overview_copy_test`／`compose_poc`／ALIGN §7.10–7.11 | **混在段階B**: 独自再生成＋FOOD_UNIT_CANON機械統一 | Git revert |
| 2026-08-31 | `PurchaseResearch.js`／HUMAN_RUN | **A-1**: 店日次カウンタリセット（auto_ck／alight done／rounds／lastNew）。週間大特大枠は残す | Git revert。要 clasp |
| 2026-08-31 | `PurchaseResearch.js` | **自動レ点**: 調査済み（巡回日／alight+shop_full）を除外。誤レ点は A-light で CK 外し | Git revert。要 clasp |
| 2026-08-31 | `PurchaseResearch.js`／`cycle_heartbeat`／CYCLE§4.4 | **根因修正**: 自動レ点を店ラウンド尽きた判定より先に。順=auto_ck→alight。lastNew未設定=-1即終了を修正。T-CHB-23 | Git revert。要 clasp＋A-1 |
| 2026-08-30 | `PurchaseResearch.js`／CYCLE§4.4 | **full_trigger**: 店ラウンド中は `prWashContinueDeferInfo_` で大続きを選定除外（`defer_shop_before_large`）。target=大→shop_first を禁止 | Git revert。要 clasp＋A-1 |
| 2026-08-30 | `PurchaseResearch.js`／`cycle_heartbeat`／CYCLE§4.4 | **実装**: 大を選ばず `prCycleShopBeforeLargeAction_`。大特大完走後は `shop_after_large`（C0→auto_ck→alight）。仮想 T-CHB-14/22 | Git revert。要 clasp＋A-1 |
| 2026-08-30 | CYCLE§4.4・§6.3bis・§7／HUMAN_RUN | **小中=0時は店ラウンド≤10尽くすまで大・特大を選定しない**（pick→defer禁止）。大特大完走後は門通過→セラー→調査 | docs revert |
| 2026-08-30 | `PurchaseResearch.js`／CYCLE§7.5／HUMAN_RUN | **A-1後のCatalog遅延対策**: abort世代・孤児ロック破棄・心拍 other_busy は1分再試行。旧Keepaは打ち切り | Git revert。要 clasp＋A-1 |
| 2026-08-30 | `PurchaseResearch.js`／CYCLE§7.1・§7.5／HUMAN_RUN | **A-1=実行中一時停止＋先頭再選定**: `a1_park`・キュー残し・続き選定スキップ・C4 recover は D1未のみ。番犬 hard_reset は park しない | Git revert。要 clasp＋A-1 |
| 2026-08-30 | `sub_image_mix_*`／`sub_image_review_launch`／ALIGN §7.10 | **混在レビュー段階A**: 独自＋競合1画面・採用export。compose最新は日時優先 | Git revert |
| 2026-08-30 | `sub_image_lp_themes`／`sub_image_ai_compose_poc`／`theme_compose_rules`／`コード.js`／ALIGN §7.9 | **競合本線**: チェック∩写真・1テーマ1枚・想像OFF。袋はパーツ外。スケール／T03／T13／needPackage／核語色 | Git revert。要 clasp（出品ルート） |
| 2026-08-30 | `review_feedback_templates`／`コード.js`／`dry_sub_overview_copy_test`／ALIGN §7.6bis | **サイズ定型＋強調色**: textColorマップ固定・微差可。テンプレ2本追加 | Git revert。要 clasp（出品ルート） |
| 2026-08-30 | `コード.js`／`photo_realism_rules`／`dry_sub_overview_copy_test`／ALIGN §7 | **P-align-2d テーマ別新規本線**: 概要＋訴求＋needPackage。人物詳細・物背景実写。袋は必要時のみ。切戻し `B5_THEME_NEW_PATH=false` | Git revert。要 clasp（出品ルート） |
| 2026-08-30 | `PurchaseResearch.js`／①プロファイル実機 | **食品・飲料・お酒**=調査ON+`57239051`。別名「食品」行削除。seed判定をルート名に合わせる | Git revert。要 clasp |
| 2026-08-30 | `PurchaseResearch.js`／schema／CYCLE§4.1bis | **店規模実装**: ①プロファイル調査カテゴリー＋ルートID。カテゴリ別query見積。A-light OR歩き。セラー内訳列。構成比は規模に使わない | Git revert。要 clasp |
| 2026-08-30 | CYCLE§4.1bis／V1／NOW／FLOW／HUMAN_RUN | **店規模分担**: 調査カテゴリー＝①プロファイル。SKU見積＝カテゴリ別query。構成比は優先のみ。501=単独≥501。在庫サンプル中止。内訳はレ点2以上 | docs revert。実装は別タスク |
| 2026-08-30 | PurchaseResearch.js／CYCLE§7.5／HUMAN_RUN | **A-1店レーン優先**: sm=0で大/特大Keepa続きを保留し自動レ点→A-lightへ。full_trigger shop_first。進捗leftは残す | Git revert。要 clasp＋A-1 |
| 2026-08-30 | `コード.js`／`dry_sub_theme_gate`／ALIGN | **P-align-2b 本線**: 材料門＋Gemini 2文。1・2必須＋ほか最大8＝10。切戻し `B5_PROPOSE_GEMINI_ENABLED=false` | Git revert。要 clasp（出品ルート） |
| 2026-08-30 | `コード.js`／`sub_image_lp_themes`／`sub_image_ai_compose_poc`／ALIGN | **P-align-2b A4**: B-⑤テーマ概要チェック→`--theme-ids-for`。切戻しは指示からフラグを外す（自動スロット） | Git revert。要 clasp（出品ルート） |
| 2026-08-30 | `PurchaseResearch.js`／CYCLE§4／HUMAN_RUN／`dry_cycle_alight_hit` | **店A-light 1店フル**: 小中大はqueryページ＋productスライス続き。特大≥501は浅いhold。日次レ点8・店目標5・上限10・新メーカー0で追加round≤10。心拍 `alight_continue`／`shop_before_large` | Git revert。要 clasp＋止めてから A-1 |
| 2026-08-30 | `sub_image_review_loop`／`sub_image_ai_compose_poc`／`sub_image_review_launch`／ALIGN | **P-align-2a サブUI寄せ**: lightbox・`__rN`・採用再export・目視フォルダ・枚数優先 | Git revert |
| 2026-08-30 | `コード.js`／`compose_set_main`／`b5_set_base_map`／review_loop／ALIGN要件 | **表示B-④＝MAIN／B-⑤＝サブ**。マップ`textColor`→C親ごと。既存3親追記 | Git revert。要 clasp（出品ルート） |
| 2026-08-30 | `b5_main_base_review_*.py`／map／`compose_set_main`／`コード.js`／B5要件 | **B-⑤ P3b 実装**: 複数土台＋レ点子マップ＋既存再利用＋拡大＋再生成世代 | Git revert。要 clasp（出品ルート） |
| 2026-08-30 | B5要件／HUMAN_RUN／CURRENT_PHASE | **B-⑤ P3b 要件ロック**（実装前） | docs revert |
| 2026-08-29 | `b5_main_base_review_*.py`／`コード.js`／B5要件 | **B-⑤ P3 ブラウザ本線**: 写真レビュー・レ点再生成・ラジオ→02。GAS採用UI外し | Git revert。要 clasp（出品ルート） |
| 2026-08-29 | `コード.js`／B5要件 | **B-⑤ P3（旧）**: スプシ採用ボタン＋Drive makeCopy（UI外し・API残置可） | Git revert。要 clasp（出品ルート） |
| 2026-08-29 | `b5_main_base_names.py`／export／compose／`コード.js` | **B-⑤ 比較名=商品名省略＋theme＋キャッチ**。02へコピー CLI。既存ランリネーム | Git revert。要 clasp（出品ルート） |
| 2026-08-29 | `b5_main_base_compose.py`／`コード.js`／B5要件 | **B-⑤ P2**: compose `--role rakuten_main_base`。チェック案JSON→OpenAI＋Pillow金丸/Octas。02未書込。実画像は明示指示時のみ | Git revert。要 clasp（出品ルート） |
| 2026-08-30 | `PurchaseResearch.js`／CYCLE§6.3bis・§7／HUMAN_RUN | **大先・週間枠を続きにも適用・大特大完走後も店レーン**。`defer_xl_large_remain`／`defer_xl_week_cap`。枠・大残りは進捗維持 | Git revert。要 clasp＋止めてから A-1 |
| 2026-08-29 | `PurchaseResearch.js`／CYCLE§6.3〜§7／HUMAN_RUN | **C4→心拍一本化**・**Catalog見積先行**（Keepa新規ゲート・上限メール）・選定から**保留除外**。旧次社 `PR_CYCLE_C4_ARM_MAKER` は成功連鎖でセットしない | Git revert。要 clasp＋A-1 |
| 2026-08-29 | `コード.js`／B5要件 | **B-⑤ M1**: ダイアログで材料→最大4案（門＋Gemini整文）→複数チェック→指示。切戻し `B5_PROPOSE_GEMINI_ENABLED=false` | Git revert。要 clasp（出品ルート） |
| 2026-08-29 | `コード.js`／`b5_main_base_propose.py`／B5要件 | **B-⑤ P1**: メニュー＋正本紐付け＋Cursor指示。Property `B5_RAKUTEN_MAIN_BASE_ENABLED=false` で非表示。切戻し同キー false | Git revert。要 clasp（出品ルート） |
| 2026-08-29 | `PurchaseResearch.js`／CYCLE§6.4ter／HUMAN_RUN | **週次転記チェック**: ①候補`転記チェック`・日曜トリガー・漏れ時D1・社内メール。切戻し `PR_CYCLE_LIST_CHECK_ENABLED=false` | Git revert。要 clasp＋A-1（トリガー武装） |
| 2026-08-29 | `PurchaseResearch.js`／FULL／HUMAN_RUN | **洗い再開待ち=次cap分**（`prKeepaWashResumeWaitMin_`）。負 tokLeft=0。特大の常時61分を解消 | Git revert。要 clasp |
| 2026-08-29 | `PurchaseResearch.js`／CYCLE§6.4bis／HUMAN_RUN | **途中D1**: Keepaスライス(`wrote>0`・未完)ごとに門通過→品番リスト。C4 D1_DONEなし。切戻し `PR_CYCLE_MID_D1_ENABLED=false` | Git revert。要 clasp＋A-1 |
| 2026-08-29 | `b5_virt_abc.py`／B5要件 | キャッチ6〜18字・途中終了禁止。袋名→トーク名は後でメモ | Git revert |
| 2026-08-29 | `b5_virt_abc.py`／B5要件 | **abc16**: キャッチ【】優先・語途中禁止。Octas聖域。上部左寄り。再生成 | Git revert |
| 2026-08-29 | `b5_virt_abc.py`／HUMAN_RUN | **abc15**: 材料提案→最大4案×OpenAI画像。Gemini画像拒否。`--legacy-sets`で旧固定 | Git revert |
| 2026-08-29 | B5要件／HUMAN_RUN | **案F旧文掃除**: 字数事前ゲート／食感マクロ／全案修正を削除。Gemini画像生成しない | docs revert |
| 2026-08-29 | B5要件／HUMAN_RUN | **売れる順=1商品ずつ**: カテゴリ差し替え表なし。点数は試行 | docs revert |
| 2026-08-29 | B5要件／HUMAN_RUN | **案F本線**: 複数チェック→画像複数→1枚採用。根拠薄い案は出さない。T4成果・売れる順。旧案A字数事前修正廃止 | docs revert |
| 2026-08-29 | `b5_virt_abc.py`／B5要件 | **v13**: T3×3＋T4×2。③Pillow Octas右下1.3マス。`compare/*_all_P3b.jpg` | ファイル削除／docs revert |
| 2026-08-29 | `b5_virt_abc.py`／B5要件 | **画像本線=OpenAI**（仮想v12人間ロック）。`--engines` 既定を openai。Gemini画像は比較用 | 既定を gemini,openai に戻す／docs revert |
| 2026-08-29 | `b5_virt_abc.py`／B5要件 | **v12**: ③b/cに主婦向け IMAGE_STYLE_TYPO（99.おしゃれ文字装飾見本）。T3新規30は既存・退避除外 | ファイル削除／docs revert |
| 2026-08-29 | `PurchaseResearch.js`／CYCLE§6.3bis〜§6.6／`cycle_heartbeat` | **案B店挟み**・**大特大週間N=2**・**title0→ブランド追加**・empty_research／偽完了。仮想 T-CHB-20/21。Keepa cap既定30 | Git revert。要 clasp＋A-1 |
| 2026-08-29 | `PurchaseResearch.js` `prCandGatePassCountForMaker_` | **偽完了走査の欠落関数修正**: 存在しない `prCandRowsForListAppend_` → D1同型の `prLoadCandPassByAsin_` | Git revert。要 clasp＋A-1 |
| 2026-08-28 | `PurchaseResearch.js`／CYCLE§7.5／仮想10件 | **フェーズ3**: `PR_WASH_TITLE0_SKIP_SET` 永久skip・選定除外・C4 needs=候補ありのみ・偽完了(`done_no_candidates`/`c4_no_candidates rewash`) | Git revert。要 clasp＋A-1 |
| 2026-08-28 | `PurchaseResearch.js`／CYCLE§7.5／HUMAN_RUN | **title0_skip**（3回→保留＋心拍）・**C4 needs/D1 recover**（完走済み複数の D1 未を A-1 で回収） | Git revert。要 clasp＋A-1 |
| 2026-08-28 | `b5_virt_abc.py`／B5要件 | **v11**: ③b/cスマホ主役サイズ（金丸半分）＋おしゃれ・可読性。Gemini/OpenAI再比較 | ファイル削除／docs revert |
| 2026-08-28 | `b5_virt_abc.py`／B5要件 | **v10**: ③に写真実写＋キャッチ必須1本。Gemini/OpenAI再比較 | ファイル削除／docs revert |
| 2026-08-28 | `b5_virt_abc.py`／B5要件 | **v9**: Gemini/OpenAI同じ指示比較。右上空け→Pillow金丸。OpenAI=gpt-image-2 | ファイル削除／docs revert |
| 2026-08-28 | `b5_virt_abc.py`／B5要件 | **v8**: ③b/c スプシ材料・IMAGE_GOLDをAI配置（後載せなし）・文字20%超は再作成 | ファイル削除／docs revert |
| 2026-08-28 | `b5_virt_abc.py`／B5要件 | **v7**: ③b/c 虚偽禁止・金丸固定・白背景不可・文字20%・材料内売り | ファイル削除／docs revert |
| 2026-08-28 | `b5_virt_abc.py`／B5要件 | **v6**: ③a添付構図固定。③b/cは短い売れる指示のみ | ファイル削除／docs revert |
| 2026-08-28 | `b5_virt_abc.py`／layout_rules／B5要件 | **v5**: ②見出しは帯ルマで黒/白。③人間ベストの対吹き出し＋三本線。P1流用 | ファイル削除／docs revert |
| 2026-08-28 | `b5_virt_abc.py`／layout_rules／B5要件 | **v4**: ①シズル+袋左下傾き+Octas。②③空金丸。③止まる仕掛け | ファイル削除／docs revert |
| 2026-08-27 | `b5_virt_ai_headline.py`／`layout_rules.json`／B5要件 | **v6**: 見本番地をAIへ。下段白。下線半透明重ね。UDモトヤマルベリ | Git revert |
| 2026-08-27 | `b5_virt_ai_headline.py`／`work_paths.py`／B5要件 | **v5**: 98筆下線参照。20マスに下線含む。金丸解決が下線PNGを拾わない | Git revert |
| 2026-08-27 | `b5_virt_ai_headline.py`／`layout_rules.json`／B5要件 | **見出しAI v4**: 金丸をAI参照渡し。100マス上2行。字数12-18/4-8。アンダーライン可 | ファイル削除／docs revert |
| 2026-08-27 | `b5_virt_ai_headline.py`／`layout_rules.json`／B5要件 | **見出しAI v2**: 金丸消し+98後貼り。文字≤20%。色=C `textColors`。02未書込 | ファイル削除／docs revert |
| 2026-08-27 | `b5_virt_ai_headline.py` | **見出しAI仮想** Gemini 4枚。02未書込。Pillowなし | ファイル削除 |
| 2026-08-26 | B5要件 | **案Aロック**。テーマ対最大4→人1枚。Pillowアピール。選定UI。§3.4–3.7 | docs revert |
| 2026-08-26 | B5要件 §3.4 | **画像もテーマ連動複数提案**（メモ→案Aでロック） | docs revert |
| 2026-08-26 | B5要件 | **Gemini4案＋テーマ**。材料＝キャッチ／箇条書き／AI情報取得data。人修正 | docs revert |
| 2026-08-26 | B5要件 | **ハイブリッド承認**。文言材料＝正本＋商品名ベース。AI発明なし | docs revert |
| 2026-08-26 | B5要件 §3.3 | **人フロー**: 文言事前＋画像は review（ハイブリッド） | docs revert |
| 2026-08-26 | `layout_rules.json`／B5要件 | **アピール2段**: name/talk。上段≤18・下段≤10。`appealHeadline1200` | Git revert |
| 2026-08-26 | `layout_rules.json`／B5要件 | **見本金丸はみ出し**: cx=1102, cy=103, d=343。数字番地は非変更 | Git revert |
| 2026-08-26 | `b5_virt_katsuo.py`／B5 HUMAN_RUN | **B-⑤仮想V1–V3**（たべるおだしかつお・空金丸→数字載せ） | ファイル削除 |
| 2026-08-26 | `D_MENU_B5_RAKUTEN_MAIN_BASE_*` | **B-⑤ P0**: 楽天MAIN基礎（右上金丸・数字は後続）。見本 `04.楽天見本` | docs revert |
| 2026-08-26 | `PurchaseResearch.js`／CYCLE§7／HUMAN_RUN | **後回し原因ログ**: sanitize/pool/defer_xl を①ログへ。不明残でも大特大後回し | Git revert。要 clasp |
| 2026-08-26 | `PurchaseResearch.js`／CYCLE§7／HUMAN_RUN | **enumOnly後**: 規模帯記録・Keepa中解除・心拍再arm。小中残存時は大特大キュー続きを後回し | Git revert。要 clasp |
| 2026-08-26 | `PurchaseResearch.js`／CYCLE§6.4／HUMAN_RUN | **C4 pick=0→心拍arm**。メニュー「心拍を今すぐ再開」追加 | Git revert。要 clasp |
| 2026-08-26 | `cycle_heartbeat`／fixtures／`PurchaseResearch.js`／CYCLE§7／HUMAN_RUN | **心拍完成形**: 仮想19/19。トリガー→列挙/A-light/自動CK/C0/大。切戻し `PR_CYCLE_HEARTBEAT_ENABLED=false` | Git revert。要 clasp＋gmail再承認 |
| 2026-08-25 | `tools/common_ss_view/registry.json`／COMMON_SS_VIEW／`common-ss-view.mdc` | **見積発注_三協乾物** 登録（メニュー見積・発注）。戻し: Git revert |
| 2026-08-25 | `PurchaseResearch.js` | **洗い強制終了耐性**: 予算20分・スライスごとckpt再予約・①ログ洗いスライス・lock待ち=残り秒 | Git revert。要 clasp |
| 2026-08-25 | `PurchaseResearch.js` | **洗いトリガー診断ログ**（expectedAt／fire／cleared／trigger_missed／診断メニュー） | Git revert。要 clasp |
| 2026-08-25 | `PurchaseResearch.js` | **取得日時JST**＋**洗いトリガー停滞対策**（未完再武装・lock_busy再予約・rescue） | Git revert。要 clasp |
| 2026-08-26 | `export_sub_images_for_rakuten_matrix.py` | **目視export上書き禁止**（同名は `__composeDir` 別名） | Git revert |
| 2026-08-26 | `sub_image_review_launch.py`／master_sets／B-④案内 | **レ点親→最新compose→ポート連番で複数タブ** | ファイル削除／Git revert |
| 2026-08-26 | B-④／review_loop／`コード.js` | 完了文言: 次の親reviewへ・手export不要。GAS案内も同文言 | Git revert。要 clasp（GAS） |
| 2026-08-26 | B-④／review_loop | 文言: JAN＋商品名・全選択（コピー）・修正なしで終了等 | Git revert。要 clasp（GAS側） |
| 2026-08-25 | `コード.js`／B3 docs | **B-③プレビュー正方形**: 行高＝列幅100px。autoResizeはプレビュー除外 | Git revert。要 clasp |
| 2026-08-25 | `コード.js`／B3 docs | **B-③**: 親のみ・1親ずつ・途中保存＋トリガー再開（`B3_IMAGE_TIME_SLICE_ENABLED`） | Prop=false／Git revert。要 clasp |
| 2026-08-25 | `PurchaseResearch.js` | **S**: 要確認文言（異種セット等）＋赤白／黄セル | Git revert。要 clasp |
| 2026-08-25 | `PurchaseResearch.js` | **S 単品名**: 合格プロンプト（最少名・異種mix・同一JAN寄せ） | Git revert。要 clasp |
| 2026-08-24 | `PurchaseResearch.js` | **S 単品商品名AI充填**（見積候補・既存非上書き・簡易存在）。`PR_SINGLE_NAME_AI_ENABLED` | Git revert。要 clasp |
| 2026-08-24 | `dry_single_name_from_list.py` | **実リスト20件＋AI** 単品名仮想（見積候補・非書・存在は簡易） | ファイル削除 |
| 2026-08-24 | `single_name_classify.py`／fixtures | **単品名パターン仮想 20/20 PASS**（A–F・コーパス限定） | ファイル削除 |
| 2026-08-24 | `docs/org/B_PURCHASE_RESEARCH_SINGLE_NAME_REQUIREMENTS.md` | **単品名・JAN 要件定義**（存在=Catalog・F不一致・コーパスはAI迷い限定） | docs revert |
| 2026-08-24 | `docs/org/B_PURCHASE_RESEARCH_SINGLE_NAME_REQUIREMENTS.md` | **単品名・JAN**: 存在＝Catalog正本・公式不採用・JAN複数パターン再整理 | docs revert |
| 2026-08-24 | `docs/org/B_PURCHASE_RESEARCH_SINGLE_NAME_REQUIREMENTS.md` | **単品名・JAN充填** 要件メモ案（AI先→JAN照合。E前） | docs revert |
| 2026-08-24 | `PurchaseResearch.js`／`d2_filter.py` | **D2**: 大型・特大ティア／重量上限g（①プロファイル・欠測は見送りにしない） | Git revert。要 clasp |
| 2026-08-28 | `PurchaseResearch.js`／liveness／HUMAN_RUN | **予定超過リセット**（Keepa最大回復+5分）・A-1単独復旧・token待ち短縮・完走→C4別実行維持 | Git revert。要 clasp |
| 2026-08-28 | `PurchaseResearch.js`／HUMAN_RUN | **洗いトリガー予約前 Keepa `/token` ログ**（liveLeft vs calcLeft）。待ち妥当性の切り分け | Git revert。要 clasp |
| 2026-08-27 | `PurchaseResearch.js`／HUMAN_RUN | **メニュー整理**: ルート「自動リサーチ」A全自動/B個別/C見積枠/99。A-1即心拍・A-2全診断 | Git revert。要 clasp |
| 2026-08-27 | `cycle_heartbeat_liveness.py`／`PurchaseResearch.js` | **心拍全体の沈黙**: 生存番犬15分・D失敗→心拍・no_maker再武装。仮想15/15 | Git revert。要 clasp |
| 2026-08-27 | `cycle_trigger_watchdog.py`／`PurchaseResearch.js` | **洗いトリガー沈黙**: inflight・phaseログ・二重発火skip・outer catch再武装。仮想12/12 | Git revert。要 clasp |
| 2026-08-24 | `docs/org/B_PURCHASE_RESEARCH_*` | **C5** CYCLE docs 整備（C3+T 本線・実機合格・HUMAN_RUN §10） | docs revert |
| 2026-08-22 | `コード.js`／STAGE | **JANカテゴリ**: 口コミ1位優先。レビュー差&lt;max(10,rc2×0.25)なら多数決。要 clasp | Git revert |
| 2026-08-22 | `コード.js`／STAGE | **Yahoo検索sort分離**: JAN=`-review_count`（口コミ）、文言=`-score`（おすすめ）。要 clasp | Git revert |
| 2026-08-23 | `PurchaseResearch.js` | **C3+T**（C3+完走トリガー1メニュー）・`prArmWashFullTriggerForContinue_` | Git revert。要 clasp |
| 2026-08-23 | `cycle_c4_chain.py`／`PurchaseResearch.js` | **C4完走連鎖**（完走→D1→D2→次社 P3 arm）。`PR_CYCLE_C4_ENABLED` | Git revert。要 clasp |
| 2026-08-23 | `cycle_c3_select.py`／`PurchaseResearch.js` | **C3自動選定**・実態プルダウン・マスタ4社クリアメニュー | Git revert。要 clasp |
| 2026-08-23 | `maker_jitta_resolve.py`／`schema.py` | **Amazonルート＋Browse棚の2列**・業種フォールバック（不明）・dry 6PASS | Git revert |
| 2026-08-23 | `docs/org/B_MAKER_JITTA_CLASSIFICATION.md`／`schema.py` | **実態仕分け正本**（平易版・マスタ/サマリミラー列・AI判定・永谷園/味の素/あさくま事例） | docs revert |
| 2026-08-22 | `PurchaseResearch.js`／FLOW／HUMAN_RUN | **C1+C2 GAS**: A-light（query+product cap・門/①候補/Keepaフルなし）→ セラーヒット全店再計算。`PR_ALIGHT_*` Property | `PR_ALIGHT_ENABLED=false`／Git revert。要 clasp |
| 2026-08-22 | `tools/competitor_store/dry_cycle_seller_accum.py` | **C0 fixture 4件**（基本/cap/429/除外）。GAS 前の正本 dry | ファイル削除 |
| 2026-08-22 | `PurchaseResearch.js`／FLOW §5.4／HUMAN_RUN | **C0 セラー貯め GAS**: 門通過→BB列+SP-API新品→`セラー`新規100/回。調査CK OFF。429/時間切れ再開 Property | Prop切戻し／Git revert。要 clasp |
| 2026-09-01 | `sub_image_food_canon_gate.py` | **見本ゲート**: 人間確定済みを再オープン時にレ点復元＋緑枠／前回バナー／手動画像を候補へ再同期 | Git revert |
| 2026-09-01 | `コード.js` | **B-④ ①′**: 見本画面表示と同時に `buildB5ProposalsForBindings` 裏実行（品番順・見本と同時進行） | Git revert。要 clasp |
| 2026-09-01 | `コード.js`／`b5_food_canon_status.py`／ALIGN §7.12.3・§7.15 | **B-④ ①′ 個品見本ゲート**: ダイアログ差し込み・Drive 状態 JSON・未確定ブロック。§7.15 ミックス5品番ドラフト | Git revert。要 clasp |
| 2026-08-22 | `tools/competitor_store/dry_cycle_alight_hit.py` | C1+C2 仮想 fixture 7件（A-light+ヒット再計算） | ファイル削除 |
| 2026-08-22 | `docs/org/B_PURCHASE_RESEARCH_CYCLE_DEV_MEMO.md` | C0–C5 仮想テスト方針 | docs revert |
| 2026-08-22 | `docs/org/B_PURCHASE_RESEARCH_CYCLE_*`／FLOW | **3レーン運用サイクル**要件＋3者多数決（条件付きYES）。A-light・セラーヒット再計算・完走連鎖・⑧週待ち廃止 | docs revert。実装 C0–C5 は別承認 |
| 2026-08-22 | `PurchaseResearch.js` | **C0 `prAppendSellerAccumRows_` setValues 修正**（37行/826行不一致。getRange 第3引数=行数） | Git revert。要 clasp |
| 2026-08-22 | `PurchaseResearch.js`／V1／HUMAN_RUN | **門**: ①プロファイルに順位下限（初期1・欠測落ち）＋価格上限（初期10万・カテゴリ別可）。列自動確保 | Git revert。要 clasp |
| 2026-08-22 | `コード.js`／STAGE | **Yahoo/Ichibaクエリ順**: JAN常に先頭（ミックスも）。FBはヒット率順（メーカー名→名→KW→N種）。要 clasp | Git revert |
| 2026-08-22 | `yahoo_category_virt/hit_rate_*` | **Yahooクエリヒット率調査**: パターン表P01–P12＋API bakeoff（読取のみ）。sample=行20/47/110相当 | ツール削除で戻す。clasp不要 |
| 2026-08-22 | `PurchaseResearch.js`／HUMAN_RUN／schema | **Keepaフル索引B1**: 書込時ASIN列だけ見て`Keepaフル索引`自動再構築。D1は必要ASINだけ読込。Prop `PR_KEEPA_FULL_INDEX_ENABLED`。99に再構築メニュー | Prop=false／Git revert。要 clasp |
| 2026-08-22 | `PurchaseResearch.js`／HUMAN_RUN | **Keepa単価**: プロファイル`product_offers20`既定6（旧2）。再開待ちは実測consumed/件数優先。残tokで次スライス継続（sliceWait≥5即終了廃止） | Prop上書き／Git revert。要 clasp |
| 2026-08-22 | `PurchaseResearch.js` | **店ゲート Keepa消費を①ログへ**: tokensConsumed列＋detailに qTok/pTok/consumed（offers=20の店調査が洗いと別枠で見える） | Git revert。要 clasp |
| 2026-08-22 | `PurchaseResearch.js` | **洗いトリガー予約修正**: arm時はProp先ON→トークン見積で予約。①ログに`トリガー予約`/`trigger_armed afterMin=`。A→D未武装時も再予約 | Git revert。要 clasp |
| 2026-08-22 | `PurchaseResearch.js`／schema／FULL | **実態ck**: 実態=要確認のみ Amazon検索URL（列名`実態ck`・末尾追加。M列実態は不変）。メニュー99「実態ck付与」 | Git revert。要 clasp |
| 2026-08-22 | FLOW §7.1 | **全店自動の最終形=案B**: 日次自動選定＋調査CK自動オン＋消化無人。日次件数上限で抑止。実装はモリタ後 | docs revert |
| 2026-08-22 | `コード.js`／STAGE | **Yahoo検索**: 打ち切り廃止・最大6qマージ・上位20・sort=-score（おすすめ順）。ログ BRANCH2_MERGE | Git revert。要 clasp |
| 2026-08-22 | `コード.js`／`yahoo_category_virt`／STAGE | **Yahooカテゴリ改修**: 点数制・ゆるいシリーズ第1FB・表記ゆれ・Driveガード。仮想9PASS（BASELINE比） | Git revert。要 clasp |
| 2026-08-22 | `コード.js`／`tools/yahoo_category_virt`／STAGE | **Yahooカテゴリ**: 同シリーズ層（+1）＋同商品（+2）。メーカー門。仮想6PASS | Git revert。要 clasp |
| 2026-08-22 | `コード.js`／`tools/yahoo_category_virt`／STAGE | **Yahooカテゴリ**: 同商品フィルタ→投票→票内重み。食品3a廃止。仮想4PASS（旧ロジックは41392等にFAIL） | Git revert。要 clasp |
| 2026-08-22 | `コード.js`／マスタGB-GD数式／STAGE | **診断ログ**: SKU監査／ミックスCOMPARE+Gemini／Yahoo BRANCH1-5+食品path優先／DUP_URL。GB「・」偽重複を数式修正 | Git revert＋数式戻し。要 clasp |
| 2026-08-22 | `docs/org/B_RERUN_CLEAR_AND_JUDGE_HUMAN_RUN.md` | **B再実行**: 消去列チェックリスト＋再B合否（G1–G7・ログ早見） | docs revert |
| 2026-08-23 | `CURRENT_PHASE` §0ab／`B_RERUN_CLEAR_AND_JUDGE` §4.3 | **G7保留**: ミックス楽天7207・同一URL複数親＝楽天に正セット無し時の汎用ヒット。非クリティカル・コード改修は問題拡大時まで | docsのみ |
| 2026-08-22 | `コード.js`／CURRENT／YAHOO_STAGE | **A→B検収FIX**: Yahoo選定ログ＋SHP数字query回避／一人分先頭数値／Step1色＝子SKU判定／ミックスWRITEログ | Git revert。要 clasp（出品） |
| 2026-08-22 | `PurchaseResearch.js` | **実態シード**（maker_alias同期）＋サマリ空欄へ実態付与。未登録=要確認。A→Dは製造/ブランド優先・Generic除外 | Git revert。要 clasp |
| 2026-08-22 | `PurchaseResearch.js`／FLOW／HUMAN_RUN | **本線メニュー**: 1.店1件通し／0.試験A→D(メーカー最大1)／5.D1→D2。旧は99サブ。全店自動はFLOW§7.1メモのみ | Git revert。要 clasp |
| 2026-08-22 | `PurchaseResearch.js`／FULL／HUMAN_RUN | **洗いGET offers=20**＋**D1案A**（フル行あればenrichしない）。単価既定2。薄い行は指紋同でもoffers上書き | Prop `KEEPA_PRODUCT_TOKEN_COST`／Git revert。要 clasp |
| 2026-08-22 | `PurchaseResearch.js`／FLOW／HUMAN_RUN | **洗い完走で調査CK解除**（トリガー完走含む）。**D1→D2 統合メニュー**（単独D1/D2は残置） | Git revert。要 clasp |
| 2026-08-21 | `PurchaseResearch.js`／FULL／HUMAN_RUN | **洗いキュー期限切れ削除**: Keepa済＋列挙日時>既定90日。メニュー＋完走時。Prop `PR_WASH_QUEUE_RETENTION_DAYS`（0=オフ） | Prop=0／Git revert。要 clasp |
| 2026-08-21 | `コード.js`／COMBO／SETCOUNT | **ミックスセット数**: A列＝何セットか（各N→N）。計N袋不使用。Vision＋参考画像ゲート・上限20。単品パース分離 | Git revert／Limit Property。要 clasp |
| 2026-08-21 | `PurchaseResearch.js`／プロファイル／FULL | 洗い完走ロック・トリガー予約のみ・Keepa upsert・門出品者数1–20・D1 offers補強 | トリガー停止／Prop。要 clasp |
| 2026-08-21 | `PurchaseResearch.js` | `PR_MAKER_WASH_FULL_PROP_` 欠落でトリガー ReferenceError → 定数追加 | Git revert。要 clasp |
| 2026-08-21 | `PurchaseResearch.js`／FULL／HUMAN_RUN | 洗い完走トリガー待ち: 残り×単価（上限300）÷rate＋1。旧30分キャップ撤廃 | Prop / Git revert。要 clasp |
| 2026-08-21 | `purchase_research/appsscript.json` | P3用に `script.scriptapp` スコープ追加（`getProjectTriggers` 権限エラー対策） | スコープ行を削除して clasp push。要**再承認** |
| 2026-08-21 | `PurchaseResearch.js`／FULL | **P3 洗い完走トリガー**（開始／停止／完走解除）。Keepa 429→RATE_LIMIT 温和停止 | Prop `PR_WASH_FULL_TRIGGER_ENABLED=false`＋トリガー停止メニュー。要 clasp |
| 2026-08-21 | `PurchaseResearch.js`／FULL | **洗い完走ループ**メニュー。1実行内スライス反復（トリガーなし・25分上限） | Git revert。要 clasp |
| 2026-08-21 | `PurchaseResearch.js`／FULL | **続きがKeepaに進むよう修正**: needEnum見直し・分割済の再推定禁止・キュー充分時は分割を済扱い | Git revert。要 clasp |
| 2026-08-21 | `PurchaseResearch.js`／FULL | **洗いキュー**: ASIN重複削除＋Keepa未済ユニーク＋同ASIN一括済。行増はユニーク増加（重複ではない） | Git revert。要 clasp |
| 2026-08-21 | `PurchaseResearch.js`／FULL | **分割語 nextToken 再開**。続きはサマリ→分割語→キューの順で対象選定 | Git revert。要 clasp |
| 2026-08-21 | `PurchaseResearch.js` | **P1bバグ**: `prWriteSplitWords_` の getRange 第3を終端行扱い→「15行 vs 16行」で FAILED。numRows に修正 | Git revert。要 clasp |
| 2026-08-21 | `PurchaseResearch.js`／dry_p1b／FULL | **洗い完走 P1b**: 分割B（タイトル→Gemini→OpenAI）・`①洗い分割語`・完走消込。尾西ドライ PASS | `PR_MAKER_WASH_FULL_ENABLED=false`。要 clasp＋①に GEMINI_API_KEY |
| 2026-08-21 | `コード.js` | **画像袋数 JSON修復**: フェンス除去・途中切れフィールド救出・MIME json・maxTokens512 | Git revert。要 clasp |
| 2026-08-21 | `コード.js` | **画像袋数ログ**: 不採用時も conf/set/texts/reason/img を Logger 出力 | Git revert。要 clasp |
| 2026-08-21 | `コード.js`／COMBO／SETCOUNT | **画像袋数改**: 抜き出しJSON・単位一般・highのみ採用。単品12-⑭もタイトル不能＋実行上限5 | Git revert／Property=false。要 clasp |
| 2026-08-21 | `コード.js`／COMBO | **ミックス袋数 Vision**: タイトル不能時のみ画像AI（親上限5・Yahoo優先）。`MIX_COMPETITOR_SET_COUNT_FROM_IMAGE` 未設定=ON | Git revert／Property=false。要 clasp |
| 2026-08-21 | `コード.js`／COMBO | **ミックス競合修正**: 除外後0件は次クエリ。N種/アソート標題のみ採用。子の楽天Yahooを空にしてから書込（単品N袋誤載せ・足し算残骸対策） | Git revert。要 clasp |
| 2026-08-21 | docs FULL／V1／FLOW | **洗い分割B**: タイトル→カテゴリ推定。Gemini主・OpenAI失敗時のみ。分割語シート＋消込 | docs revert |
| 2026-08-20 | `PurchaseResearch.js`／dry_p1／FULL | **洗い完走 P1+P2コード**。P1ドライ: 石原PASS／尾西 need_q2 でP2ドライ未着手 | `PR_MAKER_WASH_FULL_ENABLED=false`。要 clasp |
| 2026-08-20 | docs（FULL／FLOW／V1） | **メーカー洗い完走（案A+C）** 要件ロック。`①洗いキュー`・第2語分割・完走時のみ最終調査日。実装は別承認 | docs revert |
| 2026-08-20 | `コード.js`／COMBO／STORE | **ミックス楽天Yahoo**: N種モール検索→袋数一致の価格＋URL。0件は空（足し算なし）。R12足し算はAmazonのみ。選択・よりどり除外強化 | Git revert。要 clasp |
| 2026-08-20 | Keepaフル／ヒット／snapshot／①／Paste／Python | **先頭挿入**（最新=上）。セラー末尾のまま。メーカーサマリ新規のみ先頭。`dry_maker_sum --apply` 全消しは非本線 | Git revert。要 clasp（出品＋①） |
| 2026-08-20 | docs／`コード.js`／`PurchaseResearch.js` | **Keepa**: 応答優先。実測 refillRate=5 をフォールバック。洗い cap20。待ち=ceil(消費/rate) | docs／Property／Git revert。要 clasp |
| 2026-08-20 | docs／`コード.js`／`PurchaseResearch.js` | **Keepa Starter=20/分表記**（後に実測5へ訂正） | docs／Property／Git revert。要 clasp（出品＋①） |
| 2026-08-20 | docs／`コード.js` コメント | **Keepa トークン表現**: 日次枠ではない。毎分 `refillRate`・未使用60分失効（旧5/分表記はStarter20へ更新） | docs／Git revert |
| 2026-08-20 | `コード.js` | **競合 JAN≤5**: 1ページ目生ヒット5件以下ならメーカー+商品名をマージ（13-②/Yahooセット別/B Step2）。14-②は除外 | Git revert。要 clasp |
| 2026-08-20 | `コード.js` | **Yahoo 競合**: 楽天同順フォールバック（jan_code→query=JAN→メーカー+商品名→商品名）。14-②/セット別/B Step2/定時。Amazon◎は不変更 | Git revert。要 clasp |
| 2026-08-20 | `コード.js`／`bakeoff_rakuten_yahoo.py` | **楽天 Ichiba**: API版 `20220601`→`20260701`（旧版は Configuration not found）。Property `RAKUTEN_ICHIBA_ITEM_SEARCH_API_VERSION` で上書き可 | Property=`20220601`／Git revert。要 clasp |
| 2026-08-20 | `コード.js` | **楽天ジャンル（メニュー8）**: 1ロジック×多段フォールバック（セット親はN種＋JAN最後／単品はJAN優先）。投票フィルタ汎用化。Yahoo 7.6 共用 | Git revert。要 clasp |
| 2026-08-20 | `purchase_research/PurchaseResearch.js` | **未調査メーカー自動**: レ点不要・90日/空を1件選定→洗い→D1（`auto_maker_catalog`） | `PR_AUTO_MAKER_ENABLED=false`。要 clasp |
| 2026-08-20 | `purchase_research/PurchaseResearch.js` | **D1 Q/P列**: 出品 ASIN貼り付けのセット数量 regex を移植＋alt（種×袋・5食）。P=販売形態ラベル | Git revert。要 clasp |
| 2026-08-20 | `purchase_research/PurchaseResearch.js` | **D1 行2式**: skipMap 列の式が prListWriteRecs_ で空消しされていたバグ修正 | Git revert。要 clasp |
| 2026-08-20 | `purchase_research/PurchaseResearch.js` | **D1 品番リスト**: 画像URL=Keepaサブ画像からHYPERLINK、税率=0.08、挿入時に2行目の式コピー | Git revert。要 clasp |
| 2026-08-20 | `purchase_research/PurchaseResearch.js` | **D1 FBA列**: L1 first-fit（設定マスタ）＋Keepa FBA手数料フォールバック。空出品FBA手数料バグ修正 | Git revert。要 clasp |
| 2026-08-19 | `purchase_research/PurchaseResearch.js` | ブランド型メーカー洗い: 入口語の `page0=` 除去・Catalog entry 正規化・`PR_BRAND_NEEDLES_` 明示 needle | Git revert。要 clasp |
| 2026-08-19 | `tools/competitor_store/maker_alias.py` | 入口語ヒットを `brand/親 page0=…` 形式に（GAS split 整合） | Git revert |
| 2026-08-19 | `tools/common_ss_view/` | 共通スプシ参照の registry + dump。競合ツールの ID は registry 経由 | Git revert `tools/common_ss_view/` |
| 2026-08-19 | `コード.js`／COMBO／B_SPAPI | **ミックス親 FBA**: Catalog梱包空ならセット構成3辺で P1b 書込 | `B_FBA_P1B_COMBO_PACK_FALLBACK_ENABLED=false`／Git revert。要 clasp |
| 2026-08-19 | `コード.js`／COMBO／B_SPAPI | **ミックス親 FBAティア**: 1袋子ASINでCatalog。親競合ASIN非書。CQ自己発は子のまま | `B_FBA_COMBO_CHILD_ASIN_ENABLED=false`／Git revert。要 clasp |
| 2026-08-19 | メニュー成功は toast（alert は失敗のみ）。OK 待ちで別SS未反映を解消 | Git revert。要 clasp |
| 2026-08-19 | セラー再調査は機械欄上書き・同一IDは1行。切戻し `PR_SELLER_REWALK_OVERWRITE=false` | Git revert。要 clasp |
| 2026-08-19 | 同一sellerIdのA①空行へ店名行の構成比をコピー（採取元は残す） | Git revert。要 clasp |
| 2026-08-19 | レ点なしアラートに読込行・TRUE件数。同じsellerIdは店名あり行から補完 | Git revert。要 clasp |
| 2026-08-19 | 店GETはレ点の店のみ。onOpenはメニューだけ。CK枠の毎回再挿入をやめる | Git revert。要 clasp |
| 2026-08-18 | ① GAS メニューは25分で leftover（課金30分）。onOpen は30秒 | Git revert。要 clasp |
| 2026-08-18 | 店名欠けるIDは /seller storefront=0 で取り台帳へ | `PR_SELLER_NAME_GET_ENABLED=false`。要 clasp |
| 2026-08-18 | ①候補に店名列（sellerIdは参考・`|` 連結） | Git revert。要 clasp |
| 2026-08-18 | ①候補 1ASIN1行・sellerId は `|` 連結 | Git revert。要 clasp |
| 2026-08-18 | 調査CK にチェック枠＋レ点行を青字 | 条件付き書式削除／validation解除 |
| 2026-08-18 | ① GAS レ点の店／メーカー（調査CK・1000超は次へ・再巡回仮6か月） | `PR_CK_SELLER_ENABLED=false`／`PR_CK_MAKER_ENABLED=false`。要 clasp |
| 2026-08-18 | D1 新規を行5挿入（既存下へ） | Git revert。要 clasp |
| 2026-08-18 | E 実装保留メモ FLOW §5.3.1 | docs revert |
| 2026-08-18 | ① GAS D2 機械絞込 `prMenuListFilterD2` | `PR_LIST_FILTER_ENABLED=false`／列クリア。要 clasp |
| 2026-08-18 | ① GAS D1 品番リスト追記（実機PASS） | `PR_LIST_APPEND_ENABLED=false`／追記行削除。要 clasp |
| 2026-08-17 | ① GAS Keepaフル追記（GET分） | `PR_KEEPA_FULL_WRITE_ENABLED=false`／末尾行削除。要 clasp |
| 2026-08-17 | ① GAS メーカー洗い cap5（Catalog→門） | `PR_MAKER_WASH_ENABLED=false`／Git。要 clasp |
| 2026-08-17 | 接続スイート T2–T4（AGF/リフココ/BBQ） | ①候補末尾5（T2）／実行ログ | Git／行削除 |
| 2026-08-17 | 接続T1 尾西食品 cap5 | ①候補末尾5／フル+4 | Git／行削除 |

| 2026-08-16 | メーカーサマリ2層（A/B） | タブ削除／clear | Git revert |

| 2026-08-18 | R12 ミックス寸法: 長辺・短辺は最大、最短のみ数量倍 | Git revert。要 clasp |
| 2026-08-18 | R12 ミックスKeepa◎袋はAmazonその値段／物流は親1袋3辺 | Git revert。要 clasp |
| 2026-08-18 | R12 ミックスは12-⑭非書／1袋欠けは別モール1袋借用 | Git revert。要 clasp |
| 2026-08-18 | R12 競合子行セット1／Bキュー分離／構成ゲート | Git revert。要 clasp |
| 2026-08-18 | R12 大きさは単品親の競合Keepa。同シリーズは1袋のみ流用 | Git revert。要 clasp |
| 2026-08-17 | R12 構成照合は商品名（セット税込＝和・照合キーにしない） | Git revert。要 clasp |
| 2026-08-16 | R12 組み合わせ（色・Keepa＝AI行・構成記入・足し算） | Git revert。要 clasp |
| 2026-08-16 | Peacefulショップ 一連・リスト+8 | 8行目クリア／リスト末尾8 | セル・行戻し |
| 2026-08-16 | 門に使う順位段階を文字列プルダウン | 数字1に戻す／検証解除 | Git＋シート。要 clasp |
| 2026-08-16 | ストアASIN数は応答がある店だけ | docs戻し | Git revert |
| 2026-08-16 | Camel Trading 台帳列＋リスト+8 | 7行目クリア／リスト末尾8 | セル・行戻し |
| 2026-08-16 | 台帳→構成比 Camel Trading 既存行 | 店名・食品比率等クリア | セル戻し |
| 2026-08-16 | リフココ Keepa+20・リスト+14。いしとき query1 | リスト末尾14／台帳 | 行削除 |
| 2026-08-16 | 輪1周 A18COIYGTMT3AC Keepa20・リスト+11 | リスト末尾11行／台帳巡回日 | 行削除／巡回日クリア |
| 2026-08-16 | A① 2秒再開+7・まとめ+34・429 | セラー末尾41行付近 | Git revert |
| 2026-08-16 | リサーチ完成形フロー A–J / A′ を正本化 | FLOW.md 削除／DOMAIN1・V1 のポインタ戻し | docs Git revert |
| 2026-08-16 | L1–L6 一括（構成%削除・miss Keepa・JOIN・次セラー・みそ汁20） | 各塊の戻しは NOW | Git revert |

| 2026-08-16 | 処理時間 ▼実行時間（B Step＋Step5 APIのみ）。全関数ばら撒き禁止 | `OP_TIMING_ENABLED=false`。ルール `op-timing-scope.mdc`。要 clasp |
| 2026-08-16 | Step5 OpenAIはGemini失敗時のみ。AI梱包停止 | `B_STEP5_SKIP_OPENAI_IF_GEMINI_OK=false`／`B_STEP5_SKIP_PACK_AI=false`。要 clasp |
| 2026-08-16 | 行高さ B Step8 廃止（CLIP/setRowHeight削除） | Git revert。要 clasp |
| 2026-08-16 | A.準備 商品名SEOスペース（AIシートのみ） | メニュー3/4削除／Git revert。要 clasp |
| 2026-08-16 | B Step2.1＝12-⑭。ストアON時は横断シート直書きスキップ | `B_COMPETITOR_STORE_APPLY_ENABLED=false`／Git revert。要 clasp |
| 2026-08-16 | S4b セラー メイン＋構成％列（カテゴリ名） | 列削除 | Git revert |
| 2026-08-16 | S4 セラー構成比 storefront=0 | 4列クリア | Git revert |
| 2026-08-16 | S3 モリタ /query page0。台帳件数30 | 巡回日・件数を戻す | Git revert |
| 2026-08-16 | T2 品番リスト JOIN82 画像HYPERLINK | 当該セルをURLに戻す | Git revert |
| 2026-08-16 | A画像Gemini OFF検証は後続。今は実装なし | Property付けない | docs revert |
| 2026-08-16 | K8 BB/レビュー/BuyBox_FBA は今JSON空。代替はメモのみ | 実装なし | docs revert |
| 2026-08-16 | F2b 画像HYPERLINK／サブ画像\|（画像一覧リネーム） | 列クリア or ヘッダ戻し | Git revert |
| 2026-08-15 | F2 画像・画像一覧（生JSON images[]） | 列クリア | Git revert |
| 2026-08-15 | F1 Keepaフル カテゴリ・梱包・FBA手数料等を生JSON展開 | 当該列クリア | Git revert |
| 2026-08-15 | T1 品番リスト通過82追記（複製・式非上書き） | 末尾82行削除 | Git revert |
| 2026-08-15 | L1 出品FBA／自己発送 first-fit をKeepaフルへ | 5列削除 | Git revert |
| 2026-08-15 | F0 flatten 直販・新品現在・現行順位・oos180 | 列削除 | Git revert |
| 2026-08-15 | 品番リスト課題K1–K7。flattenはGETなし。L1は出品流用 | 実装なし | docs revert |
| 2026-08-15 | A.準備メニュー（Catalog空欄。12-⑮⑯同一） | メニューから項目削除 | Git revert。要 clasp |
| 2026-08-15 | P5 A末尾でP2並べ替え | `AMAZON_PASTE_P5_RANK_AFTER_A_ENABLED=false` | Git revert。要 clasp |
| 2026-08-15 | 出品者数＝COUNT_NEW。Keepaフル列化159。61空 | 列削除 | Git revert |
| 2026-08-15 | P1計画 dry（空ASIN0・Catalog GETなし） | 貼付非書 | スクリプト削除／docs revert |
| 2026-08-15 | P5計画 dry（P2並べ替え読取、A未結線） | 貼付非書。◎維持 | スクリプト削除／docs revert |
| 2026-08-15 | WRITE計画 dry（貼付 vs Keepaフル、非書） | Property OFF。raw-only フック | スクリプト削除／docs revert |
| 2026-08-15 | 競合DB「セラー」モリタ1行。フルから貯めない | タブ削除／行削除 | Git revert |
| 2026-08-15 | W4 hydrateは門stats空ならGET（T28） | 日付だけ新鮮な空JSONを載せない | Git revert。要 clasp |
| 2026-08-15 | ①タスク構造化メモをツリー正（NOW同文）に修正し他PJ共有 | 洗いと貼付を混ぜない。W4は出品 | docs Git revert |
| 2026-08-15 | ①タスク構造化メモを他PJ共有 | 洗いと貼付を混ぜない。W4は出品 | docs Git revert |
| 2026-08-15 | Keepa倉庫=Keepaフル。①候補は転記。W1が次 | Catalogはモールヒットに書かない | docs Git revert |
| 2026-08-15 | 領域1①進捗を NOW／DOMAIN1 §1.1／PHASE §0re に共有 | 他PJが洗いと貼付を混ぜない | docs Git revert |
| 2026-08-15 | P3 2種以上はマスタ非載 | ◎でもミックスを単品JANに載せない | Git revert。要 clasp |
| 2026-08-15 | 袋数 N種×各M／P2機械◎でも非候補／P3タイトル優先 | 人が直す前の機械精度。Keepaセルよりタイトル。A非改変 | Git revert。要 clasp（コード.js＋AmazonCompetitorPaste.js） |
| 2026-08-15 | P2 12-⑰⑱／paste_rank.py | 貼付非候補化。A非改変。評価◎維持 | Git revert。要 clasp。先に⑰ |
| 2026-08-15 | `purchase_research/` clone | ① GAS を Cursor で見えるようにした。出品 clasp 非変更 | フォルダ削除／出品 clasp は触らない |
| 2026-08-15 | `parseSetCountFromItemNameWithSource`／12-⑭外れ単価／ACCURACY.md | 計N袋優先。各N袋誤クラスタ防止。単価2倍は12-⑭で除外。課題を文書化 | Git revert。要 clasp push |
| 2026-08-15 | `コード.js` 12-⑭／`apply_to_master.py` | 専用ヒットを JAN＋セット数でマスタ競合列へ | Prop `COMPETITOR_MASTER_APPLY_ENABLED=false`／Git revert。要 clasp push |
| 2026-08-15 | `コード.js` flush 指紋スキップ | 変化なしモールヒット非書込 | Git revert。要 clasp push |
| 2026-08-15 | FIELDS／STORE／NOW | 楽天ポイント＝％。ヒット列は意図的に絞る。領域1実装設計草案 | docs Git revert |
| 2026-08-15 | `docs/org/COMPETITOR_STORE_DEV_MAP.md` | 開発構造の一時メモ。完了後削除 | ファイル削除 |
| 2026-08-15 | `docs/org/COMPETITOR_STORE_REQUIREMENTS.md`／`tools/competitor_store/`／`コード.js` dual-write | 競合専用スプシ段階1。マスタKeepaは残す。ENABLED未設定=OFF | Property OFF／Git revert。要 clasp push（GAS分） |
| 2026-08-15 | `docs/org/COMPETITOR_FIELDS.md`／`competitor_fields/*.csv`／`.cursor/rules/competitor-fields.mdc` | 競合項目の正本化とPJ横断の参照ルール。コードなし | docs Git revert |
| 2026-08-15 | `docs/org/B_AMAZON_COMPETITOR_PASTE_REQUIREMENTS.md` ほか | Amazon貼付自動化の制約（◎人間／候補列／キャッシュ＋領域1再利用）。コードなし | docs Git revert |
| 2026-08-14 | `コード.js` Keepaキャッシュ読取 | setCount空で旧形式誤判定しブランド等がログから消えるのを修正。ヘッダー名読取＋空上書き禁止 | Git revert。要 clasp push |
| 2026-08-14 | `コード.js` Keepa | 梱包欠落キャッシュはAPI再取得。`梱包_checked` で空返し再取得防止 | Git revert／`KEEPA_CACHE_REFETCH_WHEN_PKG_EMPTY=false`。要 clasp push |
| 2026-08-14 | `コード.js` Keepaキャッシュ | 梱包4列をキャッシュ保存＋ログ追記コピー。既存ログ非改変 | Git revert。要 clasp push |
| 2026-08-14 | `コード.js`／FOLLOWUP／PHASE | 親行CLIP後勝ち63px／選定不可はCT〜CV式維持／B 6.05改行／Yahoo query=JAN／楽天名フォールバック＋袋パース | Git revert。要 clasp push。OFF: `B_PARENT_ROW_HEIGHT_ENABLED` |
| 2026-08-15 | NOW ④の正 | `/query` 形をロック。sellerIds 可変 | docs Git revert |
| 2026-08-14 | `コード.js` `getKeepaCachedResults` | キャッシュのブランド／製造者をログ追記へコピー。既存ログ非改変 | Git revert。要 clasp push |
| 2026-08-14 | Keepa API 2品番 | CSVエクスポートと product API 比較。tokensLeft=278。キー非コミット | スクリプト削除／docs revert |
| 2026-08-14 | `docs/org/B_PURCHASE_RESEARCH_NOW.md`／`tools/purchase_research_path3/` | 領域1①を**経路3先行**に差し替え。差分PoC＋SP-APIキーワード読取骨格。GAS第1版なし | docs/tools Git revert |
| 2026-08-14 | `コード.js`／`B_RUN_20260814_FOLLOWUP`／PRICING_V1 | **B完走フォロー**: 寸法なし空欄・CPO非塗色・Step5当B clear・BにFBA P1b(6.55)・KW seen1本・親行CLIP60。⑥Nはコード非改修 | Git revert。要 clasp push。OFF: `B_FBA_P1B_WRITE_ENABLED`／`B_PARENT_ROW_HEIGHT_ENABLED` |
| 2026-08-14 | `docs/org/B_PURCHASE_RESEARCH_V1_REQUIREMENTS.md` | 領域1①詳細稿。Keepa件数キャップ削除・上限300トークン | docs Git revert |
| 2026-08-14 | `docs/DOMAIN1_RESEARCH_PURCHASING.md` ほか | 領域1正本（①第1・C下書き・調査URL） | docs Git revert |
| 2026-08-14 | `points_logic`／`points_send` | restore の `--today` をカレンダー%に渡す | Git revert |
| 2026-08-14 | `コード.js` | **AI列名引き**: ASIN貼り付け式・参考画像・卸値。`原価税込`別名 | Git revert。要 clasp push |
| 2026-08-14 | DEALS HUMAN_RUN／PHASE | **taper 1段目 prod**: b=`185913020679`／s=`185917020679` DONE。運用マニュアル化を後続タスクにメモ | docs Git revert |
| 2026-08-14 | `コード.js`／`B_HARD_DEATH_SCOPE_*`／PHASE／WATCHDOG | **Bハード死対策**: 入口3・Step1切断・insertedProductsマージ・集合フィルタ・getUi安全・番犬非設置 | Prop トグルOFF／Git revert。要 clasp push。`B_WATCHDOG_ENABLED=false` は残してよい |
| 2026-08-14 | `コード.js` | Step1で **依頼日(D)** を1行目→親／2行目→子へ式コピー（Qより左の個別対応） | 当該関数削除／revert |
| 2026-08-14 | `_local_backup/pre_B_STEP1_TOP_INSERT_20260814_001210/`（gitignore）／CHANGE_LEDGER | **復元点**: 上挿入実装前の `コード.js`／Yahoo.js／appsscript／clasp をローカル退避。HEAD=`2f9ef6c` | フォルダから Copy-Item（RESTORE.md） |
| 2026-08-14 | `docs/org/B_STEP1_TOP_INSERT_*`（多数決／要件／HUMAN_RUN）／PHASE／HANDOVER | **上挿入要件ロック**: 1A/2A/3A・既存値固定・テンプレ1–2。コード未 | docs Git revert |
| 2026-08-13 | `docs/org/B_STEP1_TOP_INSERT_REQUIREMENTS.md`／PHASE／HANDOVER | **上挿入改修メモ**初版（式方針は08-14で更新） | docs Git revert |
| 2026-08-13 | `コード.js` masterColorApplyLocked*／15-㉑㉒／色 docs §11–13／スプシ11–775 | **⑧f 列ロック確定**。親赤白薄赤。Step1自動塗 | Git revert。スプシ色は再塗or手戻し。要 clasp push |
| 2026-08-13 | `コード.js` MASTER_COLOR／CPO・Step1・カテゴリ等／色 docs | **⑧c 役割色実装**。オレンジ後回し。条件付き書式非変更 | Git revert。要 clasp push |
| 2026-08-13 | `docs/org/B_MASTER_CELL_COLOR_RULES_*`／PHASE／HANDOVER | **⑧a マスタ色ルール正本**（灰・オレンジ含む）。コード一斉塗なし | docs Git revert |
| 2026-08-13 | `コード.js` B Step8／WATCHDOG §7.1／PHASE／HANDOVER | **⑪親行高さ60px**。⑨改修不要・⑩運用検証メモ | Prop `B_PARENT_ROW_HEIGHT_ENABLED=false`／Git revert。要 clasp push |
| 2026-08-13 | `コード.js` Step1／17-⑦／MASTER_LINKAGE／RESEARCH／PHASE | **⑦ DF割引式**: 先頭子より下へ DF2 `ROUND` コピー。DDは対象外 | Prop `B_DF_STRATEGY_FORMULA_COPY_ENABLED=false`／Git revert。要 clasp push |
| 2026-08-13 | `コード.js` 15-⑳／承認包／B_SPAPI docs | **⑤競合空のみ自動**: U3高信頼→競合列。U7本線当面不要を明記 | Prop `B_COMP_ASIN_AUTOFILL_ENABLED=false`／Git revert。要 clasp push |
| 2026-08-13 | `コード.js` 15-⑲／B_SPAPI docs／⑤承認包 | **U7診断**: Catalog型番 vs Keepa。マスタ非書込。⑤競合列は承認包のみ | Prop `B_U7_PART_DIAG_ENABLED=false`／Git revert。要 clasp push |
| 2026-08-13 | CURRENT_PHASE／DEALS HUMAN_RUN／HANDOVER | **8/14前 taper dry_run**: 当日0・Cinderellas2SKU予定。FBA完了→減衰へフォーカス切替 | docs Git revert |
| 2026-08-13 | B_SPAPI要件／HUMAN_RUN／PHASE／HANDOVER | **FBAゲート暫定OK**: 必須=HTTP+梱包。型番/JAN→U7。U7未着手 | docs Git revert |
| 2026-08-13 | B_SPAPI HUMAN_RUN／PHASE／HANDOVER | **P1b実機** `P1b_20260813_162652_6111aa`: ティア/手数料 11/11書込。U7は型番JAN方針待ち | docs Git revert。マスタ値は手戻し |
| 2026-08-13 | マスタ出品CK／B_SPAPI HUMAN_RUN／PHASE | **P1a `595a71f2`集計**: HTTP20/20・梱包11/20。生姜湯2ASINレ点OFF→梱包あり6ASINのみON（P1b用） | レ点は手で戻す。docsは Git revert |
| 2026-08-13 | `コード.js` B統合 | **U6**: Step**6.6** メーカー品番（U5 quiet）。`B_INTEGRATED_STEP_FUNCTIONS` | Prop `B_MAKER_MODEL_FETCH_ENABLED=false`／Git revert。要 clasp push |
| 2026-08-13 | `コード.js` U5 | **U5自社品番**: `INT-`＋8桁hex（計12）。旧20字から短縮 | Git revert。要 clasp push |
| 2026-08-13 | `コード.js` U5 | **U5候補ガード**: JAN同一却下／純数字は4桁未満のみ却下（16100可） | Git revert。要 clasp push |
| 2026-08-13 | `コード.js` 15-⑤⑥ | **U5メーカー**: `メーカー名ベース`のみ（`メーカー名`不使用） | Git revert。要 clasp push |
| 2026-08-13 | `コード.js` 15-⑥ | **U5結果シート**: setValues行数不一致修正（getRange第3=行数） | Git revert。要 clasp push |
| 2026-08-13 | `コード.js` 15-⑤⑥ | **U5書込先**: **`メーカー品番下書き`のみ**（メーカー型番等NG） | Git revert。要 clasp push |
| 2026-08-13 | `コード.js` 15-⑤⑥ | **U5列名**: `メーカー品番`無し→`メーカー品番下書き`／`メーカー型番`フォールバック | 次行で下書きのみに訂正 |
| 2026-08-13 | `コード.js` 15-⑤⑥／B_SPAPI docs | **U5 メーカー品番**: Keepa→Serp→INT-自社（黄）。型番不明廃止。空のみ | Prop `B_MAKER_MODEL_FETCH_ENABLED=false`／Git revert。要 clasp push |
| 2026-08-13 | マスタ列／`コード.js` 15-⑱／B_SPAPI docs | **FBA P1b**: `FBAティア`／`FBA手数料_円` 空のみ書込（自己発列非改変） | Prop `B_FBA_P1B_WRITE_ENABLED=false`／Git revert。列は手削除可 |
| 2026-08-13 | `コード.js`／B_SPAPI docs | **⑤ U3.4f**: brandのみ照合・セット親×Catalog単品ヒント（0s1） | Git revert。要 clasp push |
| 2026-08-13 | `コード.js`／`AmazonCategoryPt.js`／B_SPAPI docs | **⑤ U3.4e**: brand/製造者のみ照合・タイトル除外・unknown非推奨 | Git revert。要 clasp push |
| 2026-08-13 | `コード.js`／B_SPAPI docs | **⑤ U3.4d**: U3メーカー正＝`メーカー名ベース`のみ | Git revert。要 clasp push |
| 2026-08-13 | `コード.js`／B_SPAPI docs | **⑤ U3.4c**: maker一般照合（カナ↔ローマ字＋Catalog brand。別名表なし） | Git revert。要 clasp push |
| 2026-08-13 | `コード.js`／B_SPAPI docs | **⑤ U3.4b**: brand yes Catalog優先（0s1）／conflictはmakerInTitle／兄弟は他親のみ | Git revert。要 clasp push |
| 2026-08-13 | `コード.js`／B_SPAPI docs | **⑤ U3.4a**: Catalogタイトル bi-gram≥0.80＋LCS。P4b関数は不変更 | Prop `B_KEEPA_ASIN_VOTE_U3_TITLE_SCORE`／Git revert。要 clasp push |
| 2026-08-13 | `コード.js`／B_SPAPI docs | **⑤ U3.3**: 同一JAN同セットの兄弟競合票＋Catalogタイトル不一致抑制 | Git revert。要 clasp push |
| 2026-08-13 | `コード.js`／B_SPAPI docs | **⑤ U3.2**: ブランド一致・Catalogヒント・競合404・set_mismatch を推奨に反映 | Git revert。要 clasp push |
| 2026-08-13 | `コード.js`／B_SPAPI docs | **⑤ U3.1b**: 親`A.セット商品数`空→同一親SKUの子から継承（1優先／min） | Git revert。要 clasp push |
| 2026-08-13 | `コード.js`／`AmazonCategoryPt.js`／B_SPAPI docs | **⑤ U3.1**: 親セット数ゲート・setGuess/ブランド一致列・Catalog票は単品親のみ | Prop `B_KEEPA_ASIN_VOTE_U3_ENABLED=false`／Git revert。要 clasp push |
| 2026-08-13 | `コード.js` 15-⑯／B_SPAPI docs | **⑤ U3**: 競合ASIN投票診断（◎／N列／競合店／Catalog JAN→`競合ASIN投票診断_U3`）。マスタ非書込 | Prop `B_KEEPA_ASIN_VOTE_U3_ENABLED=false`／Git revert。要 clasp push |
| 2026-08-13 | `コード.js` 12-⑧／B_SPAPI docs | **⑤ U2**: Keepa取得_ログ→archive（90日超 or JAN最新3実行以外）。マスタ非書込 | Prop `B_KEEPA_LOG_ARCHIVE_ENABLED=false`／archiveから手戻し／Git revert。要 clasp push |
| 2026-08-13 | `コード.js` Keepaログ／B_SPAPI docs | **⑤ U1**: Keepa取得_ログに partNumber／ブランド／製造者／梱包寸法・重量（追加APIなし） | Prop `B_KEEPA_LOG_EXT_ENABLED=false`／Git revert。要 clasp push |
| 2026-08-13 | `00_設定マスタ` FBA手数料／`update_fba_fee_table_00.py`／承認包 | **FBA手数料を公式現行（1,000円超）へ更新**。FBA末尾に行挿入。販売手数料等は非改変 | スプシ版履歴／承認包旧値 |
| 2026-08-13 | `コード.js` FBA P1a | **設定マスタFBA手数料連動**: 仮ティア＝区分名／手数料円。F備考をbox・三辺和でパース | 要 clasp push／Git revert |
| 2026-08-13 | `コード.js` `fbaP1aLengthToCm_` | **FBA P1a単位バグ修正**: `centimeters` が meters 誤判定→×100。cmを先判定・mは厳密一致 | 要 clasp push／Git revert |
| 2026-08-13 | `コード.js`／B_SPAPI docs | **FBA P1a**: Z 15-⑰ Catalog寸法→`FBAティア診断_P1a`・仮ティア。マスタサイズ列は書かない | Prop `B_FBA_P1A_DIAG_ENABLED=false`／Git revert。要 clasp push |
| 2026-08-13 | `コード.js` 3Dフィット | **rigid優先**（softは rigid 全滅後）。Compact stagger 当面維持 | 要 clasp push |
| 2026-08-13 | `コード.js` 3Dフィット | **ずらしバグ修正**（半辺で誤OK）。`[3D][try]` で内寸・soft・bestCap・ok をログ | 要 clasp push |
| 2026-08-13 | `コード.js` Step3.1 | **AI梱包寸法**: 商品名＋卸値(税込)優先／JAN+卸値／JANフォールバック。同JAN複数行で誤寸法を防止 | 要 clasp push |
| 2026-08-13 | `コード.js` Step3.1 | **サイズ昇順 first-fit**。送料=自己発送／箱代=資材D／合計=利益のみ。寸法あり不適合は空（ネコポス埋め禁止） | 要 clasp push。Git revert |
| 2026-08-13 | `コード.js` Step3.1 | **ネコポス利益上書き廃止**。3D時はサイズのみで箱選定。利益不足はフラグのみ | 要 clasp push。Git revert |
| 2026-08-13 | `コード.js` Step3.1／設定マスタ exclude／B_SPAPI docs | **④自己発3D**: 内寸フィット・exclude除外。`B_LOGISTICS_USE_3D_FIT` | Prop `false`／Git revert。要 clasp push |
| 2026-08-13 | `00_設定マスタ` E〜H／`write_box_inner_dims.py`／B_SPAPI_RESEARCH docs | **④⑤⑥要件＋ダンボール内寸列** | 列E〜H手削除／docs revert |
| 2026-08-12 | `コード.js` Step5/6／SERIES_INFER docs | **シリーズ不明禁止＋商品名推定（方針A）** | Prop `SERIES_INFER_FROM_NAME_ENABLED=false`／Git revert。要 clasp push |
| 2026-08-12 | AI_ADOPT 要件／HUMAN_RUN | **商品名ベース可変不可**を要件明文化（読取のみ） | docs revert |
| 2026-08-12 | `コード.js` メニュー8／Yahoo Stage／E・YAHOO docs | **7.5維持／7.6 popular_only＋38074。B Step7.6差し替え** | Git revert。要 clasp push |
| 2026-08-12 | `コード.js` Step5／メニュー8／E docs | **E本線化**: ジャンル／Yahoo未設定=ON＋Step5ジャンルAI・Drive候補スキップ | 各Prop `false`／`B_STEP5_SKIP_GENRE_YAHOO_AI_ENABLED=false`／Git revert。要 clasp push |
| 2026-08-12 | `コード.js` Step6.5／15-⑮／B_ASIN_N docs | **⑥ N列ASIN自動**: ◎×ブランド＝メーカー→空Nのみ＋黄セル要確認 | **Property** `B_ASIN_N_AUTO_FILL_ENABLED=false`／Git revert。要 clasp push |
| 2026-08-12 | `コード.js` メニュー8／AI_ADOPT docs | **v1.13**: FO二重計上解消・商品名非削り・表示文字数trim | Git revert。要 clasp push |
| 2026-08-12 | `コード.js` メニュー8／AI_ADOPT docs | **v1.12**: 短すぎ対策（FO非削り・下限ブレーキ・特徴用途を空にしない） | Git revert。要 clasp push |
| 2026-08-12 | `コード.js` メニュー8／AI_ADOPT docs／B Step7.5 | **v1.11**: =LEN()厳格trim・FO先頭・B統合載せ | **Property** `B_INTEGRATED_MENU8_ENABLED=false`／Git revert。要 clasp push |
| 2026-08-12 | `コード.js` 15-⑬⑭／要件 | **P2-A**: 月次メーカー辞書（手修正維持）＋保存方法改行→読点 | Git revert。要 clasp push |
| 2026-08-12 | `コード.js`／設定マスタA181／AI G列／要件 | **B P1**: シリーズ揃え・メーカー辞書・カタログスラッグ13 | Property各`false`／Git revert。要 clasp push |
| 2026-08-12 | `コード.js`／`B_WATCHDOG_*`／要件 | **B番犬P0**: ハード死後自動再開・Step5行CP・プルダウンTIME_SLICE・サマリ・進捗メニュー | **Property** `B_WATCHDOG_ENABLED=false` または **Git** revert。要 clasp push |
| 2026-08-12 | DEALS／HUMAN_RUN／HANDOVER | **表記揃え**: P1c実装済。G8＝終着空でもapply可。復元済＝減衰中%へ | **Git**: revert 当該docs |
| 2026-08-12 | `price_recovery_logic`／`taper_send`／`points_*`／GAS／DEALS§10.13-14 | **2層**: カレンダー減衰中%＋B中1%オーバーレイ。restore=減衰中%。fetchは現在%のみ。列`減衰開始日` | **Git**: revert。`--schema-only`＋clasp push |
| 2026-08-12 | `sheets_io` 次回減衰後%数式 | **f-string 二重引用で `{c_pct}` が残るバグ修正** | **Git**: revert。`--schema-only` 再実行 |
| 2026-08-12 | 広告運用GAS `onOpen`／HUMAN_RUN／DEALS | **メニュー分割**: タイムセール本線＋99。広告運用は分析のみ | **Git**: revert。要 clasp push |
| 2026-08-12 | `taper_send.py`／sheet_schema／GAS／DEALS§10.14 | **P1c**: taper＋メール＋実行依頼E。リカバリ手動 | **Git**: revert。要 clasp push／`--schema-only`。日次タスクは人手 |
| 2026-08-12 | DEALS§10.14／HUMAN_RUN／PHASE／HANDOVER | **P1c要件**: 現スナップ＋進捗。E先行。C／D／URL後続 | **Git**: 当該docs revert |
| 2026-08-12 | `sheet_schema` 減衰*改名／販促売価削除／GAS／DEALS | **戻し→減衰列名。減衰段%明示。販促売価円削除** | **Git**: revert。`--schema-only`＋clasp push |
| 2026-08-12 | `sheets_io` display formulas／GAS／DEALS§2.2 | **販促ポイント円・実質価格円をシート数式化**（販促売価円は非推奨で空） | **Git**: revert。`--schema-only` 再実行可 |
| 2026-08-12 | `sheet_schema`／sync realign／GAS Realign／DEALS§2.2 | **マスタ列：目的グループ＋人入力先頭＋黄セル** | **Git**: revert。要 clasp push／`--schema-only` またはメニュー並べ替え |
| 2026-08-12 | `price_recovery_*`／`sheet_schema`／GAS PriceRecovery／DEALS§10.12 | **実質戻し**: 販促ポイント%＋円表示。売価段上げ廃止 | **Git**: revert。要 clasp push（広告）・`sync --schema-only` |
| 2026-08-12 | `docs/org/MEMO_POINTS_VS_REFERENCE_PRICE.md` | **ポイント×過去価格メモ**: 記事要約・公式原文未特定・30日/90日/8週整理 | **ファイル削除**または revert |
| 2026-08-11 | `price_recovery_*`／GAS QtyMail／DEALS§10.12 | **価格戻しAPI＋G10**: our_price段階上げ・B/restoreガード | **Git**: revert。要 clasp push（広告） |
| 2026-08-11 | `TimeSaleQtyMail.js`／`コード.js`／DEALS docs | **Points G5**: apply／restore Cursor指示メニュー | **Git**: revert。要 clasp push（広告） |
| 2026-08-11 | `mail_points_remind`／HUMAN_RUN | **Points運用予行**: 日程基準リマインド＋Smile下書き確認 | **Git**: revert |
| 2026-08-11 | `points_fetch`／`test_points_fetch`／要件§10.10 | **Points G9**: listings pointsNumber÷price で%本線 | **Git**: revert |
| 2026-08-11 | `points_send`／`points_fetch`／要件§10.10 | **Points G8**: セール前%空・backup失敗でapply中止 | **Git**: revert。`--allow-missing-before` |
| 2026-08-11 | `points_logic`／`points_send`／要件§10.10.1c | **Points G4**: 施策連動SKUフィルタ（既定ON） | **Git**: revert。`--all-master`で旧挙動 |
| 2026-08-11 | `mail_points_remind`／GAS QtyMail／要件§10.10.1b | **Points G2/G3**: apply/restore リマインド | **Git**: revert。要 clasp push（広告） |
| 2026-08-11 | `points_logic`／send／fetch／要件§10.10 | **Points G7**: 状態語彙固定＋sheet更新揃え | **Git**: revert |
| 2026-08-11 | `points_send`／HUMAN_RUN | **Points G1**: 1SKU prod 2%→1%（feed DONE）。読取GETは未 | **Git**: revert。%はマスタで戻し可 |
| 2026-08-11 | `build_submit_xlsx`／HUMAN_RUN | **成功submitv2準拠**: compact+日付直書き+保護OFF（全行残しはSC処理失敗） | **Git**: revert |
| 2026-08-11 | `build_submit_xlsx`／HUMAN_RUN | **開始/終了はVLOOKUP残す**（参加中＋スケジュールのみ書込。日付直書きはSC失敗） | **Git**: revert |
| 2026-08-11 | `build_submit_xlsx`／HUMAN_RUN | **提出xlsx日付=YYYY-MM-DD文字**（データ定義準拠・シリアル混入修正） | **Git**: revert |
| 2026-08-11 | sync／template_parse／notify_mail／要件§10.9 | **②ドロップダウン日付付き全登録**＋公式衝突時独自短縮メール | **Git**: revert |
| 2026-08-11 | `mail_qty_confirm`／GAS `TimeSaleQtyMail`／HUMAN_RUN§9.7 | **数量確認メール実送信**（Gmail API＋下書きシート＋メニュー）→ contact@octas2301.com | **Git**: revert／token_gmail_send は secrets 外 |
| 2026-08-11 | HUMAN_RUN／`mail_qty_confirm`／DEALS§9.7 | **A未運用表記揃え**／数量確認はUL済対象＋`--tol` | **Git**: revert |
| 2026-08-11 | `parse_ymd`／`invariants.py`／sync／DEALS§0.1 | **並び修復＋不変条件自動検証**（抜け落ち再発防止） | **Git**: revert |
| 2026-08-11 | sync／build／sheets_io／DEALS§2.1.0 | **A未運用・作成=UL済・画像保全・台帳終了まで** | **Git**: revert |
| 2026-08-11 | sync／schedule_class／DEALS§2.1.1 | **B台帳保全**（予定消禁止）／Smile復元・BF非自動 | **Git**: revert |
| 2026-08-11 | `schedule_class`／sync／build／`TimeSaleP1bMenus.js`／DEALS docs | **P1b: ②SKU行正本・カスタム提出・UL固定文・再build** | **Git**: revert。要 clasp push（広告運用） |
| 2026-08-11 | `TimeSaleP1bMenus.js`／コード.jsメニュー／DEALS§10.9.1・P1b HUMAN | **P1b: ②確認→Cursor→③→UL手順メニュー**。公式即登録 | **Git**: revert。要 clasp push（広告運用） |
| 2026-08-11 | `a_track.py`／`lane_a_send`／sheet_schema／build `--sku`／DEALS docs | **A実施フラグ＋ログ参照**／1SKU提出 | **Git**: revert。列は手で残して可 |
| 2026-08-11 | `sheet_schema`／`sheets_io`／`TimeSalePriceRecovery.js`／DEALS§2.2 | **マスタヘッダ色分け**（戻し=琥珀） | **Git**: revert。シート背景は手でクリア可。要 clasp push（広告運用） |
| 2026-08-11 | `TimeSalePriceRecovery.js`／`price_recovery_logic`／sheet_schema／DEALS§10.12 | **戻し列＋GAS都度提案（即書き込み）** | **Git**: revert。要 clasp push（広告運用） |
| 2026-08-11 | `points_fetch`／`mail_qty_confirm`／`detect_new_schedules`／ops／DEALS§9.7・10.10-11 | **ポイントAPI退避・数量確認メール(スプシリンク)・新Sale差分** | **Git**: revert |
| 2026-08-11 | `points_*`／`sheet_schema`／DEALS§10.10 | **ポイント列意味**: 期間中／セール前退避／restoreモード | **Git**: revert。シート列は手で戻し可 |
| 2026-08-11 | `points_logic`／`points_send`／`sheet_schema`／sync／DEALS要件§10.10 | **Points Phase0**: マスタ列＋差分TSV／フィード。Phase1はメモのみ | **Git**: revert。マスタ列は手で残して可 |
| 2026-08-11 | `ops_status`／`revise_b_qty`／`lane_a_patch` | 運用ダッシュ・数量改定・AパッチJSON | **Git**: revert |
| 2026-08-11 | `tools/amazon_deals_bulk/*`／広告GAS／DEALS要件§2 | **シート改訂**: マスタ薄型・実行・A別タブ。sync＋P1b | **Git**: revert。シートは手で戻し可 |
| 2026-08-11 | `tools/amazon_deals_bulk/*`／広告GAS Cursor指示／DEALS要件・P1b HUMAN_RUN | **販促P1b**: ②人DL→Python提出xlsx→③。GASは指示のみ | **Git**: revert。③のxlsxは手削除 |
| 2026-08-11 | `広告運用GAS/コード.js`（clone push済） | **販促P0**: `bootstrapTimeSaleSheetsP0`＋メニュー。マスタ／ログ／分析タブ作成 | **Git**: revert 該当差分。シートは手削除可。要 clone 経由 clasp push |
| 2026-08-10 | `AmazonSpapiPut.js`／`AmazonApprovalExport.js`／`コード.js`／CURRENT_PHASE／D_ENTRY | **本番常時ONセット**: PUT／ALLOW_PROD／LV4／SCサマリ未設定=true。MASTER_QTY等はOFF。要 clasp push＋旧falseキー削除 | **Git**: revert。緊急停止は明示 false |
| 2026-08-10 | `tonmana_palette.py`／compose／`コード.js` B-④／SUB docs | **ベース色 PoC 3択**（beige/warm_white/soft_gray）B-④＋`--base-color`。背景のみ。要 clasp push | **Git**: revert |
| 2026-08-10 | `コード.js` B-④／`review_feedback_templates.py`／review_loop | **目視チェック必須パネル**＋再生成コメントテンプレ（偽物感等）。要 clasp push | **Git**: revert |
| 2026-08-10 | `export_sub_images_for_rakuten_matrix.py`／compose | **目視=品番キー×最大10（pick=ab）**。全セット子FO禁止。`--to-checked-children` は出品CKのみ | **Git**: revert |
| 2026-08-10 | `photo_realism_rules.py`／compose fal | **CAMERA_LOOK 案A**（Canon R5 50mm f/1.8・creamy bokeh・輪郭ソフトネス）。目視10枚/全子FOは後続 | **Git**: revert |
| 2026-08-10 | `sub_image_ai_compose_poc.py` | **auto-export前に run_meta を書く**（順序バグで SystemExit→バッチ停止）。SystemExit を捕捉して次JAN継続 | **Git**: revert |
| 2026-08-10 | `コード.js` B-④ | **`finishB4AfterTruthBind` 公開名化**（末尾`_`削除）。`google.script.run` private 固着修正。要 clasp push | **Git**: revert |
| 2026-08-10 | `コード.js` B-④ | **診断ログ diag=v1**（ms刻み・`finishB4`・Drive resolve via・クライアント時刻）。要 clasp push → ログ貼付で原因切り分け | **Git**: revert |
| 2026-08-10 | `コード.js` B-④ | 紐付け後の指示を**同一ダイアログ内表示**（`google.script.run`→`showModalDialog`禁止でハング修正） | **Git**: revert。要 clasp push |
| 2026-08-10 | `コード.js`／`photo_realism_rules.py`／`rakuten_image_names.py`／export／compose／SUB docs | **B-④ JAN↔正本人間紐付け**／写真実写ルール進化ファイル／`{sku}_{pattern}_subN` | **Git**: revert。要 clasp push。`SET_MAIN_AMAZON_BASE_FOLDER_ID` 任意 |
| 2026-08-09 | `work_paths.py`／`master_sets.py`／`export_sub_images_for_rakuten_matrix.py`／`sub_image_ai_compose_poc.py`／`コード.js`／SUB docs | **1話完結 auto-export**: compose後に楽天フォルダへ直出し。人間目視フォルダは1つ。JAN→出品CK子SKU解決 | **Git**: revert。`--no-auto-export`。要 clasp push |
| 2026-08-09 | `package_truth.py`／`sub_image_ai_compose_poc.py`／`sub_image_lp_themes.py`／`sub_image_review_loop.py`／`コード.js`／SUB docs | **PACKAGE_TRUTH必須・LOCK強化・文字量上限・SEO・人間レビュー再生成**。B-④正本確認。Vision QA／ベース色は対象外 | **Git**: revert。要 clasp push |
| 2026-08-09 | `コード.js`／SUB HUMAN_RUN／B3 docs | **B-④**: サブ採用CK JAN→Python／Cursor指示ダイアログ（非合成） | **Git**: revert。要 clasp push（メニュー反映） |
| 2026-08-09 | `sub_image_b3_curate.py`／`export_sub_images_for_rakuten_matrix.py`／`コード.js`／サブ画像楽天 docs | **サブ画像楽天本線**: B-③参照列＋compose→subN export＋E2E手順。AmazonはU4 REUSE | **Git**: revert。要 clasp push。`RAKUTEN_IMAGE_MAIN_AUTOBIND_ENABLED=false` |
| 2026-08-09 | `コード.js`／`master_sets.py`／`rakuten_image_names.py`／`export_sub_images_for_rakuten_matrix.py`／楽天SKU docs | **楽天セット数→SKU紐付け転用**: ファイル名子SKU→楽天メイン1／`_subN`→サブ。マスタでN解決。Vision Nは対象外 | **Property** `RAKUTEN_IMAGE_MAIN_AUTOBIND_ENABLED=false`。**Git**: revert。要 clasp push |
| 2026-08-09 | `sub_image_lp_themes.py`／compose／instruction_report／HUMAN_RUN | AI合成v3: 心理順スロット5×A/B・page分類・想像FO・OpenAI medium本線・pkgロック。パーツ分類は将来メモ | **Git**: 当該ファイル revert |
| 2026-08-09 | `sub_image_fal_edit_compare_poc.py`／HUMAN_RUN | fal競合改変4本比較（背景/光/湯気/文字色・商品色ロック） | **Git**: 当該ファイル revert |
| 2026-08-09 | `sub_image_ai_compose_poc.py`／HUMAN_RUN | 段階1: hybrid振り分け（T03→fal文字禁止／他→Gemini） | **Git**: 当該ファイル revert |
| 2026-08-09 | `fal_image.py`／`sub_image_fal_poc.py`／compose／HUMAN_RUN | fal.ai（FLUX Kontext/Schnell）格安画像PoC | **Git**: 当該ファイル revert。キーは secrets |
| 2026-08-09 | `sub_image_ai_compose_poc.py`／`sub_image_lp_themes.py`／`sub_image_instruction_report.py`／HUMAN_RUN | AI合成v2: 競合実在テーマ／ベージュ／AI≤50%／日本語指示ボード | **Git**: 当該ファイル revert |
| 2026-08-09 | `sub_image_ai_compose_poc.py`／`openai_image.py`／HUMAN_RUN | サブ画像をGemini+OpenAI最新画像モデルで再合成（パーツ流用≥80%／新規≤20%） | **Git**: 当該ファイル revert |
| 2026-08-09 | `コード.js` B-③／`sub_image_b3_curate.py`／`sheets_rw.py`／docs | B-③右列サブ採用レ点＋採用のみログ＋再実行復元。旧確認シート削除 | **Git**: revert。要 clasp push。ログ／旧シートは手削除可 |
| 2026-08-09 | `sub_image_b3_curate.py`／`sheets_rw.py`／`sub_image_part_poc.py`／HUMAN_RUN | JAN別仕分け＋シート「サブ画像競合候補（人間確認）」。メイン/無関係除外デフォルト | **Git**: revert。シートは手動削除可 |
| 2026-08-09 | `b3_comp_catalog.py`／`sub_image_intent.py`／`sub_image_parts.py`／`sub_image_part_poc.py`／HUMAN_RUN | B-③連動: 文字読取で選定・パーツ提案・安全BG合成（文字改変なし）。商品任意 | **Git**: 当該ファイル revert |
| 2026-08-09 | `sub_image_poc.py`／`sub_image_prompts.py`／`_sub_bg_photos/`／SUB IMAGE HUMAN_RUN | サブ画像10種（実写背景Pillow＋競合掛け合わせ＋シーンAI）。MAIN相当なし | **Git**: 当該ファイル revert。背景JPGは `_sub_bg_photos` 削除可 |
| 2026-08-09 | `sub_image_poc.py`／`sub_image_prompts.py`／SUB IMAGE HUMAN_RUN | サブ画像PoC（競合改変＋自社フル・注釈ボード・テスト出力のみ） | **Git**: 当該ファイル削除／revert |
| 2026-08-09 | `コード.js`（画像キー／画素スコア／Keepa列挙／B-③ push）／B3 docs | B-③同一画像は高画素を残す（AmazonメディアID＋SL/_ex） | **Git**: revert／clasp push |
| 2026-08-09 | `コード.js`（parseCrossMall／pick／B-③）／B3 docs | **案A**: 統合セット数不明でも楽天Yahoo競合画像を取得 | **Git**: revert／clasp push |
| 2026-08-09 | `コード.js`（KeepaログJAN／B-③ Amazon段0〜2）／B3 docs／RESEARCH | B-③: ログ◎→貼付◎→マスタ親子。0件次段。Amazon0でも楽天Yahoo。ログJAN=AI | **Git**: revert／clasp push |
| 2026-08-09 | `landscape_layout.py`（新）／`amazon_paste.py`／`amazon_paste_batch.py`／`layout_rules.json`／`コード.js`／docs | 横長合格固定をC③実装。楽天N=1をC指示に反映 | **Git**: 当該ファイル revert／clasp push でGAS反映 |
| 2026-08-09 | `rakuten_badge.py`／`compose_set_main.py`／RULES／HUMAN_RUN／landscape annot | 楽天N=1生成許可。横長N1中央+Octas。N2+合格記録 | **Git**: 当該ファイル revert |
| 2026-08-09 | `SET_MAIN_LAYOUT_RULES.md`／`_export_landscape_rule_annots.py` | 横長: N3/4階段横進み復元＋Octas。N≥5案A（右辺接触廃止） | **Git**: 当該docs/script revert |
| 2026-08-09 | `SET_MAIN_LAYOUT_RULES.md` §1.2／`_export_landscape_rule_annots.py` | 横長: N3/4枠ピン・N≥5 hero三辺＋半端行中央寄せ。注釈再出力 | **Git**: 当該docs/script revert |
| 2026-08-08 | `SET_MAIN_LAYOUT_RULES.md` §1.2／HANDOVER | **横長セットMAINルール草案**（見本154–158・実装前） | **Git**: 当該docs revert |
| 2026-08-08 | `AmazonSpapiPut.js`／`コード.js`／PUT docs | **PUT max_items 既定 5→10**（Property未設定時。ハード上限50は維持） | **Git**: revert。運用で下げるなら `APPROVAL_AMAZON_SPAPI_PUT_MAX_ITEMS` を再設定 |
| 2026-08-08 | `AmazonCategoryPt.js`／`コード.js`／P4b承認・HUMAN_RUN | **P4b-d**: ASIN検証・複数多数決・JAN/SHELF重み投票・Browse無しPT禁止＋Dエラーalert | **Git**: revert。`APPROVAL_AMAZON_P4B_PT_MAX_ASINS` 削除可 |
| 2026-08-08 | `AmazonImageMatrixExport.js`／`コード.js`／C・U2 docs | **U2-ε MAIN自動**: 子SKU名一致→BX自動投入（既存非上書き・親一致/余りは対象外） | **Property** `AMAZON_IMAGE_MAIN_AUTOBIND_ENABLED=false`。**Git**: revert |
| 2026-08-08 | `AmazonDriveImageExport.js`／`コード.js`／U4 HUMAN_RUN | **U4途中切れ対策**: URL充足スキップ・1SKU例外継続・Putリトライ・時間スライス＋トリガー再開 | **Property** `AMAZON_U4_FORCE_REUPLOAD`／`AMAZON_U4_SLICE_MS` 削除可。**Git**: revert。残留トリガーは `runAmazonU4ResumeFromTrigger` を削除 |
| 2026-08-08 | `bind_amazon_base_to_parents.py`／`amazon_paste_batch`／SET_MAIN HUMAN_RUN | **01白抜き↔楽天メインVision自動紐付け**（`{親SKU}_単体.png`・複数商品量産） | **Git**: revert。リネームは `BASE_BIND_*.json` の old_name で手戻し |
| 2026-08-08 | `layout_rules`／`portrait_layout`／`amazon_paste`／SET_MAIN RULES・HUMAN_RUN | **N=3細長合格確定**: Q(四隅上辺中点)→M＋緑法線上+35px。C②→07→D | **Git**: revert。ノブ `unit0NudgeUpAlongNormalPx=0` でも旧挙動に近い |
| 2026-08-08 | `portrait_layout`／RULES | **N=3細長**: Q＝四隅上辺(TL–TR)中点→円弧中点M（AABB不可） | **Git**: revert |
| 2026-08-08 | `portrait_layout`／RULES | **N=3細長**: Q＝オレンジAABB上辺中点→円弧中点M | **Git**: revert |
| 2026-08-08 | `portrait_layout`／RULES | **N=3細長**: Q＝オレンジ外枠(ProductQuad)上辺中点→円弧中点M | **Git**: revert |
| 2026-08-08 | `portrait_layout`／RULES | **N=3細長**: unit1上辺中点Q＝円弧中点M（法線∩弧） | **Git**: revert |
| 2026-08-08 | `portrait_layout`／`layout_rules`／RULES | **N=3細長**: unit1上辺左右中点＝円弧中点の法線上 | **Git**: revert |
| 2026-08-08 | `portrait_layout`／`layout_rules`／RULES | **N=3細長**: hero枠内＋unit0 AABB中心中点（見た目バランス） | **Git**: revert |
| 2026-08-08 | `portrait_layout`／`layout_rules`／`octas_prep`／`amazon_paste`／RULES | **N=3細長**: 中角8.8°（頂X中点）＋composeでOctas未傾き時に8°適用 | **Git**: revert（Pythonのみ） |
| 2026-08-08 | `portrait_layout.py`／`layout_rules.json`／SET_MAIN RULES | **N=3細長**: hero AABB下辺＝枠下辺（扇剛体平行移動／`pinHeroFrameBottom`） | **Git**: revert（Pythonのみ・clasp不要） |
| 2026-08-08 | `portrait_layout.py`／`layout_rules.json`／SET_MAIN RULES | **N=3分岐**: 細長→top_arc_legacy＋高さ≤92%（N=4同型）／幅広→upright現行 | **Git**: revert（Pythonのみ・clasp不要） |
| 2026-08-08 | `portrait_layout.py`／`layout_rules.json`／SET_MAIN RULES | **N=4分岐見直し**: 細長(H/W≥1.75・缶含む)→**n4_legacy頂円弧**／幅広→upright_quad現行 | **Git**: revert（Pythonのみ・clasp不要） |
| 2026-08-08 | `amazon_paste.py`／`layout_rules.json`／SET_MAIN RULES | **N≥5 Octas**: unit非被覆＋下枠接触＋**hero隠し≤5%で縮小**／upright_quad素材1枚化 | **Git**: revert（Pythonのみ・clasp不要） |
| 2026-08-08 | `layout_rules.json`／`portrait_layout.py`／`コード.js`／`work_paths.py`／SET_MAIN RULES・HUMAN_RUN・LV4／HANDOVER | **縦長タイプ合格固定**（N=2〜≥5・Cメニュー②登録・unit3 BR枠下） | **Git**: revert。GASは clasp push 前ならローカル戻し |
| 2026-08-05 | `portrait_layout.py`／`layout_rules.json`／SET_MAIN RULES・HUMAN_RUN | **縦長N=3合格固定（頂円弧）＋N=4拡張初版** | **Git**: revert。`portraitTilt.n3/n4` を戻す |
| 2026-08-05 | `portrait_layout.py`／`amazon_paste*.py`／`layout_rules.json`／`コード.js`／SET_MAIN docs | **縦長パターン（斜めファン）** | **Git**: revert。GASは clasp push 前ならローカル戻し |
| 2026-08-05 | `c1_packaged.py`／C1 HUMAN_RUN／HANDOVER | **定価空時のみ販売価格フォールバック**（B3はマスタ定価で再PACKAGED） | **Git**: revert `resolve_list_price`。03の旧xlsm（timestampなし）は販売価格誤反映のため使わない |
| 2026-08-02 | `layout_rules`／`rakuten_layer` 3段組版／`master_sets` バリエーション単位／RULES・HUMAN | **楽天Python本線**（金丸数値組版＋単位ヘッダ連携） | **Git**: revert |
| 2026-08-02 | `tools/set_main_image/ai_*`／`gemini_image`／`model_policy`／AI承認・HUMAN_RUN／PHASE／LEDGER | **セットMAIN AI PoC**（Nano Banana＝通常Flash最新・PRO禁止・見本03/04） | **Git**: revert。secrets/はコミットしない |
| 2026-08-02 | `LV4_SET_MAIN_IMAGE_PHASE_AI_APPROVAL`／PHASE／LEDGER | **セットMAINを見本参照AI生成へ振替提案**（Pillow本線凍結） | **Git**: 当該docs revert |
| 2026-08-02 | `tools/set_main_image/*`／`D_MENU_SET_MAIN_IMAGE_*`／承認／PHASE／HANDOVER／LEDGER | **セットMAIN Phase A 実装**（Amazonコラージュ＋楽天数字レイヤ・07出力） | **Git**: revert。07の生成jpgは人手削除 |
| 2026-08-02 | `LV4_SET_MAIN_IMAGE_PHASE_A_APPROVAL`／PHASE／LEDGER | **セットMAIN Phase A 方針ロック**（数字レイヤ・07/C・一括・品質でリサーチ切替可） | **Git**: 当該docs revert |
| 2026-08-02 | `LV4_SET_MAIN_IMAGE_PHASE_A_APPROVAL`／PHASE／LEDGER | **セットMAIN Phase A 承認ドラフト**（Pillow合成・生成AI不使用・B/C後回し） | **Git**: 当該docs revert |
| 2026-08-02 | 台帳／PHASE／ROADMAP／REMAKE／HANDOVER／LEDGER | **七味出品中**（48h待ち削除）＋D内レ点 clasp push済反映 | **Git**: 当該docs revert |
| 2026-08-02 | `コード.js`／`AmazonApprovalExport.js`／`LV4_D_REMAKE_*`／Facade／PHASE／HANDOVER／LEDGER | **D内レ点実装**（失敗後再GENERATED・案A・確認後に最新openのみ UPLOAD_FAILED） | **Git**: revert。Z 21-④は従来どおり |
| 2026-08-02 | `LV4_D_REMAKE_MENU_APPROVAL`／MAP HUMAN_RUN／Facade§6／SHELF・PHASE／HANDOVER／LEDGER | **D内レ点方針ロック（案A）**＋MAP sync HUMAN_RUN §0b 固め。コードなし | **Git**: 当該docs revert |
| 2026-08-02 | `c1_packaged.py`／`コード.js`（D手渡し人間手順）／MAP・LANE_B台帳／SHELF HUMAN_RUN／LEDGER | **browse=Node IDのみ**（90194/90225対策）＋人間手順を成功/失敗で分離（失敗=05・08禁止） | **Git**: revert |
| 2026-08-02 | `c1_packaged.py`／`コード.js`／MAP・LANE台帳／LEDGER | **browse=BrowsePath（プルダウン表記）**＋人間手順成功/失敗分離。数値ID単独禁止 | **Git**: revert |
| 2026-08-02 | `LV4_MAP_SHEET_JSON_SYNC_*`／`sync_map_sheet_to_column_json.py`／`push_map_attr_patches_to_sheet.py`／`append_map_sheet_error.py`／`MAP_SC_ERROR_LEDGER.md`／PHASE／HANDOVER／LEDGER | **正本=sheet／派生=MD／実行=JSON（sheet→生成）** | **Git**: revert。JSON は `.bak_map_sync` から戻す |
| 2026-08-02 | `c1_packaged.py`／`food_fish_grocery_column_map.json`／MAP生成／LEDGER | **缶飯数量ポリシー**（入数/ユニット=セット缶数、重量=サイズg、AT=size、色・形態・温度未出力） | **Git**: revert。03の該当PACKAGEDを旧版に戻す |
| 2026-08-02 | `food_fish_grocery_column_map.json`／`create_amazon_mapping_sheet.py`／C1 FOOD列マップ／SHELF承認／LEDGER | **GROCERYテーマ=`サイズ`**（純正プルダウンのみ。`SET_NAME`廃止） | **Git**: revert。旧xlsmは03のタイムスタンプ付きを使わない |
| 2026-08-02 | `LV4_SHELF_BROWSE_CATALOG_*`／`c1_shelf_browse_extract.py`／`shelf_browse_catalog.json`／sync MAP／`AmazonCategoryPt.js`／`コード.js`／T2・P4b・PHASE／HANDOVER／LEDGER | **SHELF Browse網羅＋Nodeルーティング**（缶飯GROCERY・MEATエイリアスはフォールバック） | **Git**: revert。Drive catalog／registry 旧版 |
| 2026-08-02 | `shelf_registry.json`／`food_fish_grocery_column_map.json`／指紋／`AmazonCategoryPt.js`／`コード.js`／PHASE／HANDOVER／LEDGER | **GROCERY本線**（FOOD_FISH_GROCERY・Catalog MEAT→GROCERYエイリアス・Dゲート書換） | **Git**: revert。Drive `shelf_registry.json` を旧版に戻す |
| 2026-08-02 | `c1_packaged.py`／`food_seasoning_column_map.json`／`create_amazon_mapping_sheet.py`／C1 FOOD列マップ／P4b承認／PHASE／HANDOVER／LEDGER | **C1 FOOD: PT/browse必須・既定禁止・FOOD許可・ハイライトB（楽天→Yahoo→箇条書き①）・HJ型番** | **Git**: revert 当該差分 |
| 2026-08-02 | `コード.js`／`AmazonCategoryPt.js`／D新規ゲート承認・HUMAN_RUN／Facade／P4b／PHASE／HANDOVER／LEDGER | **D新規ゲート＋Cursor手渡し実装**（Drive棚 File ID） | **Git**: revert。Property `AMAZON_SHELF_REGISTRY_FILE_ID` 削除可 |
| 2026-08-02 | `AmazonSpapiPut.js`／`コード.js`／D_ENTRY／デュアル承認／レ点本線承認／HANDOVER／LEDGER | **相乗り通常運用=prod直**（SKU空でも生成・成功後保存。dry_run任意） | **Git**: revert 当該差分 |
| 2026-08-02 | `コード.js`／`AmazonDriveImageExport.js`／C画像コース HUMAN_RUN／U2要件／HANDOVER／LEDGER | **C表示名をAma新カタログ①／②へ明確化＋U4対象SKU限定の兄弟PT URL補完** | **Git**: revert 当該差分。補完済みURLは必要な行だけ人手で空欄へ戻す |
| 2026-08-01 | `AmazonImageMatrixExport.js`／C画像コース HUMAN_RUN／U2要件／PHASE／HANDOVER／LEDGER | **Amazon候補を1子SKU＝1枚に**（全行タイル廃止・名前一致優先） | **Property**: `AMAZON_IMAGE_CANDIDATE_TILE_ALL_ENABLED=true` で旧表示。**Git**: revert |
| 2026-08-01 | `コード.js`（`runBatchExportAmazonFacade`）／D新規ゲート承認§2.1／D_ENTRY／HUMAN_RUN／PHASE／HANDOVER／LEDGER | **相乗りASIN空 soft skip**（相乗りのみでも行スキップ・全体継続） | **Git**: revert 当該差分 |
| 2026-08-01 | `LV4_D_NEW_PT_SHELF_GATE_…`／`D_MENU_D_NEW_PT_SHELF_GATE_…`／Facade§6／P4b-c／PHASE／HANDOVER／LEDGER | **D新規ゲート＋Cursor手渡し方針ロック**（docsのみ・実装未） | **Git**: 当該docs revert |
| 2026-08-01 | B-T2承認／HUMAN_RUN／T0・T1誘導／PHASE／HANDOVER／LEDGER／ROADMAP | **B-T2§4ロック**（複合複数明示・三点スキップ・実装は実需後） | **Git**: 当該docs revert |
| 2026-08-01 | `LV4_LANE_B_BULK_TEMPLATE_T2_…`／`D_MENU_LANE_B_BULK_T2_…`／T0・T1誘導／PHASE／HANDOVER／LEDGER／ROADMAP | **B-T2方針ドラフト**（コードなし・複合優先・実需起動） | **Git**: 当該docs revert |
| 2026-08-01 | `コード.js`／`AmazonApprovalExport.js`／P0承認／D_ENTRY／PHASE／HANDOVER／LEDGER | **D送信在庫: 新規もMASTER時qty内訳確認**（相乗りと同水準・UI明示） | **Git**: revert 当該差分 |
| 2026-08-01 | 台帳／C1 §0b／D_ENTRY §1f／PHASE／HANDOVER／LEDGER | **七味再送サマリ確定**（100521のみ・99016解消） | **Git**: 当該docs revert |
| 2026-08-01 | `コード.js`（`createZSplitMenu`） | **Zメニュー連番サブ化**（1〜10／11〜17／18〜21・21内分割。関数名・番号不変） | **Git**: revert |
| 2026-08-01 | `c1_bulk_fill_by_name.py`／T1_PROD承認／HUMAN_RUN／PHASE／HANDOVER／LEDGER／ROADMAP | **B-T1 prod実装**（`--mode prod`・三点スキップ・スモークOK） | **Git**: revert 当該差分 |
| 2026-08-01 | `LV4_LANE_B_BULK_TEMPLATE_T1_PROD_…`／`D_MENU_LANE_B_BULK_T1_PROD_…`／親T1／PHASE／HANDOVER／LEDGER／ROADMAP | **B-T1 prod第2段方針ドラフト**（コードなし） | **Git**: 当該docs revert |
| 2026-08-01 | `c1_bulk_shelf_lookup.py`／`c1_bulk_fill_by_name.py`／`c1_bulk_name_map.py`／`shelf_registry.json`／`food_seasoning_column_map.json`／B-T1 docs | **B-T1実装**（棚引き＋項目名 dry_run・三点スキップ・スモークOK） | **Git**: revert 当該ツール＋docs。prod未 |
| 2026-08-01 | B-T1承認／HUMAN_RUN／PHASE／HANDOVER／LEDGER | **B-T1に棚引き・DL要否案内を追記** | **Git**: 当該docs revert |
| 2026-08-01 | `LV4_LANE_B_BULK_TEMPLATE_T1_…`／`D_MENU_LANE_B_BULK_T1_…`／PHASE／HANDOVER／LEDGER／ROADMAP | **B-T1方針ドラフト**（項目名マップ・複合SEASONING・コードなし） | **Git**: 当該docs revert |
| 2026-08-01 | `c1_bulk_fingerprint.py`／B-T0承認／HUMAN_RUN／PHASE／HANDOVER／LEDGER／ROADMAP | **B-T0実装**（09指紋→05・三点スキップ・スモークmatch） | **Git**: revert ツール＋docs |
| 2026-08-01 | B-T0承認／HUMAN_RUN／PHASE／HANDOVER／LEDGER／ROADMAP | **B-T0入口を09に確定** | **Git**: 当該docs revert |
| 2026-08-01 | `LV4_LANE_B_BULK_TEMPLATE_T0_…`／`D_MENU_LANE_B_BULK_T0_…`／PHASE／HANDOVER／LEDGER／ROADMAP | **B-T0承認ドラフト**（指紋差分・`04`集約・コードなし） | **Git**: 当該docs revert |
| 2026-08-01 | P4b／C1 HUMAN_RUN／PHASE／HANDOVER／LEDGER／ROADMAP | **P4b-b合格**（HERB→C1 PT=SEASONING）＋ローカルPython=Agent実行明記 | **Git**: 当該docs revert |
| 2026-08-01 | P4b承認／HUMAN_RUN／PHASE／HANDOVER／LEDGER／ROADMAP | **P4b-a合格・P4b-b=HERBフォールバック手順** | **Git**: 当該docs revert |
| 2026-08-01 | `AmazonCategoryPt.js`／P4b承認／HUMAN_RUN／PHASE／HANDOVER／LEDGER／ROADMAP | **P4bマスタ競合版実装**（多数決廃止・WARNのみ） | **Git**: revert。Property OFF |
| 2026-08-01 | P4b承認§2／HUMAN_RUN／PHASE／HANDOVER／LEDGER／ROADMAP | **P4b§2マスタ競合改定**（多数決廃止・WARNのみ・コード未） | **Git**: 当該docs revert |
| 2026-08-01 | `AmazonCategoryPt.js`／P4b承認§2／HUMAN_RUN／PHASE／HANDOVER／LEDGER／ROADMAP | **P4b多数決改定**（◎上位5・同票は売上1位・SEASONING優先廃止） | **Git**: revert。Property OFF |
| 2026-08-01 | `AmazonCategoryPt.js`／`AmazonSpapiPut.js`／`コード.js`／`c1_packaged.py`／column_map／P4b承認／HUMAN_RUN／PHASE等 | **P4b実装**（21-⑱・空セル書込・C1マスタ優先・三点スキップ） | **Git**: revert。Property OFF |
| 2026-08-01 | P4b承認§2／HUMAN_RUN／PHASE／HANDOVER／LEDGER／ROADMAP | **P4b方針ロック**（Keepa既存browse・提案HERB可・C1本線SEASONING/HPC） | **Git**: 当該docs revert |
| 2026-08-01 | `LV4_AMAZON_CATEGORY_PT_P4B_…`／`D_MENU_P4B_…`／PHASE／HANDOVER／LEDGER／ROADMAP／P4a | **P4b承認包起草**（方針待ち・コードなし） | **Git**: 当該docs revert |
| 2026-08-01 | Phase2承認§5／HUMAN_RUN§3／PHASE／HANDOVER／LEDGER／ROADMAP／D_ENTRY | **デュアル Phase2検収OK**（dry `…f372b8`／`…40d85e`・prod `…8fcbc2`／`…cc7c72`） | **Git**: 当該docs revert |
| 2026-08-01 | `AmazonSpapiPut.js` | **`amazonSpapiPutOfferSellerSkuHeader_` 復元**（Phase2編集で消失→`ReferenceError`。系統→保存列名） | **Git**: revert |
| 2026-08-01 | `コード.js`／`AmazonSpapiPut.js`／`LV4_DUAL_OFFER_PHASE2_…`／`D_MENU_DUAL_OFFER_PHASE2_…`／PHASE／HANDOVER／LEDGER／ROADMAP／D_ENTRY | **デュアル Phase2実装**（チェックUI・順次PUT・部分成功） | **Git**: revert。UIはラジオに戻る |
| 2026-08-01 | `LV4_LANE_B_…`／`LANE_B_SC_ERROR_LEDGER`／`D_MENU_LANE_B_LEDGER_…`／PHASE／HANDOVER／LEDGER／ROADMAP／P2・C1・D_ENTRY | **レーンB台帳初版＋シード** | **Git**: 当該docs revert |
| 2026-08-01 | A3承認／HUMAN_RUN／PHASE／HANDOVER／LEDGER／ROADMAP／D_ENTRY | **A3 dry／prod検収OK**（`…49a49e`／`…f677a3`） | **Git**: 当該docs revert |
| 2026-08-01 | `LV4_LANE_A3_…`／`D_MENU_LANE_A3_HUMAN_RUN`／PHASE／HANDOVER／LEDGER／ROADMAP／D_ENTRY | **A3承認包・HUMAN_RUN起草**（実機待ち・原則コードなし） | **Git**: 当該docs revert |
| 2026-08-01 | A2承認／HUMAN_RUN／PHASE／HANDOVER／LEDGER／ROADMAP／D_ENTRY | **A2検収OK記録** | **Git**: 当該docs revert |
| 2026-08-01 | `LV4_LANE_A2_…`／`D_MENU_LANE_A2_HUMAN_RUN`／PHASE／HANDOVER／LEDGER／ROADMAP／D_ENTRY | **A2承認包・HUMAN_RUN起草**（原則コードなし・実機待ち） | **Git**: 当該docs revert |
| 2026-08-01 | デュアル承認§6／D_ENTRY§1b2／A1 HUMAN_RUN／PHASE／HANDOVER／LEDGER／ROADMAP | **デュアル prod検収OK**（`…4ed30e`／`…eb2511`） | **Git**: 当該docs revert |
| 2026-08-01 | デュアル承認§6／D_ENTRY§1b2／A1 HUMAN_RUN／PHASE／HANDOVER／LEDGER／ROADMAP | **デュアル dry_run実機OK記録**（`…8fa79e`／`…d6ed67`。prod未） | **Git**: 当該docs revert |
| 2026-08-01 | `AmazonSpapiPut.js`／`コード.js`／`AmazonApprovalExport.js`／デュアル承認／D_ENTRY・A1 HUMAN_RUN／PHASE／HANDOVER／LEDGER／ROADMAP | **デュアル Phase1実装**（系統別SKU列・三点スキップ） | **Git**: revert。マスタの `_FBA` 列は人手削除可 |
| 2026-08-01 | A1 HUMAN_RUN／属性・A1承認／PHASE／ROADMAP／HANDOVER／D_ENTRY／LEDGER | **A1検収OK記録**（FBA dry_run／prod runId） | **Git**: 当該docs revert |
| 2026-08-01 | `AmazonSpapiPut.js`／`コード.js`／属性承認／A1 HUMAN_RUN／PHASE／HANDOVER／LEDGER | **FBA compliance属性実装**（電池・危険物既定＋失敗時おすすめUI） | **Git**: revert。Property `…FBA_COMPLIANCE_ATTRS=false` |
| 2026-08-01 | `LV4_A1_FBA_COMPLIANCE_ATTRS_…`／PHASE／ROADMAP／HANDOVER／LEDGER | **FBA属性方針承認**（§3折衷・三点スキップ。実装待ち） | **Git**: 当該docs revert |
| 2026-08-01 | `LV4_A1_FBA_COMPLIANCE_ATTRS_APPROVAL`／A1 HUMAN_RUN／PHASE／ROADMAP／HANDOVER／LEDGER | **A1対比記録＋FBA属性承認ドラフト**（コードなし） | **Git**: 当該docs revert |
| 2026-08-01 | `LV4_DUAL_OFFER_MFN_FBA_APPROVAL`／A1 HUMAN_RUN／PHASE／ROADMAP／HANDOVER／LEDGER | **デュアル Phase1方針ロック**＋A1暫定NF空七味。コードなし | **Git**: 当該docs revert |
| 2026-08-01 | `LV4_LANE_A1_…`／HUMAN_RUN／PHASE／ROADMAP／HANDOVER／LEDGER | **A1方針承認**（prodまで・三点スキップ。実機待ち） | **Git**: 当該docs revert |
| 2026-08-01 | `LV4_LANE_A1_FBA_…APPROVAL`／`D_MENU_LANE_A1_FBA_HUMAN_RUN`／PHASE／ROADMAP／HANDOVER／LEDGER | **レーンA1承認包ドラフト**（相乗りFBA検収。コード原則なし） | **Git**: 当該docs revert |
| 2026-08-01 | `LV4_P2_…APPROVAL` §7.4／PHASE／ROADMAP／HANDOVER／P2 HUMAN_RUN／LEDGER | **方針ロック**: 既存API本線(A)／新規xlsm取り貯め(B)／新規JSONゲート後(C)。コードなし | **Git**: 当該docs revert |
| 2026-08-01 | `LV4_P2_DC_123_…APPROVAL`／P2 HUMAN_RUN／PHASE／ROADMAP／HANDOVER／LEDGER | **P2調査承認＋結論**（①②API不可／③温存。コードなし） | **Git**: 当該docs revert |
| 2026-08-01 | `LV4_P2_DC_123_INVESTIGATION_APPROVAL`／`D_MENU_P2_DC_HUMAN_RUN`／PHASE／ROADMAP／HANDOVER／LEDGER | **P2調査承認包起草**（Dc①②③。コードなし） | **Git**: 当該docs revert |
| 2026-08-01 | PHASE／ROADMAP／HANDOVER／LEDGER／P1承認・HUMAN_RUN | **P1検収OK＋吉野家承認済**記録（コードなし） | **Git**: 当該docs revert |
| 2026-08-01 | `コード.js`／`AmazonImageMatrixExport.js`／`AmazonDriveImageExport.js`／P1・C・U2 docs／PHASE／LEDGER | **P1実装**: 02手File禁止の案内・Dゲート／U4／④メッセージ | **Git**: revert。要 clasp push |
| 2026-08-01 | `LV4_P1_FILE_MIN_APPROVAL`／PHASE／ROADMAP／HANDOVER／LEDGER | **P1方針承認**（§2既定・三点スキップ。コードなし） | **Git**: 当該docs revert |
| 2026-08-01 | `LV4_P1_FILE_MIN_APPROVAL`／`D_MENU_P1_HUMAN_RUN`／PHASE／ROADMAP／HANDOVER／LEDGER | **P1承認包起草**（07以降 File最少。コードなし） | **Git**: 当該docs revert |
| 2026-08-01 | `tools/spapi_smoke/spapi_smoke.py`／config.example／P4a承認・HUMAN_RUN／PHASE／ROADMAP／HANDOVER | **P4a読取PoC**: `--poc-category`（Definitions search/get＋Catalog分類。マスタ非書込） | **Git**: revert。フラグ無しなら従来スモークのみ |
| 2026-08-01 | `コード.js`／C承認／C HUMAN_RUN／PHASE／HANDOVER／LEDGER | **CをD型選択モーダル1本へ**（C-0〜C-2はZ互換） | **Git**: revert。トップを `createCImageCourseMenu` サブメニューへ戻す |
| 2026-08-01 | `コード.js`／C承認／C HUMAN_RUN／PHASE／HANDOVER／LEDGER | **C本線 Property一時ON→復元**（常設は07フォルダID。C-2確認＋ScriptLock） | **Git**: revert。復元失敗時は U2/U4 を手動false/削除 |
| 2026-08-01 | `コード.js`／`LV4_C_COURSE_…APPROVAL`／`D_MENU_C_IMAGE_COURSE_HUMAN_RUN`／PHASE／HANDOVER／LEDGER | **Cコース実装**（C-0/1/2・①〜④をZ・E寄せ） | **Git**: revert。メニューは旧C/C-Amazon行に戻す |
| 2026-08-01 | `LV4_C_COURSE_CONSOLIDATION_APPROVAL`／PHASE／HANDOVER／ROADMAP／ファサード§7 | **Cコース統合 方針ロック**（IA・三点スキップ・P1非同梱。コードなし） | **Git**: 当該docs revert |
| 2026-08-01 | `コード.js`／`AmazonSpapiPut.js`／`AmazonApprovalExport.js`／P0承認／D_ENTRY HUMAN_RUN／PHASE／LEDGER | **P0実装**: D在庫0\|マスタ・prod既定・ALLOW_MASTER_QTY・E→D補助/Z互換 | **Property** MASTER_QTY/ALLOW_PROD false。**Git**: revert |
| 2026-08-01 | `LV4_D_P0_THREE_REVIEW_MAJORITY`／P0承認／CHECKBOX／マトリクス／Lv4§3.4／PHASE | **P0三点(B)反映**（マスタqty＝承認②。コードなし） | **Git**: 当該docs revert |
| 2026-08-01 | `AMAZON_DEV_ROADMAP_P0_P4`／`LV4_D_P0_…APPROVAL`／`LV4_AMAZON_CATEGORY_PT_POC_APPROVAL`／PHASE／ファサード／HANDOVER | **Amazon P0〜P4／Dcロードマップ確定**（P4a並列調査・P0三点対象。コードなし） | **Git**: 当該docs revert |
| 2026-07-31 | `AmazonApprovalExport.js`／D_ENTRY HUMAN_RUN §1e | **SCサマリ監視の自動ON/OFF**（GENERATED成功→待ちリスト＋トリガー設置。終端ステータスで削除。上限72h。Property `…_SC_SUMMARY_WAIT`） | **21-⑰**でトリガー＋待ちクリア。**Property** ENABLED=false。**Git**: 対象差分revert |
| 2026-07-31 | `Yahoo.js`／`コード.js`／`YAHOO_OAUTH_REAUTH_HUMAN_RUN.md` | **Yahoo再認証半自動**（認可code貼付→B14更新＋**C14取得日**。残り≤7日でB17へメール。任意日次トリガー。出品本体非改変） | **メニュー**でトリガー削除。**Property** `YAHOO_OAUTH_REDIRECT_URI` 削除可。**Git**: 対象差分revert＋メニュー行削除 |
| 2026-07-31 | `c1_packaged.py`／`food_seasoning_column_map.json` | **SEASONING検索KWは1枠**（SC 99016: generic_keywordは最大1回。空白結合。keyword2〜5列マップ削除） | **Git**: 対象差分をrevert。旧5枠分割に戻る |
| 2026-07-31 | `AmazonDriveImageExport.js`／`コード.js` autoU4／U4 HUMAN_RUN | **楽天サブ→R2流用**（PT参照空かつ非AMAZON_ONLYなら `楽天サブ画像1〜8` を取得→R2→`Amazon PT URL`。D自動U4も楽天サブ不足を検知。MAINファイル無し時は既存MAIN URL維持可） | **Git**: 対象差分をrevert。autoU4判定は旧「ptRef無しならスキップ」に戻る |
| 2026-07-31 | `AmazonDriveImageExport.js`／`コード.js`（D新規）／`c1_packaged.py`／SEASONING列マップ／U4・C1 docs | **サブ画像＋商品の形式＋U4自動化**（U4はREUSE_RAKUTENでもPTをR2へ→`Amazon PT URL`（親へも伝播）、D新規はGENERATED前にU4自動実行、C1はPT URLを24〜31列へ、`item_form` 既定=粉末） | **Property**: `AMAZON_U4_AUTO_IN_D_ENABLED=false` でD自動U4のみ停止（21-⑦手動は従来どおり）。**Git**: 対象差分をrevert。列マップは `item_form`／`other_image1..8` の行を削れば旧挙動 |
| 2026-07-31 | `AmazonApprovalExport.js`／`コード.js` 21-⑮⑯⑰／D_ENTRY HUMAN_RUN §1e | **SC処理サマリ自動記録**（監視フォルダ＋時間主導トリガー→ファイル名から subBatchId→`UPLOADED_OK`／`UPLOAD_FAILED`→`_処理済`退避） | **Property**: `APPROVAL_AMAZON_LV4_SC_SUMMARY_ENABLED=false` で即停止。**21-⑰** でトリガー削除。**Git**: 当該差分revert＋メニュー3行削除。誤記録行はシート上で削除 |
| 2026-07-31 | `AmazonApprovalExport.js`／D_ENTRY HUMAN_RUN | **21-⑭ 証跡おすすめ**（同カテゴリマスタASIN→過去成功文面。OKで記録／Cancelで手入力） | **Git**: 当該差分revert |
| 2026-07-31 | `AmazonApprovalExport.js`／`コード.js` 21-⑭／Lv4正本／D_ENTRY HUMAN_RUN | **GTIN免除証跡の記録メニュー**（レ点カテゴリ検出→人間確認→EXEMPTION追記。既存証跡は追記せず、`*` はProperty＋警告） | **Git**: 対象差分をrevert＋メニュー行削除。誤記録した EXEMPTION 行はシート上で削除（追記のみのため他行に影響なし） |
| 2026-07-30 | `AmazonApprovalExport.js`／レ点本線承認包／Lv4正本／D_ENTRY HUMAN_RUN | **Dレ点新規は在庫>0でもGENERATED**（別カタログのため。承認①経路は従来どおりスキップ／マスタ在庫非改変） | **Property**: `APPROVAL_AMAZON_LV4_CK_ALLOW_IN_STOCK=false`。**Git**: 対象差分をrevert |
| 2026-07-30 | `コード.js`／`AmazonImageMatrixExport.js`／U2 docs | **Amazon画像の旧sheet誤紐付け防止**（U2時は楽天画像0件でもC再生成、②で現在レ点／sheet完全一致ゲート） | **Git**: 対象差分をrevert。`AMAZON_IMAGE_U2_ENABLED=false` |
| 2026-07-30 | `c1_packaged.py`／SEASONING列マップ・指紋／C1 docs | **C1-1c SEASONING**（新321列テンプレ、唐辛子ノード、行7保持・行8開始、PT別出力） | **Git**: 対象差分をrevert。HPCマップは維持 |
| 2026-07-30 | `AmazonApprovalExport.js`／`コード.js`／D_ENTRY docs | **同一レ点行を新規＋相乗りへ同時出品**（N列分割を廃止。識別子は子SKU／Amazon相乗りSKUで分離） | **Git**: revert。Property false |
| 2026-07-30 | D_ENTRY HUMAN_RUN／PHASE／HANDOVER／承認包 | **D相乗り自己発 dry_run／prod 実機合格記録**（`…48as12`／B084RJSH7W） | **Git**: revert docs。Propertyは作業後false |
| 2026-07-30 | `コード.js`／`AmazonApprovalExport.js`／`AmazonSpapiPut.js`／HUMAN_RUN | **D相乗り修正**（Dで自己発/FBA選択、X非依存、ASIN済みSKUそのまま、在庫>0でも送信0） | **Git**: 対象差分をrevert。Propertyをfalseへ |
| 2026-07-30 | `コード.js`／`AmazonApprovalExport.js`／`AmazonSpapiPut.js`／レ点本線docs | **Dレ点本線＋Amazon相乗りSKU実装**（新規／相乗り同時、N列ASINのみ、NF列VALID後保存、prod再利用） | **Git**: 対象差分をrevert。`APPROVAL_AMAZON_LV4_ENABLED=false`／`APPROVAL_AMAZON_SPAPI_PUT_ENABLED=false`。承認①メニューへ戻す |
| 2026-07-30 | レ点本線承認包／憲章／承認マトリクス／Lv4正本／商品マスタ要件／PHASE／HANDOVER | **3者多数決反映**（人間レ点＝当面承認①相当、A/M2専用Amazon相乗りSKU、フル＋相乗りprod許可）。コードなし | **Git**: revert docs。実装前のためProperty／GAS影響なし |
| 2026-07-29 | `コード.js` Dラジオ／offer facade／D_ENTRY HUMAN_RUN／PHASE | **A（D入口版）実装**（既存相乗りを D から呼出） | **Git**: revert。要 clasp push |
| 2026-07-29 | `LV4_SPAPI_D_ENTRY_APPROVAL` §6／PHASE／HANDOVER | **A（D入口版）社長承認**（実装可・コードなし） | **Git**: revert docs |
| 2026-07-29 | `LV4_SPAPI_D_ENTRY_APPROVAL`（新規）／PHASE／HANDOVER | **A（D入口版）承認起草**（Dラジオで新規／既存相乗り・コードなし） | **Git**: revert docs |
| 2026-07-29 | GAS PUT HUMAN_RUN／第2段承認／PHASE／HANDOVER | **v1.4 第2段 実機合格記録**（21-⑫ VALID／21-⑬ ACCEPTED） | **Git**: revert docs |
| 2026-07-29 | `AmazonSpapiPut.js`／`コード.js` 21-⑫⑬／GAS PUT docs | **v1.4 第2段実装**（承認①済→GAS直PUT） | **Git**: revert。Property false。clasp push 済 |
| 2026-07-29 | `LV4_SPAPI_GAS_PUT_STAGE2_APPROVAL` §6／GAS PUT承認・HUMAN_RUN／PHASE／HANDOVER | **v1.4 第2段 社長承認**（実装可・コードなし） | **Git**: revert docs |
| 2026-07-29 | `LV4_SPAPI_GAS_PUT_STAGE2_APPROVAL`（新規）／GAS PUT承認／PHASE | **v1.4 第2段 承認起草**（承認①済→GAS直PUT・コードなし） | **Git**: revert docs |
| 2026-07-29 | `AmazonSpapiPut.js`／GAS PUT docs／PHASE | **v1.4 API実機合格記録＋ENDPOINT `https:\` 正規化** | **Git**: revert。clasp push 済 |
| 2026-07-29 | `AmazonSpapiPut.js`／`コード.js` 21-⑩⑪／GAS PUT docs | **v1.4 GAS Listings直呼び実装** | **Git**: revert。Property false。要 clasp push |
| 2026-07-29 | `LV4_SPAPI_GAS_PUT_APPROVAL`／GAS PUT HUMAN_RUN／PHASE | **v1.4 承認起草**（コードなし） | **Git**: revert docs |
| 2026-07-29 | APPROVED／BRIDGE／CHECKBOX／CURRENT_PHASE／HANDOVER | **v1.2b実機合格＋親レ点完了扱い** | **Git**: revert docs |
| 2026-07-29 | CHECKBOX／BRIDGE／CURRENT_PHASE／HANDOVER | **v1.2c／v1.3 実機合格記録**（`…48s11` prod） | **Git**: revert docs |
| 2026-07-28 | `AmazonSpapiExport.js`／`コード.js` 21-⑧／CHECKBOX承認 | **v1.2c: 子SKUレ点のみ→Drive CSV（選択行廃止）** | **Git**: revert。Property false。要 clasp push |
| 2026-07-28 | `spapi_fetch_drive_csv.py`／listings_write／21-⑨／DRIVE・APPROVED承認 | **v1.3 Drive取得＋v1.2b承認一括＋v1.2a文字コード** | **Git**: revert。Property false。要 clasp push |
| 2026-07-28 | CURRENT_PHASE／HANDOVER／BRIDGE・WRITE HUMAN_RUN／承認包 | **橋渡し v1.2 実機合格を正本反映**（ride01 SC確認・Property false） | **Git**: revert docs |
| 2026-07-28 | `AmazonSpapiExport.js`／`コード.js` 21-⑧／BRIDGE docs | **スプシ→Drive CSV**（SP-API直呼びなし） | **Git**: revert。Property false。要 clasp push |
| 2026-07-28 | `spapi_listings_write` v1.1／BATCH承認／HUMAN_RUN | **複数行CSV**（max_items・安眠/アルコール相乗り1000/0） | **Git**: revert。items.csv・config.local は追跡しない |
| 2026-07-28 | `tools/spapi_listings_write/*`／WRITE HUMAN_RUN／承認包 | **Listings書込 v1**（1SKU offer／dry_run・prod） | **Git**: revert。config.local は追跡しない |
| 2026-07-28 | SPAPI SMOKE／CURRENT_PHASE／書込承認包 | **読取スモーク合格反映**＋書込は承認待ち docs | **Git**: revert docs |
| 2026-07-28 | `tools/spapi_smoke/*`／D_MENU_SPAPI_SMOKE_HUMAN_RUN | **SP-API読取スモーク**（LWA＋Catalog）。秘密は local | **Git**: revert。config.local は追跡しない |
| 2026-07-28 | `tools/m2_offer_packaged/m2_listing_loader_fill.py` 等／M2 docs | **公式 ListingLoader 自動埋め**（人手DL＋fill） | **Git**: revert 当該ファイル |
| 2026-07-28 | M2／CURRENT_PHASE／HANDOVER／GAP docs | **M2実機合格＋SP-API認証完了を正本反映**。公式LoaderがSC正 | **Git**: revert docs |
| 2026-07-27 | `コード.js`（E-3）／`ApprovalQueue.js`／E HUMAN_RUN | **E-3完了時**: 承認Web URL付きダイアログ（`approvalQueueBuildWebUrl_`） | **Git**: revert。要 clasp push |
| 2026-07-27 | `コード.js`（メニュー8）／AI_ADOPT docs | **v1.10**: 共有4列の類似横断dedupe（品/用・包含） | **Git**: revert。要 clasp push |
| 2026-07-27 | `コード.js`（`syncAiDataToMaster`）／商品マスタ作業エリア要件 | **キャッチ1行化**: 楽天／Yahooキャッチ作業エリアへは先頭1案のみ（全文転記禁止） | **Git**: revert。要 clasp push |
| 2026-07-27 | `コード.js`（メニュー8）／AI_ADOPT docs | **v1.9**: 区切り＝半角スペースのみ／メインKW1語／最終名は式の後ろ列・後ろワードから厳格trim | **Git**: revert。要 clasp push |
| 2026-07-27 | `AmazonApprovalExport.js`／`tools/m2_offer_packaged`／M2 docs | **M2 v1**: 案L CSV・競合ASIN・manual_ok。試験=発汗 | **Git**: revert。Property false／BRAND_GATE 削除 |
| 2026-07-27 | `docs/org/LV4_M2_*`／`D_MENU_M2_HUMAN_RUN`／CURRENT_PHASE／Facade U6 | **M2キックオフ**: ギャップ＋承認下書き（コードなし） | **Git**: revert docs |
| 2026-07-27 | `c1_packaged.py`／`config.example.json`／C1 HUMAN_RUN | **C1-clean**: `{subBatchId}`＋`--sub-batch`／`relax`固定廃止 | **Git**: revert `c1_packaged.py`。local は gitignore |
| 2026-07-27 | `tools/c1_hpc_packaged/c1_fetch_inputs.py`／要件／HUMAN_RUN | **C1入力案A**: Drive GENERATED＋マスタCSV自動取得（OAuth読取） | **Git**: revert。secrets は追跡しない |
| 2026-07-27 | `AmazonDriveImageExport.js`／`AmazonApprovalExport.js`／`コード.js`／E・C1 HUMAN_RUN | **親MAIN URL**: U4子→親コピー＋Lv4 Build子フォールバック | **Git**: revert。要 clasp push |
| 2026-07-26 | `コード.js`／`AmazonDriveImageExport.js`／`AmazonApprovalExport.js`／E HUMAN_RUN | **Eコース silent**: 成功時OKなし・エラー時のみ停止。U4/Lv4 silent | **Git**: revert。要 clasp push |
| 2026-07-26 | `コード.js`／`D_MENU_E_AMAZON_COURSE_HUMAN_RUN` | **メニューE**: Amazonコース（E-0〜E-5）薄いファサード | **Git**: revert。要 clasp push |
| 2026-07-26 | `Yahoo.js`／`コード.js`／STAGE | **ブランド正本=getShopBrandList**（検証＋メーカー名API取得） | **Git**: revert。要 clasp push |
| 2026-07-26 | `Yahoo.js`（SHP読取）／`コード.js`／YAHOO_CATEGORY_BRAND_STAGE | **正本=getShopCategoryList**: 書込前検証・無効ID却下／名前解決 | **Git**: revert。要 clasp push（Yahoo.js含む） |
| 2026-07-26 | `コード.js`（Yahooフォールバック）／YAHOO_CATEGORY_BRAND_STAGE | **競合ゼロ時**: AI推奨列→DriveマスタCSV選定 | **Git**: revert。要 clasp push |
| 2026-07-26 | `コード.js`（メニュー8）／YAHOO_CATEGORY_BRAND_STAGE／AI_ADOPT docs | **v1.8**: Yahooカテゴリ／ブランド都度API（売れ筋＋自社最安）をメニュー8へ | **Property**: `AMAZON_AI_ADOPT_YAHOO_CATEGORY_BRAND_ENABLED=false`。**Git**: revert。要 clasp push |
| 2026-07-26 | `YAHOO_CATEGORY_BRAND_STAGE`／承認包／AI_ADOPT | **Yahooカテゴリ選定改定**: 売れ筋＋自社最安優先（docsのみ） | **Git**: revert |
| 2026-07-26 | `YAHOO_CATEGORY_BRAND_STAGE`／承認包／HUMAN_RUN／AI_ADOPT docs／CURRENT_PHASE | **Yahooカテゴリ／ブランド Stage 要件起草**（都度API・`:`階層名・実装未） | **Git**: revert（コード差分なし） |
| 2026-07-26 | `コード.js`（メニュー8）／RAKUTEN_NAV_GENRE_STAGE3／AI_ADOPT docs | **Stage3**: 楽天ジャンル都度API（Ichiba投票→Nav）をメニュー8へ。AI推奨不使用 | **Property**: `AMAZON_AI_ADOPT_RAKUTEN_GENRE_ENABLED=false`。**Git**: revert。要 clasp push |
| 2026-07-26 | `コード.js`（メニュー8）／AI_ADOPT docs | **v1.6**: バリエーションテーマ選定（HPCサイズ／食品パッケージサイズ＋Amazon裏マップ）をメニュー8へ | **Git**: revert。要 clasp push |
| 2026-07-26 | `コード.js`（メニュー8）／AI_ADOPT docs | **v1.5**: 下限=上限−5・半角0.5・容量1語・特徴/用途不適なら空 | **Git**: revert。要 clasp push |
| 2026-07-26 | `コード.js`（メニュー8）／AI_ADOPT 要件・HUMAN_RUN | **v1.4**: 楽天・Yahoo横断dedupe＋最終名上限（75/120/75）。検索KWから弱語削除 | **Git**: revert。要 clasp push |
| 2026-07-26 | `コード.js`（メニュー8）／AI_ADOPT 要件・HUMAN_RUN | **v1.3**: 最終商品名amazon式の左優先・完全一致横断dedupe（部分一致なし） | **Git**: revert。要 clasp push |
| 2026-07-26 | `コード.js`（メニュー8）／AI_ADOPT 要件・HUMAN_RUN | **v1.2**: 再生成停止。KW9選択＋dedupeのみ | **Git**: revert。要 clasp push |
| 2026-07-26 | `コード.js`（メニュー8）／AI_ADOPT 要件・HUMAN_RUN | **緊急**: syncAiDataToMaster をメニュー8から除外（レ点外破壊防止） | **Git**: revert。要 clasp push |
| 2026-07-26 | `コード.js`（メニュー7.5・AI一括採用）／要件・HUMAN_RUN・承認包 | Amazon AI生成→空欄のみ採用→要確認。トグル既定false | **Property** OFF。**Git**: revert |
| 2026-07-26 | `コード.js`（MASTER_TO_WORKAREA／AI_COLUMN_MAP）／商品マスタ要件／RESEARCH／C1 HUMAN_RUN | 市場価格調査→定価転記を廃止。定価＝数値／式専用 | **Git**: revert。**clasp push** 反映後は再pushで戻す |
| 2026-07-26 | `tools/c1_hpc_packaged/**`（C1-1b）／HUMAN_RUN／要件 | マスタCSV併読・必須列。Propertyなし | **Git**: revert |
| 2026-07-26 | `D_MENU_C1_MASTER_HPC_COLUMN_MAP`／C1要件・HUMAN_RUN・CURRENT_PHASE | マスタ→HPC付け合わせ下書き。コードなし | **Git**: revert |
| 2026-07-26 | `tools/c1_hpc_packaged/**`／C1 HUMAN_RUN／承認包／`.claspignore` | C1実装承認＋ローカルPACKAGEDツール | **Git**: revert。成果xlsmはDownloads削除 |
| 2026-07-26 | `D_MENU_C1_THREE_REVIEW_MAJORITY`／C1要件／`LV4_C1_IMPLEMENTATION_APPROVAL` §6／CURRENT_PHASE 等 | C1三点＋社長決定反映。コードなし。次＝実装承認 | **Git**: revert |
| 2026-07-26 | `D_MENU_C1_…REQUIREMENTS`／`LV4_C1_IMPLEMENTATION_APPROVAL`／POC§5／CURRENT_PHASE 等 | C1方針ロック＋要件起草。コードなし。次＝3者 | **Git**: revert |
| 2026-07-26 | U4 HUMAN_RUN／CURRENT_PHASE／AGENT_HANDOVER／CHANGE_LEDGER | **U4実機合格**記録（21-⑦・マスタURL） | **Git**: revert |
| 2026-07-26 | `AmazonDriveImageExport.js`／`AmazonApprovalExport.js`／`コード.js`／U4 HUMAN_RUN 等 | **U4 v1**: 21-⑦ R2→マスタURL＋GENERATED優先 | Property OFF。**Git**: revert |
| 2026-07-26 | `D_MENU_U4_…REQUIREMENTS`／`LV4_U4_IMPLEMENTATION_APPROVAL`／CURRENT_PHASE 等 | U4要件＋承認パッケージ起草。コードなし | **Git**: revert |
| 2026-07-26 | LV4_T2_HUMAN_RUN／CURRENT_PHASE／AGENT_HANDOVER | T2再検証合格（80s10・URL単独・18320なし）記録 | **Git**: revert |
| 2026-07-26 | `AmazonImageMatrixExport.js`／U2 HUMAN_RUN | ④成功MAIN/PTのみ `07/アップロード済み画像` へ退避 | Property `AMAZON_IMAGE_CANDIDATE_ARCHIVE_ENABLED=false` または git revert |
| 2026-07-25 | U2 HUMAN_RUN／CURRENT_PHASE／U2要件／AGENT_HANDOVER／CHANGE_LEDGER | **U2実機合格**記録（②③④・冪等D）。コード差分は同コミットの実装分 | **Git**: revert |
| 2026-07-25 | `コード.js`（`generateAiImageMatrix` 選定）／U2 HUMAN_RUN／CURRENT_PHASE | C: 子レ点優先（親のみレ点は全子のまま）。候補=`07` Property済 | **Git**: revert 該当差分 |
| 2026-07-25 | `AmazonImageMatrixExport.js`（新規）／`コード.js`（C-Amazonメニュー・フック）／U2 HUMAN_RUN／CURRENT_PHASE | U2 v1: sheet MAIN/PT・マスタ永続・02コピー。トグル既定false | **Property** OFF。**Git**: revert＋新規js削除 |
| 2026-07-25 | `D_MENU_U2_THREE_REVIEW_MAJORITY.md`（新規）／U2要件／POC／D要件／CURRENT_PHASE 等 | 三点採用＋社長回答反映。コードなし | **Git**: `git revert` |
| 2026-07-25 | `D_MENU_U2_…REQUIREMENTS.md`／D要件§7／CURRENT_PHASE | U2方針: 案α本線・MAIN=sheet紐付け・02=出口・εバックログ。コードなし | **Git**: `git revert` |
| 2026-07-25 | U3実機合格記録／`D_MENU_U2_…REQUIREMENTS.md`（新規）／CURRENT_PHASE 等 | U3クローズ＋U2起草。コード追加なし（本行） | **Git**: `git revert` |
| 2026-07-25 | `コード.js`（D amazon／full_amazon・`runBatchExportAmazonFacade`）／`AmazonApprovalExport.js`（21-①戻り値）／`D_MENU_U3_HUMAN_RUN.md`／要件・CURRENT_PHASE／AGENT_HANDOVER | U3 v1: 薄いファサード。トリガーにAmazon非搭載 | **Git**: `git revert`。Property OFF |
| 2026-07-24 | `D_MENU_AMAZON_FACADE_REQUIREMENTS.md`（新規）／CURRENT_PHASE／AGENT_HANDOVER／CHANGE_LEDGER／LV4・POCリンク | D本線Amazonファサード要件U0。T3保留。コードなし | **Git**: `git revert` |
| 2026-07-24 | CURRENT_PHASE／AGENT_HANDOVER／CHANGE_LEDGER（T2 PoC成功記録） | runId `R2T2_20260724_221107_7f9cf7`・URL画像表示。コード追加なし | **Property** `AMAZON_DRIVE_R2_UPLOAD_ENABLED=false`。**Git**: `git revert` |
| 2026-07-24 | `AmazonDriveImageExport.js` HMAC Byte[]揃え／clasp push | SigV4で (String,Byte[]) 例外を修正。T2のみ | **Git**: `git revert`／再push |
| 2026-07-24 | CURRENT_PHASE §0（clasp push済反映） | T2: `clasp push` 8 files 成功。コード差分なし（反映のみ） | 戻し不要（GASは前回リビジョンへ） |
| 2026-07-24 | CURRENT_PHASE §0 帰宅引き継ぎ／AGENT_HANDOVER | 自宅PC続行用に本日分を記録。コード追加なし（T2は既存） | **Git**: `git revert` |
| 2026-07-24 | `AmazonDriveImageExport.js`（新規）／`コード.js`（21-⑥）／LV4_T2_HUMAN_RUN／POC／CURRENT_PHASE／AGENT_HANDOVER／CHANGE_LEDGER | T2: Drive→R2 MAIN1枚PoC。トグル既定false | **Property** `AMAZON_DRIVE_R2_UPLOAD_ENABLED=false`。**Git**: `git revert`／新規js削除＋メニュー削除 |
| 2026-07-24 | `RAKUTEN_NAV_STAGE1_HUMAN_RUN.md`／AI_ROUTING §5.1静的／`LV4_T2_IMPLEMENTATION_APPROVAL.md`／CURRENT_PHASE／AGENT_HANDOVER | Nav・AI静的＋T2承認書。GAS未実装 | **Git**: `git revert` |
| 2026-07-24 | CURRENT_PHASE／AGENT_HANDOVER／Drive `05` success JSON／FOOD_在庫確認_STATUS | FOOD v5成功（6バリエーション）。コードなし | **Git**: `git revert`／Drive success JSON削除可 |
| 2026-07-24 | Drive `03/…_v5_corrective.xlsm`・`04/…_MISSING5_MAIN_for_SC.zip`／CURRENT_PHASE／AGENT_HANDOVER／CHANGE_LEDGER | FOOD欠け5子corrective（親部分更新・30s175除外・袋数タイトル）。GASなし | **Drive**: v5ファイル削除。**Git**: `git revert` |
| 2026-07-24 | CURRENT_PHASE §0／LV4_R2_IMAGE_PIPELINE_POC（§2.1二モード）／AGENT_HANDOVER／CHANGE_LEDGER | 外出リモート引き継ぎ＋サブREUSE/ONLY方針。コードなし | **Git**: `git revert` |
| 2026-07-24 | Drive `04` 配下01〜06作成／POC §2 ID表／CURRENT_PHASE／`DRIVE_04_FOLDER_IDS.md`(Downloads) | フォルダ実体＋Folder ID控え。GAS・Property未設定 | **Drive**: 空フォルダ削除可。**Git**: `git revert` |
| 2026-07-24 | `LV4_R2_IMAGE_PIPELINE_POC.md`／CURRENT_PHASE／AGENT_HANDOVER／CHANGE_LEDGER／LV4要件 | Drive起点GAS案を正に。xlsm自動は提案のみ。コードなし | **Git**: `git revert` |
| 2026-07-24 | `docs/org/LV4_R2_IMAGE_PIPELINE_POC.md`／CURRENT_PHASE／AGENT_HANDOVER／CHANGE_LEDGER／LV4要件リンク | R2×Amazon画像PoC設計（楽天・Yahoo対象外）。コードなし | **Git**: `git revert`／PoC md削除 |
| 2026-07-23 | CURRENT_PHASE §0／AGENT_HANDOVER／CHANGE_LEDGER | 帰宅引き継ぎ: HPC suburl試験・FOOD再UP禁止。コードなし | **Git**: `git revert` |
| 2026-07-23 | LV4要件 §11.0・§11表・§11.5.3画像／CURRENT_PHASE／AGENT_HANDOVER／CHANGE_LEDGER | HPC §11-1〜8・10＋U5クローズ。画像ZIP優先。FOOD／M2／21-⑤は別ゲート。コードなし | **Git**: `git revert` |
| 2026-07-23 | CURRENT_PHASE §0／AGENT_HANDOVER／CHANGE_LEDGER／accepted_values_db(Downloads) | 外出リモート引き継ぎ: HPC 21-③済・FOOD待機。コードなし | **Git**: `git revert`／DBはDownloads削除可 |
| 2026-07-22 | LV4要件 §11.5.4・§11.6／CURRENT_PHASE／`accepted_values_db/` | FOODテンプレ確認＋成功値DB方針。種子JSONはDownloads配下 | **Git**: `git revert`／DBフォルダ削除可 |
| 2026-07-22 | CURRENT_PHASE §0／AGENT_HANDOVER／CHANGE_LEDGER | 外出引き継ぎ: corrective画像18320・中期R2方針。コードなし | **Git**: `git revert` |
| 2026-07-22 | LV4要件 §6.1.2・§10・§11.5.3／CURRENT_PHASE / AGENT_HANDOVER / CHANGE_LEDGER | 修正登録（SKU維持）・21-⑤・実データ7行目起点。GASメニュー実装は後続。ローカル corrective_link.xlsm | **Git**: `git revert`／PACKAGEDファイル削除 |
| 2026-07-21 | LV4要件 §11.5・§4／CURRENT_PHASE / AGENT_HANDOVER | HPC列対応ドラフト＋マスタ必須候補方針＋成功値辞書化メモ（後続）。コードなし | **Git**: `git revert` |
| 2026-07-21 | LV4要件 §5・§6.1.1・§14・§17 Q5／CURRENT_PHASE / CHANGE_LEDGER | subBatchId方針を新番号に統一＋push手順。コードは既存実装どおり | **Git**: `git revert` |
| 2026-07-20 | AmazonApprovalExport.js | 再レビュー採用: DRY_RUN／冪等汚染／subBatchId単調／ブランド厳密／archive失敗停止 | **Git**: `git revert` |
| 2026-07-20 | AmazonApprovalExport.js | 実装レビュー採用修正（レジューム・SKIPPED記録・GTIN・冪等・追記ログ）。戻し: git revert | **Git**: `git revert` |
| 2026-07-20 | AmazonApprovalExport.js（新規）／ApprovalQueue.js／コード.js（メニュー21）／LV4要件・CURRENT_PHASE / AGENT_HANDOVER | Lv4実装。ENABLED=false既定。戻し: Property OFF＋git revert＋新規js削除 | **Git**: `git revert` / Property `APPROVAL_AMAZON_LV4_ENABLED=false` |
| 2026-07-20 | LV4要件 §3.2・§2・§13・§17／AMAZON_REQUIREMENTS §3／多数決 §4.2／CURRENT_PHASE / AGENT_HANDOVER | Q15/Q16＋列メモ注記。コードなし | **Git**: `git revert` |
| 2026-07-20 | LV4要件 §3.4–3.6・§6・§17／多数決／CURRENT_PHASE / AGENT_HANDOVER | Q11–Q14反映。方針Q&A一通り閉じ。コードなし | **Git**: `git revert` |
| 2026-07-20 | LV4要件 §3.1–3.2・§3.5・§17／多数決 §4.1／CURRENT_PHASE / AGENT_HANDOVER | Q7–Q10b: 3モール同一承認・TRACK=B強制・SKU/メーカー品番。コードなし | **Git**: `git revert` |
| 2026-07-20 | LV4要件 §1.4・§17／多数決メモ §4.1／CURRENT_PHASE / AMAZON_REQUIREMENTS | 社長Q&A: D-1 PACKAGED・JAN出力規則・バリエーションのみ・再生成＋ログ。コードなし | **Git**: `git revert` |
| 2026-07-20 | `docs/org/LV4_THREE_REVIEW_MAJORITY.md`・`LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md`・CURRENT_PHASE / AGENT_HANDOVER / LEVELLED_PLAN / AMAZON_REQUIREMENTS | Lv4三点レビュー多数決保存＋社長確定反映（在庫書込禁止・DONE分離・親SKU・GTINゲート）。コードなし | **Git**: `git revert` |
| 2026-07-20 | `docs/org/LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md`・CURRENT_PHASE / AGENT_HANDOVER / LEVELLED_PLAN / AMAZON_REQUIREMENTS | Lv4要件ドラフト（M1=B新規・M2=A既存・手動UP・限定ツール）。コードなし | **Git**: `git revert` または当該ファイル削除 |
| 2026-07-20 | CURRENT_PHASE / LV3要件 §8.1・§8.2 / AGENT_HANDOVER / LEVELLED_PLAN | Lv3人間検収完了を共有。フォーカスを Lv4 Amazon へ。在庫列二重ヘッダー運用メモ | **Git**: `git revert` |
| 2026-07-20 | `YahooApprovalExport.js`（resolveログ） | Lv3スキップ切り分け: sheet名・在庫生値・ORPHAN/IN_STOCK を Logger 出力 | **Git**: `git revert`／当該 Logger 行を削除 |
| 2026-07-20 | `YahooApprovalExport.js`（新規）・`ApprovalQueue.js`（Yahoo読取）・`コード.js`（メニュー20）・Lv3 docs | Lv3: 承認①済Yahoo子を案A＋主25分／副ユニーク50で `runYahooExport` 呼出。Yahoo本体非改変 | **Property**: `APPROVAL_YAHOO_LV3_ENABLED=false`。**Git**: `git revert`／新規js削除＋メニュー20削除 |
| 2026-07-20 | `docs/org/LV3_YAHOO_ORCHESTRATION_REQUIREMENTS.md`・CURRENT_PHASE / AGENT_HANDOVER / LEVELLED_PLAN | Lv3要件ドラフト（runYahooExport呼出・主25分／副ユニーク50運用揃え）。コードなし | **Git**: `git revert` または当該ファイル削除 |
| 2026-07-20 | CURRENT_PHASE / LV2要件 §8.1・§10.1 / AGENT_HANDOVER / LEVELLED_PLAN | Lv2人間検収完了を共有。フォーカスを Lv3 Yahoo へ。目視手順追記 | **Git**: `git revert` |
| 2026-07-20 | `RakutenApprovalExport.js`（案Aレ点）・Lv2 docs | 案Aを「承認済み親＋紐づく子も一時レ点」に変更。CSV本体非改変。バリエーション親のシングルSKU誤認を解消 | **Git**: `git revert`。関数 `rakutenApprovalLv2ApplyPlanACheckboxes_` を親だけONに戻す |
| 2026-07-21 | `AmazonApprovalExport.js`（価格解決）・LV4 docs | 親`販売価格amazon`空→承認済み子へフォールバック。`[Lv4Price]`ログ | **Git**: `git revert` |
| 2026-07-21 | `AmazonApprovalExport.js`（カテゴリ解決）・LV4 docs | T列`カテゴリー`フォールバック＋見出しゆれ対応。`[Lv4Cat]`ログ | **Git**: `git revert`／当該関数を旧2行lookupに戻す |
| 2026-07-17 | CURRENT_PHASE / LV1要件 §8.1 / AGENT_HANDOVER / LEVELLED_PLAN / LV2前提 | Lv1人間検収完了をプロジェクト全体へ共有。フォーカスを Lv2 へ | **Git**: `git revert` |
| 2026-07-17 | `ApprovalQueue.js`・`Yahoo.js`(doGet分岐のみ)・`コード.js`(メニュー18)・`appsscript.json` | Lv1承認キュー実装。シート追記＋Web承認。出品API/CSV非呼出。Property既定false | **Property**: `APPROVAL_QUEUE_V1_ENABLED=false`。シート削除可。**Git**: `git revert` |
| 2026-07-17 | `docs/org/LV1_APPROVAL_QUEUE_REQUIREMENTS.md` 等 | Lv1承認キュー要件（シート列・batchId・レ点抽出・Web・EC書込なし） | **Git**: `git revert` または当該ファイル削除 |
| 2026-07-17 | org 多数決・Lvプラン・マトリクス・CURRENT_PHASE 等 | Lv0最終承認・モール順楽天先・レ点/スキップ/U1手動上書きの文書化（コードなし） | **Git**: `git revert` |
| 2026-07-17 | `docs/org/LEVELLED_IMPLEMENTATION_PLAN.md`、CURRENT_PHASE / AGENT_HANDOVER 追記 | AI組織の Lv0〜5＋並行P の実装順叩き台（コードなし） | **Git**: `git revert` または当該ファイル削除 |
| 2026-07-17 | `コード.js`（`menuDiagnoseRakutenNavigationGenreStage1Write` 等）、`docs/RAKUTEN_NAV_GENRE_STAGE1.md` | Stage1: 専用シート `▼診断(楽天ジャンルNav)` への追記のみ（マスタ/CSV非書込・楽天CSV非改変）。Property 既定 false | **Property**: `RAKUTEN_NAV_GENRE_STAGE1_WRITE_ENABLED=false`。シート削除可。**Git**: `git revert`。メニュー17-⑥/99-⑩と Stage1 関数を削除 |
| 2026-07-17 | `docs/RAKUTEN_NAV_GENRE_STAGE1.md`（要件） | Stage0完了後の Stage1：テストシート追記隔離の要件正本化 | **Git**: `git revert` または当該ファイル削除 |
| 2026-07-16 | `parseRakutenNavigationGenreResponse_` | NavigationAPI 2.0 の `nameJa` / `nameJaPath` を解釈して namePath を出す | **Git**: `git revert` |
| 2026-07-16 | `diagnoseRakutenNavigationGenreGetOne_` の URL | NavigationAPI 2.0 を `/es/2.0/navigation/genres/{id}` に修正（GF0002 404 対策） | **Git**: `git revert` または URL 配列を旧 query 形式に戻す |
| 2026-07-16 | `コード.js`（`menuDiagnoseRakutenNavigationGenreGet` 等）、`docs/RAKUTEN_NAV_GENRE_DIAG.md` | 楽天ジャンルID NavigationAPI 読取疎通診断（マスタ/CSV書込なし・CSV経路非改変） | **Property**: `RAKUTEN_NAV_GENRE_DIAG_ENABLED=false`。**Git**: `git revert`。診断関数・メニュー17-⑤/99-⑨を削除 |
| 2026-07-15 | `docs/org/PHASE0_THREE_REVIEW_MAJORITY.md`・`THREE_REVIEW_RUNBOOK.md`、憲章・マトリクス更新、`.cursor/rules/three-review-runbook.mdc`、CURRENT_PHASE / AGENT_HANDOVER | 3者多数決の記録・採用項反映・親1＋並列サブ3運用の正本化（実装コードなし） | **Git**: `git revert`。ルール削除は `three-review-runbook.mdc` を削除 |
| 2026-07-15 | `docs/org/AI_ORG_CHARTER.md`・`docs/org/AI_APPROVAL_MATRIX.md`、および `CURRENT_PHASE.md` / `AGENT_HANDOVER.md` の参照追記 | 一人社長＋AI部門の組織憲章・承認マトリクスを docs 正本化（実装なし。Phase0→3者検証前提） | **ファイル削除**: `docs/org/` の当該2ファイルを削除し、CURRENT_PHASE / AGENT_HANDOVER / 本台帳の該当行を戻す。**Git**: `git revert` |
| 2026-07-11 | `Yahoo.js`（`findYahooMasterHeaderRowIndex_`・`getYahooMasterHeaderContext_`・`_loadMasterData`・`updateMasterYahooId`・`updateMasterDeleteFlag`・`showDeleteSelectionDialog`・`listDeletableItems`） | マスタヘッダー行の `data[7]` 固定を廃止し、出品 Builder と同じ動的検出に統一（`docs/REVIEW_Yahoo_js.md` 指摘#1） | **ファイル**: `Yahoo.js.bak_before_header_unify_20260711` または `Yahoo.js.local_backup` を `Yahoo.js` に上書きコピー。**Git**: 本変更コミット後は `git revert` |
| 2026-05-09 | `コード.js`（`keepaProductPrimaryImageToken_`・`keepaAmazonMediaUrlFromImageToken_`・`getKeepaProductImageUrl`・`getKeepaProductImageUrlsAll`・`getKeepaProductImageUrlsMaxForCompetitorTest_`・`buildKeepaTableRow`） | Keepa API が返す **images 配列（l / m）** を解釈する。従来の image / imagesCSV のみでは現行レスポンスで画像URLが空になる問題を修正 | **Git**: `git revert`。**GAS**: 同ブロックを旧実装に戻す |
| 2026-05-06 | `コード.js`・`.claspignore` | R-Cabinet 診断ログ（パイプラインタグ・送信 bytes・JPEG 比・cap 付きサムネ）、Capacity 時に長辺 cap を段階的に下げる再試行、**`clasp push` で本番反映**。`.claspignore` に `**/*.html` 等で誤 push 防止 | **Git**: `git revert`。**GAS**: script.google.com で誤追加の `.html` を削除（以前の誤 push 分）。再試行を単発に戻す場合は `executeRenameAndUploadFromMatrixProgrammatic_` 内の `capList` ループを旧 1 回 `tryFetch` に戻す |
| 2026-05-06 | `RAKUTEN_CABINET_MAX_IMAGE_BYTES`・`convertBlobToJpegForRakuten_`・`prepareBlobForRakutenCabinetUpload_`・Drive サムネ系（`コード.js`） | R-Cabinet `Capacity` 回避のため目安を **1,900,000 バイト**に締め、サムネ／オリジナルとも **送信直前に JPEG 化**し **JPEG 後バイト**で上限判定する一本化 | **Git**: `git revert`。**手動**: 定数を `2*1024*1024` に戻し `convertBlobToJpegForRakuten_` を削除、サムネ戻り値の `convertBlobToJpegForRakuten_` 呼び出しと `prepare` の二段ロジックを旧実装に戻す（判断は `docs/BATCH_EXPORT_IMAGE_GATE_REQUIREMENTS.md`） |
| 2026-05-06 | `generateRakutenCSV`（`コード.js`）・`getRakutenSpecHtmlCell_` | PC/スマホ説明文の「商品情報」表を、▼マスタ列なしの明示フォールバック（メーカー名・マスタシリーズ／特記すべき原材料・原材料／括弧付き賞味・保存＋従来名）で取得 | **Git**: `git revert`。スペック4行を旧 `getMVal(ブランド/シリーズ/原材料/賞味/保存)` に戻し `getRakutenSpecHtmlCell_` 関数削除 |
| 2026-05-02 | `generateRakutenCSV`（`コード.js`）・ヘルパー `isMasterCellEmptyForRakutenAttr_` / `isRakutenAttrValueEmptyForCleaning_` | マルチSKUで親の `楽天セット数`（0/空）が `virtualRow`・`parentDataValues` 経由で子の属性値を潰し、かつ `attr_value` の `!val` で項目・単位だけ消える問題を防止 | **Git**: 該当コミットを `git revert`。**手動**: `virtualRow` のコピー・SKU属性継承・`attr_value` クリーニング3点を元に戻す（判断経緯は `docs/RAKUTEN_CSV_ATTR_AND_SETCOUNT_REQUIREMENTS.md`） |
| 2026-03-22 | `runProductNameProposalsForRows`（`コード.js`） | OpenAI 失敗時もマスタ商品名があればバリエーション（単位・内容量）を実行。Script Property `PRODUCT_NAME_PROPOSALS_CONTINUE_VARIATION_ON_OPENAI_FAIL`（既定 `true`）で旧挙動に戻せる | **Git**: 本変更のコミットを `git revert <commit>`。**運用**: Script Properties に `PRODUCT_NAME_PROPOSALS_CONTINUE_VARIATION_ON_OPENAI_FAIL` = `false` で旧挙動（OpenAI 失敗行はバリエーションもスキップ） |
| 2026-03-22 | `inferVariationFromAsinCircleForJan_`・`pickPerSetContentFromCircleTitles_` 等（`コード.js`） | 同一 JAN ブロック内の **全 ◎ 行**から 1セットあたり内容量を総合判断（g/ml 優先）。`CIRCLE_COMBINED_PER_SET_CONTENT`（既定 `true`）で無効化すると旧「先頭ヒット打切り」 | **Git**: `git revert <commit>`。**運用**: `CIRCLE_COMBINED_PER_SET_CONTENT` = `false` |

**コミットハッシュ**: ローカルで `git log -1 --oneline` を実行したうえで、上表に追記すること。
