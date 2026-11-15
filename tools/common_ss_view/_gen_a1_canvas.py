# -*- coding: utf-8 -*-
"""Generate a1-efficiency.canvas.tsx with correct cursor/canvas Table API."""
from __future__ import annotations

import json
from pathlib import Path

DATA = json.loads(Path("tools/common_ss_view/out/_a1_canvas_data.json").read_text(encoding="utf-8"))
OUT = Path(
    r"C:\Users\takuy\.cursor\projects\c-Users-takuy-Desktop-gas-project\canvases\a1-efficiency.canvas.tsx"
)
data_lit = json.dumps(DATA, ensure_ascii=False, indent=2)

tsx = """import {
  BarChart,
  Callout,
  Card,
  CardBody,
  CardHeader,
  Divider,
  Grid,
  H1,
  H2,
  LineChart,
  Row,
  Spacer,
  Stack,
  Stat,
  Table,
  Text,
} from "cursor/canvas";

const DATA = """ + data_lit + """ as const;

export default function A1EfficiencyCanvas() {
  const s = DATA.stats;
  const tone =
    DATA.overall === "合格"
      ? "success"
      : DATA.overall === "条件付き合格"
        ? "warning"
        : "danger";

  const rowTone = DATA.rows.map((r) => {
    if (r.ok.includes("合格") || r.ok === "完走") return "success" as const;
    if (r.ok.includes("継続") || r.ok.includes("保留") || r.ok.includes("再試行"))
      return "warning" as const;
    return undefined;
  });

  const tableRows = DATA.rows.map((r) => {
    const diff = r.diff_s;
    const kLeft = r.k_left == null ? "—" : `${r.k_left}/300`;
    const kNeed =
      r.k_need == null
        ? "—"
        : r.k_consumed != null
          ? `${r.k_need} (実${r.k_consumed})`
          : String(r.k_need);
    const kWait =
      r.k_wait == null
        ? "—"
        : r.k_tofull != null
          ? `${r.k_wait}分 / 満タン想定${r.k_tofull}分`
          : `${r.k_wait}分`;
    const cat = r.c_pages == null ? "—" : `${r.c_pages}頁 / 予算50`;
    const catCd = r.c_cd > 0 ? `${r.c_cd}分` : "0分 (CD60未使用)";
    const miss = r.miss > 0 ? ` miss+${r.miss}分` : "";
    const name = r.parent ? `${r.name} / ${r.parent}` : r.name;
    return [
      r.time,
      name,
      r.ok + miss,
      `${r.actual_s}s`,
      `${r.exp_s}s`,
      `${diff >= 0 ? "+" : ""}${diff}s`,
      kLeft,
      kNeed,
      kWait,
      cat,
      catCd,
    ];
  });

  return (
    <Stack gap={20} style={{ padding: 16, maxWidth: 1100 }}>
      <Stack gap={6}>
        <H1>A-1以降 合否・効率</H1>
        <Text tone="secondary" size="small">
          Source: ①ログ · {DATA.meta.rid} · {DATA.meta.at.slice(0, 19)} →{" "}
          {DATA.meta.window_end.slice(0, 19)} · {DATA.meta.events} events
        </Text>
      </Stack>

      <Row gap={12} wrap>
        <Stat label="総合判定" value={DATA.overall} tone={tone} />
        <Stat label="C4合格" value={String(s.c4_ok)} tone="success" />
        <Stat
          label="C4不合格"
          value={String(s.c4_ng)}
          tone={s.c4_ng ? "danger" : "neutral"}
        />
        <Stat label="Catalog列挙" value={String(s.enum_ok)} />
        <Stat
          label="trigger_missed"
          value={`${s.missed_slices}/${s.slices}`}
          tone={s.missed_slices ? "warning" : "success"}
        />
      </Row>

      <Callout tone={tone} title="判定メモ">
        {DATA.reasons.join(" · ")}
      </Callout>

      <Grid columns={4} gap={12}>
        <Card>
          <CardHeader>Keepa残トークン</CardHeader>
          <CardBody>
            <Stat label="最小→最大" value={`${s.keepa_min}→${s.keepa_max}`} />
            <Text size="small" tone="secondary">
              直近 {s.keepa_last} / 想定満タン 300
            </Text>
          </CardBody>
        </Card>
        <Card>
          <CardHeader>Keepa回復</CardHeader>
          <CardBody>
            <Stat label="rate" value={`${s.keepa_rate}/s`} />
            <Text size="small" tone="secondary">
              空→満タン想定 {s.keepa_full_recovery_min}分（maxWait61）
            </Text>
          </CardBody>
        </Card>
        <Card>
          <CardHeader>Catalog回復</CardHeader>
          <CardBody>
            <Stat
              label="429/503"
              value={s.has_429 ? "あり" : "なし"}
              tone={s.has_429 ? "danger" : "success"}
            />
            <Text size="small" tone="secondary">
              CD想定 {s.catalog_cd_min}分 · 本窓は未発動
            </Text>
          </CardBody>
        </Card>
        <Card>
          <CardHeader>効率のボトルネック</CardHeader>
          <CardBody>
            <Text>
              Keepaスライス後の GAS トリガー遅延（trigger_missed
              13〜26分）が主因。Catalog側は回復待ちゼロで高効率。
            </Text>
          </CardBody>
        </Card>
      </Grid>

      <Divider />

      <H2>Keepaトークン推移（実測残 vs 想定満タン300）</H2>
      <Text size="small" tone="secondary">
        想定レート 5/s。縦軸はトークン数。
      </Text>
      <LineChart
        categories={[...DATA.keepa_line.cats]}
        series={[
          { name: "実測 left", data: [...DATA.keepa_line.left], tone: "info" },
          { name: "想定満タン", data: [...DATA.keepa_line.full], tone: "neutral" },
        ]}
        height={220}
        beginAtZero
        yMax={320}
      />

      <H2>実行時間: 実測 vs 想定（分）</H2>
      <Text size="small" tone="secondary">
        Catalog=pages×8s+60s / A-light≈12分 / Keepaスライス作業想定15分 / C4≈100s
      </Text>
      <BarChart
        categories={[...DATA.bar.cats]}
        series={[
          { name: "実測(分)", data: [...DATA.bar.actual], tone: "info" },
          { name: "想定(分)", data: [...DATA.bar.expected], tone: "neutral" },
        ]}
        height={320}
        horizontal
      />

      <H2>各実行の詳細（回復・想定トークン併記）</H2>
      <Table
        headers={[
          "開始",
          "処理 / parent",
          "合否",
          "実測",
          "想定",
          "差",
          "Keepa残 実/満タン",
          "想定消費tok",
          "Keepa回復 実/想定",
          "Catalog頁 実/予算",
          "Catalog回復",
        ]}
        rows={tableRows}
        rowTone={rowTone}
        stickyHeader
        striped
      />

      <Spacer height={8} />
      <Text size="small" tone="secondary">
        coding-rule.mdc / common-ss-view.mdc。Keepa想定消費=want×unit〜6 または
        consumed。Catalogは頁数を予算トークン相当（soft50）として併記。429時CD想定60分。
      </Text>
    </Stack>
  );
}
"""

OUT.write_text(tsx, encoding="utf-8")
print("wrote", OUT, "bytes", OUT.stat().st_size)
