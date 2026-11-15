# -*- coding: utf-8 -*-
"""かつお caption / フリースペース生成。

貼付分担（テスト合意）:
- 商品説明PC: 日本語のみ（CTA・画像・セット）。多言語なし
- フリースペース1: EN / 中 / 韓 / TH / VI 長文
- フリースペース2: PT / Filipino 長文
- スマホFS: 日本語ブロック + EN / 中 / 韓 / TH / VI（残り枠で充実。PT/TLはPC FS2）

外国語対策:
- 中・韓・タイは数値文字参照（NCR）。直書きUTF-8は Yahoo で ? 化
- VI/PT はアクセントなしASCII（〓 化回避）
"""
from pathlib import Path

sets = [
    (2, "ishi-4906283045119-64s18"),
    (3, "ishi-4906283045119-97s12"),
    (4, "ishi-4906283045119-12s196"),
    (5, "ishi-4906283045119-16s120"),
    (6, "ishi-4906283045119-19s144"),
    (8, "ishi-4906283045119-25s192"),
    (10, "ishi-4906283045119-32s140"),
    (12, "ishi-4906283045119-38s188"),
    (15, "ishi-4906283045119-48s160"),
]
BASE_ITEM = "https://store.shopping.yahoo.co.jp/octas"
CTA_SETS = "https://shopping.c.yimg.jp/lib/octas/cta_gourmet_sets.jpg"
CTA_BEST = "https://shopping.c.yimg.jp/lib/octas/cta_gourmet_bestsellers.jpg"
BODY1 = "https://shopping.c.yimg.jp/lib/octas/katsuo_body_t03_sizzle.jpg"
BODY2 = "https://shopping.c.yimg.jp/lib/octas/katsuo_body_t04_aji.jpg"
BODY3 = "https://shopping.c.yimg.jp/lib/octas/katsuo_body_t11_arrange.jpg"
STORE_TOP = "https://store.shopping.yahoo.co.jp/octas/"
PC_W = 600
LIMIT = 10000


def ncr(text: str) -> str:
    out = []
    for ch in text:
        o = ord(ch)
        if ch in ("\n", "\r", "\t") or 32 <= o < 127:
            out.append(ch)
        else:
            out.append(f"&#{o};")
    return "".join(out)


def wrap_pc(body: str) -> str:
    return (
        f'<table width="{PC_W}" border="0" align="center" cellpadding="0" cellspacing="0">\n'
        f'<tr><td align="center">\n{body}\n</td></tr>\n</table>\n'
    )


def sets_table(table_w: str) -> str:
    rows = []
    for i in range(0, len(sets), 3):
        cells = []
        for j in range(3):
            if j > 0:
                cells.append('<td width="2%">&nbsp;</td>')
            if i + j < len(sets):
                n, code = sets[i + j]
                cells.append(
                    f'<td width="32%" bgcolor="#f5efe4" align="center">'
                    f'<a href="{BASE_ITEM}/{code}.html">'
                    f"<b><br>{n}袋セット<br>　</b></a></td>"
                )
            else:
                cells.append('<td width="32%" align="center">&nbsp;</td>')
        rows.append("<tr>\n" + "\n".join(cells) + "\n</tr>")
        if i + 3 < len(sets):
            rows.append('<tr><td colspan="5" height="8"></td></tr>')
    return (
        f'<table width="{table_w}" border="0" cellpadding="8" cellspacing="0">\n'
        + "\n".join(rows)
        + "\n</table>"
    )


def jp_story(*, img_w: str, table_w: str, with_bottom_cta: bool, compact: bool = False) -> str:
    """日本語ストーリー（CTA・画像・セット）。商品情報の短文とは別枠。"""
    bottom = ""
    if with_bottom_cta:
        bottom = f"""
<br>
<a href="#sets"><img src="{CTA_SETS}" width="{img_w}" alt="お得なセットを見る"></a><br>
<br>
<a href="{STORE_TOP}"><img src="{CTA_BEST}" width="{img_w}" alt="当店の売れ筋を見る"></a><br>
"""
    if compact:
        howto = """<b>おすすめの食べ方</b><br>
01　そのまま、おつまみに<br>
02　温かいご飯や卵かけご飯に<br>
03　お茶漬け、サラダ、冷奴に<br>
"""
        lead = """<b>忙しい日のもう一品に。</b><br>
<b>袋を開けて、そのまま食卓へ。</b><br>
<br>
国産かつおを、だし香る醤油仕立てで仕上げた「食べるおだし」。<br>
ふわりとほどける食感と、噛むほどに広がる旨味が特長です。<br>
<br>
ご飯・お茶漬け・サラダ・冷奴・パスタやおつまみに。<br>
"""
    else:
        howto = """<b>おすすめの食べ方</b><br>
01　そのまま、おつまみに<br>
02　温かいご飯や卵かけご飯に<br>
03　お茶漬け、おにぎり、サラダ、冷奴に<br>
04　パスタや豆腐料理のトッピングに<br>
"""
        lead = """<b>忙しい日のもう一品に。</b><br>
<b>袋を開けて、そのまま食卓へ。</b><br>
<br>
国産かつおを、だし香る醤油仕立てで仕上げた「食べるおだし」。<br>
ふわりとほどける食感と、噛むほどに広がる旨味が特長です。<br>
<br>
温かいご飯にのせるだけでもう一品。<br>
お茶漬け・卵かけご飯・おにぎり・サラダ・冷奴・パスタやおつまみにも使えます。<br>
"""
    return f"""<a name="top"></a>
<table width="{table_w}" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center">
      <a href="#sets">
        <img src="{CTA_SETS}" width="{img_w}" alt="お得なセットを見る">
      </a>
    </td>
  </tr>
  <tr>
    <td align="center">
      <br>
      <a href="{STORE_TOP}">
        <img src="{CTA_BEST}" width="{img_w}" alt="当店の売れ筋を見る">
      </a>
    </td>
  </tr>
</table>
<br>
{lead}
<br>
<img src="{BODY1}" width="{img_w}" alt="シズル・食感"><br>
<br>
<img src="{BODY2}" width="{img_w}" alt="味の正しいアピール"><br>
<br>
<img src="{BODY3}" width="{img_w}" alt="アレンジレシピ"><br>
<br>
{howto}
<br>
<a name="sets"></a>
<hr>
<b>【セットコーナー】</b><br>
袋数を選ぶと、各セットの商品ページへ移動します。<br>
<br>
{sets_table(table_w)}
<br>
<hr>
<b>【商品について】</b><br>
名称：食べるおだし（かつお）<br>
内容量：1袋あたり50g<br>
主原料：国産かつお<br>
保存：直射日光・高温多湿を避けて常温<br>
<br>
※原材料・アレルゲン・賞味期限はパッケージ表示が正式です。<br>
※クーポンはカート画面でもご確認ください。<br>
{bottom}"""


def i18n_fs1_long() -> str:
    """PCフリースペース1: EN / 中 / 韓 / TH / VI。"""
    zh_nav = ncr("简体中文")
    ko_nav = ncr("한국어")
    th_nav = ncr("ภาษาไทย")

    zh_title = ncr("简体中文｜开袋即食的日式调味鲣鱼")
    zh_p1 = ncr(
        "本产品使用日本国产鲣鱼，搭配酱油和日式高汤风味调制而成。"
        "口感松软易入口，鲣鱼的鲜味浓郁，适合日常用餐或作为方便的小菜。"
    )
    zh_p2 = ncr(
        "开袋后可直接食用，也可铺在热米饭上。"
        "还适合搭配茶泡饭、饭团、沙拉、冷豆腐、意大利面，或作为佐酒小食。"
    )
    zh_net = ncr("净含量：每袋50克")
    zh_store = ncr("保存方法：请置于常温环境，避免阳光直射、高温及潮湿。开封后请密封保存并尽快食用。")
    zh_imp = ncr("重要提示：完整的配料、过敏原、保质期及制造商信息，请以实际商品包装上的日文标签为准。")

    ko_title = ncr("한국어｜바로 즐기는 일본식 가다랑어 조미식품")
    ko_p1 = ncr(
        "일본산 가다랑어에 간장과 다시 풍미를 더한 제품입니다. "
        "부드러운 식감과 진한 감칠맛으로, 일상 식사나 간단한 밑반찬으로 즐기기 좋습니다."
    )
    ko_p2 = ncr(
        "개봉 후 그대로 드시거나 따뜻한 밥 위에 올려 드세요. "
        "오차즈케, 주먹밥, 샐러드, 냉두부, 파스타, 안주로도 잘 어울립니다."
    )
    ko_net = ncr("내용량: 1봉 50g")
    ko_store = ncr(
        "보관: 실온, 직사광선·고온다습을 피하세요. 개봉 후 밀봉하여 빨리 섭취하세요."
    )
    ko_imp = ncr(
        "중요: 원재료·알레르기·소비기한·제조자 정보는 실제 포장의 일본어 표시가 정식입니다."
    )

    th_title = ncr("ภาษาไทย｜ปลาโอญี่ปุ่นปรุงรส พร้อมรับประทาน")
    th_p1 = ncr(
        "ทำจากปลาโอจากญี่ปุ่น ปรุงด้วยรสโชยุและดาชิ "
        "เนื้อนุ่ม อูมามิเข้ม เหมาะกับมื้อประจำวันหรืองานว่างง่ายๆ"
    )
    th_p2 = ncr(
        "เปิดซองรับประทานได้ทันที หรือโรยบนข้าวร้อน "
        "เหมาะกับโอชาซุเกะ ข้าวปั้น สลัด เต้าหู้เย็น พาสตา หรือของว่างคู่เครื่องดื่ม"
    )
    th_net = ncr("น้ำหนักสุทธิ: 50 กรัมต่อซอง")
    th_store = ncr(
        "การเก็บรักษา: อุณหภูมิห้อง หลีกเลี่ยงแสงแดด ความร้อน และความชื้น "
        "หลังเปิดแล้วให้ปิดสนิทและรับประทานโดยเร็ว"
    )
    th_imp = ncr(
        "ข้อมูลสำคัญ: ส่วนประกอบ สารก่อภูมิแพ้ วันหมดอายุ และผู้ผลิต "
        "ให้ยึดฉลากภาษาญี่ปุ่นบนบรรจุภัณฑ์จริงเป็นหลัก"
    )

    return f"""<a name="international"></a>
<b>【For international customers / 海外のお客様へ】</b><br>
<a href="#lang-en">English</a> ｜
<a href="#lang-zh">{zh_nav}</a> ｜
<a href="#lang-ko">{ko_nav}</a> ｜
<a href="#lang-th">{th_nav}</a> ｜
<a href="#lang-vi">Tieng Viet</a><br>
<br>
<a name="lang-en"></a>
<b>English｜Seasoned Bonito, Ready to Enjoy</b><br>
<br>
This Japanese food product is made with bonito from Japan and finished with a savory soy sauce and dashi-style flavor. Its tender, flaky texture and rich umami make it easy to enjoy as part of an everyday meal.<br>
<br>
Open the package and enjoy it as it is, or serve it over warm rice. It also works well with ochazuke, rice balls, salads, chilled tofu, pasta, or as a savory snack with drinks.<br>
<br>
<b>Net weight:</b> 50 g per bag<br>
<b>Storage:</b> Keep at room temperature away from direct sunlight, high temperatures, and humidity. After opening, seal well and consume as soon as possible.<br>
<br>
<b>Important:</b> Please check the Japanese label on the actual package for the complete ingredients, allergens, best-before date, and manufacturer information. The Japanese package label is the official product information.<br>
<br>
<hr>
<a name="lang-zh"></a>
<b>{zh_title}</b><br>
<br>
{zh_p1}<br>
<br>
{zh_p2}<br>
<br>
<b>{zh_net}</b><br>
<b>{zh_store}</b><br>
<br>
<b>{zh_imp}</b><br>
<br>
<hr>
<a name="lang-ko"></a>
<b>{ko_title}</b><br>
<br>
{ko_p1}<br>
<br>
{ko_p2}<br>
<br>
<b>{ko_net}</b><br>
<b>{ko_store}</b><br>
<br>
<b>{ko_imp}</b><br>
<br>
<hr>
<a name="lang-th"></a>
<b>{th_title}</b><br>
<br>
{th_p1}<br>
<br>
{th_p2}<br>
<br>
<b>{th_net}</b><br>
<b>{th_store}</b><br>
<br>
<b>{th_imp}</b><br>
<br>
<hr>
<a name="lang-vi"></a>
<b>Tieng Viet｜Ca ngu van Nhat Ban tam vi, dung ngay</b><br>
<br>
San pham thuc pham Nhat Ban lam tu ca ngu van cua Nhat, tam vi nuoc tuong va dashi. Thit mem, umami dam, hop cho bua an hang ngay hoac mon an kem de dang.<br>
<br>
Mo goi va dung ngay, hoac rac len com nong. Hop voi ochazuke, com nam, salad, dau phu lanh, mi, hoac mon nham voi do uong.<br>
<br>
<b>Khoi luong tinh:</b> 50 g / goi<br>
<b>Bao quan:</b> Nhiet do phong, tranh nang, nong va am. Sau khi mo, dong kin va dung som.<br>
<br>
<b>Luu y:</b> Thanh phan, di ung, HSD va nha san xuat: lay theo nhan tieng Nhat tren bao bi that.<br>
"""


def i18n_fs2_long() -> str:
    """PCフリースペース2: PT / Filipino。"""
    return """<a name="lang-pt"></a>
<b>Portugues｜Bonito japones temperado, pronto para consumir</b><br>
<br>
Alimento japones feito com bonito do Japao, finalizado com molho de soja e sabor dashi. Textura macia e umami intenso, facil de incluir nas refeicoes do dia a dia.<br>
<br>
Abra a embalagem e consuma puro, ou sirva sobre arroz quente. Combina com ochazuke, onigiri, salada, tofu gelado, massa ou como petisco.<br>
<br>
<b>Peso liquido:</b> 50 g / embalagem<br>
<b>Conservacao:</b> Temperatura ambiente, longe de sol, calor e umidade. Apos abrir, feche bem e consuma logo.<br>
<br>
<b>Importante:</b> Ingredientes, alergenicos, validade e fabricante: o rotulo em japones da embalagem real e a informacao oficial.<br>
<br>
<hr>
<a name="lang-tl"></a>
<b>Filipino｜Tinimplahang bonito mula Japan, handa nang kainin</b><br>
<br>
Gawa sa bonito mula Japan, may lasang toyo at dashi. Malambot, mayamang umami — madaling idagdag sa pang-araw-araw na pagkain.<br>
<br>
Buksan at kainin agad, o ilagay sa mainit na kanin. Bagay sa ochazuke, onigiri, salad, chilled tofu, pasta, o bilang pulutan.<br>
<br>
<b>Net weight:</b> 50 g / pack<br>
<b>Storage:</b> Room temperature; iwasan ang araw, init, at halumigmig. Pagkabukas, isara nang mabuti at ubusin agad.<br>
<br>
<b>Important:</b> Ingredients, allergens, best-before, at manufacturer: ang Japanese label sa aktuwal na pakete ang opisyal.<br>
"""


def i18n_sp_rich() -> str:
    """スマホFS: EN/中/韓/TH/VI。残りバイトでPCのFS1に近い厚み（PT/TLはPCのFS2）。"""
    zh_nav = ncr("简体中文")
    ko_nav = ncr("한국어")
    th_nav = ncr("ภาษาไทย")

    zh_title = ncr("简体中文｜开袋即食的日式调味鲣鱼")
    zh_p1 = ncr(
        "本产品使用日本国产鲣鱼，搭配酱油和日式高汤风味。"
        "口感松软、鲜味浓郁，适合日常用餐或作为方便的小菜。"
    )
    zh_p2 = ncr(
        "开袋后可直接食用，也可铺在热米饭上。"
        "还适合茶泡饭、饭团、沙拉、冷豆腐、意大利面，或作为佐酒小食。"
    )
    zh_meta = ncr(
        "净含量：每袋50克。"
        "保存：常温，避光防潮。开封后请密封并尽快食用。"
        "重要：配料・过敏原・保质期以包装日文标签为准。"
    )

    ko_title = ncr("한국어｜바로 즐기는 일본식 가다랑어 조미식품")
    ko_p1 = ncr(
        "일본산 가다랑어에 간장과 다시 풍미. "
        "부드러운 식감과 진한 감칠맛으로 일상 식사나 밑반찬으로 좋습니다."
    )
    ko_p2 = ncr(
        "그대로 드시거나 따뜻한 밥 위에. "
        "오차즈케, 주먹밥, 샐러드, 냉두부, 파스타, 안주로도 잘 어울립니다."
    )
    ko_meta = ncr(
        "내용량: 1봉 50g. "
        "보관: 실온, 직사광선·고온다습 피함. 개봉 후 밀봉·빨리 섭취. "
        "중요: 원재료·알레르기·소비기한은 포장 일본어 표시가 정식입니다."
    )

    th_title = ncr("ภาษาไทย｜ปลาโอญี่ปุ่นปรุงรส พร้อมรับประทาน")
    th_p1 = ncr(
        "ปลาโอจากญี่ปุ่น รสโชยุและดาชิ เนื้อนุ่ม อูมามิเข้ม เหมาะกับมื้อประจำวัน"
    )
    th_p2 = ncr(
        "เปิดซองทานได้ทันที หรือโรยบนข้าว สลัด เต้าหู้เย็น พาสตา"
    )
    th_meta = ncr(
        "50 กรัม/ซอง. เก็บอุณหภูมิห้อง. ข้อมูลสำคัญ: ดูฉลากญี่ปุ่นบนบรรจุภัณฑ์"
    )

    return f"""
<a name="international"></a>
<hr>
<b>【For international customers / 海外のお客様へ】</b><br>
<a href="#lang-en">English</a> ｜
<a href="#lang-zh">{zh_nav}</a> ｜
<a href="#lang-ko">{ko_nav}</a> ｜
<a href="#lang-th">{th_nav}</a> ｜
<a href="#lang-vi">Tieng Viet</a><br>
<br>
<a name="lang-en"></a>
<b>English｜Seasoned Bonito, Ready to Enjoy</b><br>
<br>
This Japanese food product is made with bonito from Japan and finished with a savory soy sauce and dashi-style flavor. Soft, flaky texture and rich umami — easy for everyday meals.<br>
<br>
Enjoy as-is or over warm rice. Also great with ochazuke, rice balls, salad, chilled tofu, pasta, or as a snack with drinks.<br>
<br>
<b>Net weight:</b> 50 g per bag<br>
<b>Storage:</b> Room temperature, away from sun, heat, and humidity. After opening, seal well and consume soon.<br>
<br>
<b>Important:</b> The Japanese package label is the official source for ingredients, allergens, best-before date, and manufacturer information.<br>
<br>
<hr>
<a name="lang-zh"></a>
<b>{zh_title}</b><br>
<br>
{zh_p1}<br>
<br>
{zh_p2}<br>
<br>
{zh_meta}<br>
<br>
<hr>
<a name="lang-ko"></a>
<b>{ko_title}</b><br>
<br>
{ko_p1}<br>
<br>
{ko_p2}<br>
<br>
{ko_meta}<br>
<br>
<hr>
<a name="lang-th"></a>
<b>{th_title}</b><br>
<br>
{th_p1}<br>
<br>
{th_p2}<br>
<br>
{th_meta}<br>
<br>
<hr>
<a name="lang-vi"></a>
<b>Tieng Viet｜Ca ngu van Nhat Ban tam vi, dung ngay</b><br>
<br>
San pham thuc pham Nhat Ban lam tu ca ngu van cua Nhat, tam vi nuoc tuong va dashi. Thit mem, umami dam, hop cho bua an hang ngay.<br>
<br>
Mo goi dung ngay, hoac rac len com nong. Hop voi ochazuke, com nam, salad, dau phu lanh, mi, hoac mon nham.<br>
<br>
<b>Khoi luong tinh:</b> 50 g / goi<br>
<b>Bao quan:</b> Nhiet do phong, tranh nang, nong va am. Sau khi mo, dong kin va dung som.<br>
<br>
<b>Luu y:</b> Thanh phan, di ung, HSD: lay theo nhan tieng Nhat tren bao bi that.<br>
"""


def report(name: str, html: str) -> None:
    remain = LIMIT - len(html)
    flag = "OK" if remain > 0 else "WARN over"
    print(f"{name}: {len(html)} chars, remain {remain} [{flag}]")


out = Path(__file__).resolve().parent / "paste"
out.mkdir(parents=True, exist_ok=True)

# PC 商品説明: 日本語のみ（下部CTAあり・余白に余裕）
pc_caption = wrap_pc(
    jp_story(img_w=str(PC_W), table_w=str(PC_W), with_bottom_cta=True)
)
# PC FS1 / FS2
pc_fs1 = wrap_pc(i18n_fs1_long())
pc_fs2 = wrap_pc(i18n_fs2_long())
# スマホFS: 日本語 + EN/中/韓/TH/VI（残り枠で充実。PT/TLはPC FS2）
sp = (
    '<div align="center">\n'
    + jp_story(img_w="100%", table_w="100%", with_bottom_cta=False, compact=True)
    + i18n_sp_rich()
    + "\n</div>\n"
)

(out / "katsuo_1bag_caption_PASTE.html").write_text(pc_caption, encoding="utf-8")
(out / "katsuo_1bag_fs1_PASTE.html").write_text(pc_fs1, encoding="utf-8")
(out / "katsuo_1bag_fs2_PASTE.html").write_text(pc_fs2, encoding="utf-8")
(out / "katsuo_1bag_caption_PASTE_sp.html").write_text(sp, encoding="utf-8")

report("PC caption", pc_caption)
report("PC FS1", pc_fs1)
report("PC FS2", pc_fs2)
report("SP", sp)

assert "&#31616;" in pc_fs1, "FS1 must contain NCR for 简"
assert "&#31616;" in sp, "SP must contain NCR for 简"
assert "lang-th" in pc_fs1 and "lang-vi" in pc_fs1
assert "lang-pt" in pc_fs2 and "lang-tl" in pc_fs2
assert "lang-zh" not in pc_caption  # 商品説明に多言語なし
assert "lang-th" in sp and "lang-vi" in sp
assert "lang-pt" not in sp  # スマホはPT/TLなし
assert all(len(x) < LIMIT for x in (pc_caption, pc_fs1, pc_fs2, sp))
print("ALL_OK")
