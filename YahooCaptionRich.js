// =======================================================
// Yahoo 装飾説明（caption / FS）テンプレ組立＋API投入
// 要件: docs/org/YAHOO_CAPTION_TEMPLATE_LOCK_REQUIREMENTS.md 案A
// 正本HTML: tools/yahoo_page_mock/paste/（本ファイルは生成物。再生成は tools/yahoo_page_mock/_gen_yahoo_caption_rich_js.py）
// =======================================================

var YAHOO_CAPTION_TEMPLATE_ID_COL = '▼マスタ(Yahoo説明テンプレID)';
var YAHOO_CAPTION_IMAGE1_COL = '▼マスタ(Yahoo説明画像1)';
var YAHOO_CAPTION_IMAGE2_COL = '▼マスタ(Yahoo説明画像2)';
var YAHOO_CAPTION_IMAGE3_COL = '▼マスタ(Yahoo説明画像3)';
var YAHOO_CAPTION_RICH_LIB_BASE = 'https://shopping.c.yimg.jp/lib/octas/';
var YAHOO_CAPTION_RICH_STORE_ITEM_BASE_ = 'https://store.shopping.yahoo.co.jp/octas';
var YAHOO_CAPTION_RICH_P6_IMG_PROP_ = 'YAHOO_CAPTION_RICH_P6_IMG_ENABLED';
var YAHOO_CAPTION_RICH_P6C_PROP_ = 'YAHOO_CAPTION_RICH_P6C_THEMES_ENABLED';
var YAHOO_CAPTION_RICH_RAKUTEN_CABINET_ = 'https://image.rakuten.co.jp/octas/cabinet';
// P6C_PACK_BEGIN
var YAHOO_CAPTION_P6C_PACK_ = {
  "ishi-4906283045119-oya2": {
    "slug": "5119_oya2",
    "parentSku": "ishi-4906283045119-oya2",
    "jan": "4906283045119",
    "ranked": [
      {
        "themeId": 3,
        "catch": "噛むほどに溢れる、濃厚なだしの旨味。",
        "overview": "焼津の石原水産が手掛ける、食べるおだしのかつおとぶりを一度に楽しめるセットです。しっとり柔らかな食感から噛むほどに濃厚なだしの旨味が溢れ、じゅわっと広がる芳醇な香りを堪能できます。",
        "subs": [
          "しっとり柔らかな食感と、深みのある飴色の輝き。"
        ],
        "salesTalk": "",
        "libFile": "5119_oya2_s03_t03_shizuru_sizzle.jpg",
        "overviewSp": "食べるおだしのかつおとぶりを楽しめるセット。噛むほどに濃厚なだしの旨味が溢れます。",
        "subsSp": [
          "しっとり柔らかな食感。"
        ]
      },
      {
        "themeId": 10,
        "catch": "炊きたてご飯が止まらない、至福の相性。",
        "overview": "かつおとぶりの濃厚なだしがお米の甘みを引き立て、炊きたてご飯との相性は抜群です。おにぎりの具にすれば冷めても旨味が続き、最後の一口まで至福の余韻を楽しめます。",
        "subs": [
          "お米の甘みを引き立てる、濃厚なだしの余韻。"
        ],
        "salesTalk": "",
        "libFile": "5119_oya2_s06_t10_tabekata_howto.jpg",
        "overviewSp": "かつおとぶりの濃厚なだしがお米の甘みを引き立て、炊きたてご飯との相性は抜群です。おにぎりの具にすれば冷めても旨味が続き、最後の一口まで至福の余韻を楽しめます。",
        "subsSp": []
      },
      {
        "themeId": 1,
        "catch": "献立に迷う忙しい夜の救世主。",
        "overview": "かつおもぶりも開けるだけで一品が完成し、味付け不要で手軽に本格的な味わいを用意できます。献立に迷う忙しい夜でも準備の手間を減らし、食卓に濃厚なだしの美味しさを届けられます。",
        "subs": [
          "開けるだけで一品完成、味付け不要の時短おだし。"
        ],
        "salesTalk": "",
        "libFile": "5119_oya2_s01_t01_nayami_kyokan.jpg",
        "overviewSp": "かつおもぶりも開けるだけで一品が完成し、味付け不要で手軽に本格的な味わいを用意できます。献立に迷う忙しい夜でも準備の手間を減らし、食卓に濃厚なだしの美味しさを届けられます。",
        "subsSp": []
      }
    ]
  },
  "5119_oya2": {
    "slug": "5119_oya2",
    "parentSku": "ishi-4906283045119-oya2",
    "jan": "4906283045119",
    "ranked": [
      {
        "themeId": 3,
        "catch": "噛むほどに溢れる、濃厚なだしの旨味。",
        "overview": "焼津の石原水産が手掛ける、食べるおだしのかつおとぶりを一度に楽しめるセットです。しっとり柔らかな食感から噛むほどに濃厚なだしの旨味が溢れ、じゅわっと広がる芳醇な香りを堪能できます。",
        "subs": [
          "しっとり柔らかな食感と、深みのある飴色の輝き。"
        ],
        "salesTalk": "",
        "libFile": "5119_oya2_s03_t03_shizuru_sizzle.jpg",
        "overviewSp": "食べるおだしのかつおとぶりを楽しめるセット。噛むほどに濃厚なだしの旨味が溢れます。",
        "subsSp": [
          "しっとり柔らかな食感。"
        ]
      },
      {
        "themeId": 10,
        "catch": "炊きたてご飯が止まらない、至福の相性。",
        "overview": "かつおとぶりの濃厚なだしがお米の甘みを引き立て、炊きたてご飯との相性は抜群です。おにぎりの具にすれば冷めても旨味が続き、最後の一口まで至福の余韻を楽しめます。",
        "subs": [
          "お米の甘みを引き立てる、濃厚なだしの余韻。"
        ],
        "salesTalk": "",
        "libFile": "5119_oya2_s06_t10_tabekata_howto.jpg",
        "overviewSp": "かつおとぶりの濃厚なだしがお米の甘みを引き立て、炊きたてご飯との相性は抜群です。おにぎりの具にすれば冷めても旨味が続き、最後の一口まで至福の余韻を楽しめます。",
        "subsSp": []
      },
      {
        "themeId": 1,
        "catch": "献立に迷う忙しい夜の救世主。",
        "overview": "かつおもぶりも開けるだけで一品が完成し、味付け不要で手軽に本格的な味わいを用意できます。献立に迷う忙しい夜でも準備の手間を減らし、食卓に濃厚なだしの美味しさを届けられます。",
        "subs": [
          "開けるだけで一品完成、味付け不要の時短おだし。"
        ],
        "salesTalk": "",
        "libFile": "5119_oya2_s01_t01_nayami_kyokan.jpg",
        "overviewSp": "かつおもぶりも開けるだけで一品が完成し、味付け不要で手軽に本格的な味わいを用意できます。献立に迷う忙しい夜でも準備の手間を減らし、食卓に濃厚なだしの美味しさを届けられます。",
        "subsSp": []
      }
    ]
  },
  "ishi-4906283045119-oya3": {
    "slug": "5119_oya3",
    "parentSku": "ishi-4906283045119-oya3",
    "jan": "4906283045119",
    "ranked": [
      {
        "themeId": 3,
        "catch": "噛むほどに溢れる、濃厚なだしの旨味。",
        "overview": "食べるおだしのかつお、まぐろ、ぶりという3種を一度に楽しめるセットです。じゅわっと広がる醤油とだしの芳醇な香りと、噛むほどにしっとり柔らかな食感から濃厚なだしの旨味が溢れます。",
        "subs": [
          "しっとり柔らかな食感と、深みのある飴色の輝き。"
        ],
        "salesTalk": "",
        "libFile": "5119_oya3_s03_t03_shizuru_sizzle.jpg",
        "overviewSp": "かつお、まぐろ、ぶりの3種セット。濃厚なだしの旨味が溢れます。",
        "subsSp": [
          "しっとり柔らかな食感。"
        ]
      },
      {
        "themeId": 10,
        "catch": "炊きたてご飯が止まらない、至福の相性。",
        "overview": "かつお、まぐろ、ぶりの濃厚なだしの余韻が、炊きたてのお米の甘みを贅沢に引き立てます。おにぎりの具にすれば冷めても旨味がしっかりと続き、毎日の食卓でご飯が進む美味しさを届けます。",
        "subs": [
          "お米の甘みを引き立てる、濃厚なだしの余韻。"
        ],
        "salesTalk": "",
        "libFile": "5119_oya3_s08_t10_tabekata_howto.jpg",
        "overviewSp": "濃厚なだしの余韻が、お米の甘みを引き立てます。",
        "subsSp": []
      },
      {
        "themeId": 1,
        "catch": "献立に迷う忙しい夜の救世主。",
        "overview": "開けるだけで味付け不要の一品が完成し、献立に迷う忙しい夜の準備を手軽に助けます。かつおやまぐろ、ぶりの本格的なだしの味わいがそのまま楽しめ、夕食の準備の焦りを笑顔に変えてくれます。",
        "subs": [
          "開けるだけで一品完成、味付け不要の時短おだし。"
        ],
        "salesTalk": "",
        "libFile": "5119_oya3_s01_t01_nayami_kyokan.jpg",
        "overviewSp": "開けるだけで味付け不要の一品が完成し、献立に迷う忙しい夜の準備を手軽に助けます。かつおやまぐろ、ぶりの本格的なだしの味わいがそのまま楽しめ、夕食の準備の焦りを笑顔に変えてくれます。",
        "subsSp": []
      }
    ]
  },
  "5119_oya3": {
    "slug": "5119_oya3",
    "parentSku": "ishi-4906283045119-oya3",
    "jan": "4906283045119",
    "ranked": [
      {
        "themeId": 3,
        "catch": "噛むほどに溢れる、濃厚なだしの旨味。",
        "overview": "食べるおだしのかつお、まぐろ、ぶりという3種を一度に楽しめるセットです。じゅわっと広がる醤油とだしの芳醇な香りと、噛むほどにしっとり柔らかな食感から濃厚なだしの旨味が溢れます。",
        "subs": [
          "しっとり柔らかな食感と、深みのある飴色の輝き。"
        ],
        "salesTalk": "",
        "libFile": "5119_oya3_s03_t03_shizuru_sizzle.jpg",
        "overviewSp": "かつお、まぐろ、ぶりの3種セット。濃厚なだしの旨味が溢れます。",
        "subsSp": [
          "しっとり柔らかな食感。"
        ]
      },
      {
        "themeId": 10,
        "catch": "炊きたてご飯が止まらない、至福の相性。",
        "overview": "かつお、まぐろ、ぶりの濃厚なだしの余韻が、炊きたてのお米の甘みを贅沢に引き立てます。おにぎりの具にすれば冷めても旨味がしっかりと続き、毎日の食卓でご飯が進む美味しさを届けます。",
        "subs": [
          "お米の甘みを引き立てる、濃厚なだしの余韻。"
        ],
        "salesTalk": "",
        "libFile": "5119_oya3_s08_t10_tabekata_howto.jpg",
        "overviewSp": "濃厚なだしの余韻が、お米の甘みを引き立てます。",
        "subsSp": []
      },
      {
        "themeId": 1,
        "catch": "献立に迷う忙しい夜の救世主。",
        "overview": "開けるだけで味付け不要の一品が完成し、献立に迷う忙しい夜の準備を手軽に助けます。かつおやまぐろ、ぶりの本格的なだしの味わいがそのまま楽しめ、夕食の準備の焦りを笑顔に変えてくれます。",
        "subs": [
          "開けるだけで一品完成、味付け不要の時短おだし。"
        ],
        "salesTalk": "",
        "libFile": "5119_oya3_s01_t01_nayami_kyokan.jpg",
        "overviewSp": "開けるだけで味付け不要の一品が完成し、献立に迷う忙しい夜の準備を手軽に助けます。かつおやまぐろ、ぶりの本格的なだしの味わいがそのまま楽しめ、夕食の準備の焦りを笑顔に変えてくれます。",
        "subsSp": []
      }
    ]
  },
  "maguro": {
    "slug": "maguro",
    "parentSku": "",
    "jan": "4906283045317",
    "ranked": [
      {
        "themeId": 2,
        "catch": "罪悪感ゼロで楽しむ夜の贅沢時間",
        "overview": "国産まぐろを主原料とした食べるおだし（まぐろ）は、噛むたびに燻製の芳醇な香りが広がります。高タンパク・低糖質で体も喜ぶため、夜の贅沢時間にも罪悪感ゼロでお楽しみいただけます。",
        "subs": [
          "高タンパク・低糖質で体も喜ぶ"
        ],
        "salesTalk": "",
        "libFile": "maguro_s02_t02_benefit_mirai.jpg",
        "overviewSp": "国産まぐろを主原料とした食べるおだし。高タンパク・低糖質で、夜の贅沢時間にも罪悪感ゼロでお楽しみいただけます。",
        "subsSp": [
          "噛むたび広がる燻製の芳醇な香り"
        ]
      },
      {
        "themeId": 4,
        "catch": "噛むほどに深まる旨塩の余韻",
        "overview": "国産まぐろを使用した食べるおだし（まぐろ）は、燻製の香ばしさとまぐろの甘みが特長です。絶妙な塩加減が後を引く美味しさで、噛むほどに深まる旨塩の余韻を存分にご堪能いただけます。",
        "subs": [
          "燻製の香ばしさとまぐろの甘み"
        ],
        "salesTalk": "",
        "libFile": "maguro_s03_t04_aji_kitai.jpg",
        "overviewSp": "",
        "subsSp": [
          "燻製の香ばしさとまぐろの甘み"
        ]
      },
      {
        "themeId": 10,
        "catch": "ほかほかご飯に乗せるだけの贅沢",
        "overview": "主原料に国産まぐろを使用した食べるおだし（まぐろ）は、ひと振りで料理のコクを格上げします。ほかほかご飯に乗せるだけの贅沢はもちろん、お茶漬けで広がる燻製の香りもお楽しみいただけます。",
        "subs": [
          "お茶漬けで広がる燻製の香り"
        ],
        "salesTalk": "",
        "libFile": "maguro_s07_t10_tabekata_howto.jpg",
        "overviewSp": "",
        "subsSp": []
      }
    ]
  },
  "4906283045317": {
    "slug": "maguro",
    "parentSku": "",
    "jan": "4906283045317",
    "ranked": [
      {
        "themeId": 2,
        "catch": "罪悪感ゼロで楽しむ夜の贅沢時間",
        "overview": "国産まぐろを主原料とした食べるおだし（まぐろ）は、噛むたびに燻製の芳醇な香りが広がります。高タンパク・低糖質で体も喜ぶため、夜の贅沢時間にも罪悪感ゼロでお楽しみいただけます。",
        "subs": [
          "高タンパク・低糖質で体も喜ぶ"
        ],
        "salesTalk": "",
        "libFile": "maguro_s02_t02_benefit_mirai.jpg",
        "overviewSp": "国産まぐろを主原料とした食べるおだし。高タンパク・低糖質で、夜の贅沢時間にも罪悪感ゼロでお楽しみいただけます。",
        "subsSp": [
          "噛むたび広がる燻製の芳醇な香り"
        ]
      },
      {
        "themeId": 4,
        "catch": "噛むほどに深まる旨塩の余韻",
        "overview": "国産まぐろを使用した食べるおだし（まぐろ）は、燻製の香ばしさとまぐろの甘みが特長です。絶妙な塩加減が後を引く美味しさで、噛むほどに深まる旨塩の余韻を存分にご堪能いただけます。",
        "subs": [
          "燻製の香ばしさとまぐろの甘み"
        ],
        "salesTalk": "",
        "libFile": "maguro_s03_t04_aji_kitai.jpg",
        "overviewSp": "",
        "subsSp": [
          "燻製の香ばしさとまぐろの甘み"
        ]
      },
      {
        "themeId": 10,
        "catch": "ほかほかご飯に乗せるだけの贅沢",
        "overview": "主原料に国産まぐろを使用した食べるおだし（まぐろ）は、ひと振りで料理のコクを格上げします。ほかほかご飯に乗せるだけの贅沢はもちろん、お茶漬けで広がる燻製の香りもお楽しみいただけます。",
        "subs": [
          "お茶漬けで広がる燻製の香り"
        ],
        "salesTalk": "",
        "libFile": "maguro_s07_t10_tabekata_howto.jpg",
        "overviewSp": "",
        "subsSp": []
      }
    ]
  },
  "buri": {
    "slug": "buri",
    "parentSku": "",
    "jan": "4906283047410",
    "ranked": [
      {
        "themeId": 3,
        "catch": "噛むほどに溢れ出す、しっとり濃厚な出汁の旨味。",
        "overview": "国産の天然ぶりを主原料に、伝統の燻製技術で深みを引き出した「食べるおだし（ぶり）」です。噛むほどにしっとり濃厚な出汁の旨味が溢れ出す贅沢な味わいを、ぜひご堪能ください。",
        "subs": [
          "伝統の燻製技術が引き出す、天然ぶりの深み"
        ],
        "salesTalk": "",
        "libFile": "buri_s02_t03_shizuru_sizzle.jpg",
        "overviewSp": "国産の天然ぶりを主原料に、深みを引き出しました。",
        "subsSp": [
          "天然ぶりの深み"
        ]
      },
      {
        "themeId": 2,
        "catch": "袋を開けるだけで、極上の晩酌がすぐ始まる。",
        "overview": "主原料に国産ぶりを使用した「食べるおだし（ぶり）」は、常温保存でいつでも本格的な魚の旨味を味わえます。袋を開けるだけで極上の晩酌がすぐに始まり、本格的なぶりの美味しさを手軽にお楽しみいただけます。",
        "subs": [
          "常温保存で、いつでも本格的な魚の旨味を"
        ],
        "salesTalk": "",
        "libFile": "buri_s04_t02_benefit_mirai.jpg",
        "overviewSp": "常温保存でいつでも本格的な魚の旨味を味わえます。",
        "subsSp": []
      },
      {
        "themeId": 4,
        "catch": "生姜醤油が香る、ふっくら柔らかな「なまり節」。",
        "overview": "国産ぶり本来の味を活かし、生姜醤油が香る濃密で上品な味付けに仕上げた「食べるおだし（ぶり）」です。ふっくら柔らかな「なまり節」ならではの豊かな風味を、ぜひご堪能ください。",
        "subs": [
          "魚本来の味を活かした、濃密で上品な味付け"
        ],
        "salesTalk": "",
        "libFile": "buri_s03_t04_aji_kitai.jpg",
        "overviewSp": "",
        "subsSp": []
      }
    ]
  },
  "4906283047410": {
    "slug": "buri",
    "parentSku": "",
    "jan": "4906283047410",
    "ranked": [
      {
        "themeId": 3,
        "catch": "噛むほどに溢れ出す、しっとり濃厚な出汁の旨味。",
        "overview": "国産の天然ぶりを主原料に、伝統の燻製技術で深みを引き出した「食べるおだし（ぶり）」です。噛むほどにしっとり濃厚な出汁の旨味が溢れ出す贅沢な味わいを、ぜひご堪能ください。",
        "subs": [
          "伝統の燻製技術が引き出す、天然ぶりの深み"
        ],
        "salesTalk": "",
        "libFile": "buri_s02_t03_shizuru_sizzle.jpg",
        "overviewSp": "国産の天然ぶりを主原料に、深みを引き出しました。",
        "subsSp": [
          "天然ぶりの深み"
        ]
      },
      {
        "themeId": 2,
        "catch": "袋を開けるだけで、極上の晩酌がすぐ始まる。",
        "overview": "主原料に国産ぶりを使用した「食べるおだし（ぶり）」は、常温保存でいつでも本格的な魚の旨味を味わえます。袋を開けるだけで極上の晩酌がすぐに始まり、本格的なぶりの美味しさを手軽にお楽しみいただけます。",
        "subs": [
          "常温保存で、いつでも本格的な魚の旨味を"
        ],
        "salesTalk": "",
        "libFile": "buri_s04_t02_benefit_mirai.jpg",
        "overviewSp": "常温保存でいつでも本格的な魚の旨味を味わえます。",
        "subsSp": []
      },
      {
        "themeId": 4,
        "catch": "生姜醤油が香る、ふっくら柔らかな「なまり節」。",
        "overview": "国産ぶり本来の味を活かし、生姜醤油が香る濃密で上品な味付けに仕上げた「食べるおだし（ぶり）」です。ふっくら柔らかな「なまり節」ならではの豊かな風味を、ぜひご堪能ください。",
        "subs": [
          "魚本来の味を活かした、濃密で上品な味付け"
        ],
        "salesTalk": "",
        "libFile": "buri_s03_t04_aji_kitai.jpg",
        "overviewSp": "",
        "subsSp": []
      }
    ]
  }
};
// P6C_PACK_END

/** 食べるおだし系 JAN（共通骨格。チーズかつお混在 JAN は含めない） */
var YAHOO_CAPTION_ODASHI_JANS_ = ['4906283045119', '4906283045317', '4906283047410'];

/** 検証済みかつおテンプレ＝おだし系共通骨格（シリーズ全袋同一HTML＋P5変数埋込） */
var YAHOO_CAPTION_RICH_TMPL_KATSUO_ODASHI_ = {
  caption: "<table width=\"600\" border=\"0\" align=\"center\" cellpadding=\"0\" cellspacing=\"0\">\n<tr><td align=\"center\">\n<a name=\"top\"></a>\n<table width=\"600\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\">\n  <tr>\n    <td align=\"center\">\n      <a href=\"#sets\">\n        <img src=\"https://shopping.c.yimg.jp/lib/octas/cta_gourmet_sets.jpg\" width=\"600\" alt=\"お得なセットを見る\">\n      </a>\n    </td>\n  </tr>\n  <tr>\n    <td align=\"center\">\n      <br>\n      <a href=\"https://store.shopping.yahoo.co.jp/octas/\">\n        <img src=\"https://shopping.c.yimg.jp/lib/octas/cta_gourmet_bestsellers.jpg\" width=\"600\" alt=\"当店の売れ筋を見る\">\n      </a>\n    </td>\n  </tr>\n</table>\n<br>\n<b>忙しい日のもう一品に。</b><br>\n<b>袋を開けて、そのまま食卓へ。</b><br>\n<br>\n国産かつおを、だし香る醤油仕立てで仕上げた「食べるおだし」。<br>\nふわりとほどける食感と、噛むほどに広がる旨味が特長です。<br>\n<br>\n温かいご飯にのせるだけでもう一品。<br>\nお茶漬け・卵かけご飯・おにぎり・サラダ・冷奴・パスタやおつまみにも使えます。<br>\n\n<br>\n<img src=\"https://shopping.c.yimg.jp/lib/octas/katsuo_body_t03_sizzle.jpg\" width=\"600\" alt=\"シズル・食感\"><br>\n<br>\n<img src=\"https://shopping.c.yimg.jp/lib/octas/katsuo_body_t04_aji.jpg\" width=\"600\" alt=\"味の正しいアピール\"><br>\n<br>\n<img src=\"https://shopping.c.yimg.jp/lib/octas/katsuo_body_t11_arrange.jpg\" width=\"600\" alt=\"アレンジレシピ\"><br>\n<br>\n<b>おすすめの食べ方</b><br>\n01　そのまま、おつまみに<br>\n02　温かいご飯や卵かけご飯に<br>\n03　お茶漬け、おにぎり、サラダ、冷奴に<br>\n04　パスタや豆腐料理のトッピングに<br>\n\n<br>\n<a name=\"sets\"></a>\n<hr>\n<b>【セットコーナー】</b><br>\n袋数を選ぶと、各セットの商品ページへ移動します。<br>\n<br>\n<table width=\"600\" border=\"0\" cellpadding=\"8\" cellspacing=\"0\">\n<tr>\n<td width=\"32%\" bgcolor=\"#f5efe4\" align=\"center\"><a href=\"https://store.shopping.yahoo.co.jp/octas/ishi-4906283045119-64s18.html\"><b><br>2袋セット<br>　</b></a></td>\n<td width=\"2%\">&nbsp;</td>\n<td width=\"32%\" bgcolor=\"#f5efe4\" align=\"center\"><a href=\"https://store.shopping.yahoo.co.jp/octas/ishi-4906283045119-97s12.html\"><b><br>3袋セット<br>　</b></a></td>\n<td width=\"2%\">&nbsp;</td>\n<td width=\"32%\" bgcolor=\"#f5efe4\" align=\"center\"><a href=\"https://store.shopping.yahoo.co.jp/octas/ishi-4906283045119-12s196.html\"><b><br>4袋セット<br>　</b></a></td>\n</tr>\n<tr><td colspan=\"5\" height=\"8\"></td></tr>\n<tr>\n<td width=\"32%\" bgcolor=\"#f5efe4\" align=\"center\"><a href=\"https://store.shopping.yahoo.co.jp/octas/ishi-4906283045119-16s120.html\"><b><br>5袋セット<br>　</b></a></td>\n<td width=\"2%\">&nbsp;</td>\n<td width=\"32%\" bgcolor=\"#f5efe4\" align=\"center\"><a href=\"https://store.shopping.yahoo.co.jp/octas/ishi-4906283045119-19s144.html\"><b><br>6袋セット<br>　</b></a></td>\n<td width=\"2%\">&nbsp;</td>\n<td width=\"32%\" bgcolor=\"#f5efe4\" align=\"center\"><a href=\"https://store.shopping.yahoo.co.jp/octas/ishi-4906283045119-25s192.html\"><b><br>8袋セット<br>　</b></a></td>\n</tr>\n<tr><td colspan=\"5\" height=\"8\"></td></tr>\n<tr>\n<td width=\"32%\" bgcolor=\"#f5efe4\" align=\"center\"><a href=\"https://store.shopping.yahoo.co.jp/octas/ishi-4906283045119-32s140.html\"><b><br>10袋セット<br>　</b></a></td>\n<td width=\"2%\">&nbsp;</td>\n<td width=\"32%\" bgcolor=\"#f5efe4\" align=\"center\"><a href=\"https://store.shopping.yahoo.co.jp/octas/ishi-4906283045119-38s188.html\"><b><br>12袋セット<br>　</b></a></td>\n<td width=\"2%\">&nbsp;</td>\n<td width=\"32%\" bgcolor=\"#f5efe4\" align=\"center\"><a href=\"https://store.shopping.yahoo.co.jp/octas/ishi-4906283045119-48s160.html\"><b><br>15袋セット<br>　</b></a></td>\n</tr>\n</table>\n<br>\n<hr>\n<b>【商品について】</b><br>\n名称：食べるおだし（かつお）<br>\n内容量：1袋あたり50g<br>\n主原料：国産かつお<br>\n保存：直射日光・高温多湿を避けて常温<br>\n<br>\n※原材料・アレルゲン・賞味期限はパッケージ表示が正式です。<br>\n※クーポンはカート画面でもご確認ください。<br>\n\n<br>\n<a href=\"#sets\"><img src=\"https://shopping.c.yimg.jp/lib/octas/cta_gourmet_sets.jpg\" width=\"600\" alt=\"お得なセットを見る\"></a><br>\n<br>\n<a href=\"https://store.shopping.yahoo.co.jp/octas/\"><img src=\"https://shopping.c.yimg.jp/lib/octas/cta_gourmet_bestsellers.jpg\" width=\"600\" alt=\"当店の売れ筋を見る\"></a><br>\n\n</td></tr>\n</table>\n",
  additional1: "<table width=\"600\" border=\"0\" align=\"center\" cellpadding=\"0\" cellspacing=\"0\">\n<tr><td align=\"center\">\n<a name=\"international\"></a>\n<b>【For international customers / 海外のお客様へ】</b><br>\n<a href=\"#lang-en\">English</a> ｜\n<a href=\"#lang-zh\">&#31616;&#20307;&#20013;&#25991;</a> ｜\n<a href=\"#lang-ko\">&#54620;&#44397;&#50612;</a> ｜\n<a href=\"#lang-th\">&#3616;&#3634;&#3625;&#3634;&#3652;&#3607;&#3618;</a> ｜\n<a href=\"#lang-vi\">Tieng Viet</a><br>\n<br>\n<a name=\"lang-en\"></a>\n<b>English｜Seasoned Bonito, Ready to Enjoy</b><br>\n<br>\nThis Japanese food product is made with bonito from Japan and finished with a savory soy sauce and dashi-style flavor. Its tender, flaky texture and rich umami make it easy to enjoy as part of an everyday meal.<br>\n<br>\nOpen the package and enjoy it as it is, or serve it over warm rice. It also works well with ochazuke, rice balls, salads, chilled tofu, pasta, or as a savory snack with drinks.<br>\n<br>\n<b>Net weight:</b> 50 g per bag<br>\n<b>Storage:</b> Keep at room temperature away from direct sunlight, high temperatures, and humidity. After opening, seal well and consume as soon as possible.<br>\n<br>\n<b>Important:</b> Please check the Japanese label on the actual package for the complete ingredients, allergens, best-before date, and manufacturer information. The Japanese package label is the official product information.<br>\n<br>\n<hr>\n<a name=\"lang-zh\"></a>\n<b>&#31616;&#20307;&#20013;&#25991;&#65372;&#24320;&#34955;&#21363;&#39135;&#30340;&#26085;&#24335;&#35843;&#21619;&#40099;&#40060;</b><br>\n<br>\n&#26412;&#20135;&#21697;&#20351;&#29992;&#26085;&#26412;&#22269;&#20135;&#40099;&#40060;&#65292;&#25645;&#37197;&#37233;&#27833;&#21644;&#26085;&#24335;&#39640;&#27748;&#39118;&#21619;&#35843;&#21046;&#32780;&#25104;&#12290;&#21475;&#24863;&#26494;&#36719;&#26131;&#20837;&#21475;&#65292;&#40099;&#40060;&#30340;&#40092;&#21619;&#27987;&#37057;&#65292;&#36866;&#21512;&#26085;&#24120;&#29992;&#39184;&#25110;&#20316;&#20026;&#26041;&#20415;&#30340;&#23567;&#33756;&#12290;<br>\n<br>\n&#24320;&#34955;&#21518;&#21487;&#30452;&#25509;&#39135;&#29992;&#65292;&#20063;&#21487;&#38138;&#22312;&#28909;&#31859;&#39277;&#19978;&#12290;&#36824;&#36866;&#21512;&#25645;&#37197;&#33590;&#27873;&#39277;&#12289;&#39277;&#22242;&#12289;&#27801;&#25289;&#12289;&#20919;&#35910;&#33104;&#12289;&#24847;&#22823;&#21033;&#38754;&#65292;&#25110;&#20316;&#20026;&#20304;&#37202;&#23567;&#39135;&#12290;<br>\n<br>\n<b>&#20928;&#21547;&#37327;&#65306;&#27599;&#34955;50&#20811;</b><br>\n<b>&#20445;&#23384;&#26041;&#27861;&#65306;&#35831;&#32622;&#20110;&#24120;&#28201;&#29615;&#22659;&#65292;&#36991;&#20813;&#38451;&#20809;&#30452;&#23556;&#12289;&#39640;&#28201;&#21450;&#28526;&#28287;&#12290;&#24320;&#23553;&#21518;&#35831;&#23494;&#23553;&#20445;&#23384;&#24182;&#23613;&#24555;&#39135;&#29992;&#12290;</b><br>\n<br>\n<b>&#37325;&#35201;&#25552;&#31034;&#65306;&#23436;&#25972;&#30340;&#37197;&#26009;&#12289;&#36807;&#25935;&#21407;&#12289;&#20445;&#36136;&#26399;&#21450;&#21046;&#36896;&#21830;&#20449;&#24687;&#65292;&#35831;&#20197;&#23454;&#38469;&#21830;&#21697;&#21253;&#35013;&#19978;&#30340;&#26085;&#25991;&#26631;&#31614;&#20026;&#20934;&#12290;</b><br>\n<br>\n<hr>\n<a name=\"lang-ko\"></a>\n<b>&#54620;&#44397;&#50612;&#65372;&#48148;&#47196; &#51600;&#44592;&#45716; &#51068;&#48376;&#49885; &#44032;&#45796;&#46993;&#50612; &#51312;&#48120;&#49885;&#54408;</b><br>\n<br>\n&#51068;&#48376;&#49328; &#44032;&#45796;&#46993;&#50612;&#50640; &#44036;&#51109;&#44284; &#45796;&#49884; &#54413;&#48120;&#47484; &#45908;&#54620; &#51228;&#54408;&#51077;&#45768;&#45796;. &#48512;&#46300;&#47084;&#50868; &#49885;&#44048;&#44284; &#51652;&#54620; &#44048;&#52832;&#47579;&#51004;&#47196;, &#51068;&#49345; &#49885;&#49324;&#45208; &#44036;&#45800;&#54620; &#48145;&#48152;&#52268;&#51004;&#47196; &#51600;&#44592;&#44592; &#51339;&#49845;&#45768;&#45796;.<br>\n<br>\n&#44060;&#48393; &#54980; &#44536;&#45824;&#47196; &#46300;&#49884;&#44144;&#45208; &#46384;&#46907;&#54620; &#48165; &#50948;&#50640; &#50732;&#47140; &#46300;&#49464;&#50836;. &#50724;&#52264;&#51592;&#52992;, &#51452;&#47673;&#48165;, &#49360;&#47084;&#46300;, &#45257;&#46160;&#48512;, &#54028;&#49828;&#53440;, &#50504;&#51452;&#47196;&#46020; &#51096; &#50612;&#50872;&#47549;&#45768;&#45796;.<br>\n<br>\n<b>&#45236;&#50857;&#47049;: 1&#48393; 50g</b><br>\n<b>&#48372;&#44288;: &#49892;&#50728;, &#51649;&#49324;&#44305;&#49440;&#183;&#44256;&#50728;&#45796;&#49845;&#51012; &#54588;&#54616;&#49464;&#50836;. &#44060;&#48393; &#54980; &#48128;&#48393;&#54616;&#50668; &#48744;&#47532; &#49453;&#52712;&#54616;&#49464;&#50836;.</b><br>\n<br>\n<b>&#51473;&#50836;: &#50896;&#51116;&#47308;&#183;&#50508;&#47112;&#47476;&#44592;&#183;&#49548;&#48708;&#44592;&#54620;&#183;&#51228;&#51312;&#51088; &#51221;&#48372;&#45716; &#49892;&#51228; &#54252;&#51109;&#51032; &#51068;&#48376;&#50612; &#54364;&#49884;&#44032; &#51221;&#49885;&#51077;&#45768;&#45796;.</b><br>\n<br>\n<hr>\n<a name=\"lang-th\"></a>\n<b>&#3616;&#3634;&#3625;&#3634;&#3652;&#3607;&#3618;&#65372;&#3611;&#3621;&#3634;&#3650;&#3629;&#3597;&#3637;&#3656;&#3611;&#3640;&#3656;&#3609;&#3611;&#3619;&#3640;&#3591;&#3619;&#3626; &#3614;&#3619;&#3657;&#3629;&#3617;&#3619;&#3633;&#3610;&#3611;&#3619;&#3632;&#3607;&#3634;&#3609;</b><br>\n<br>\n&#3607;&#3635;&#3592;&#3634;&#3585;&#3611;&#3621;&#3634;&#3650;&#3629;&#3592;&#3634;&#3585;&#3597;&#3637;&#3656;&#3611;&#3640;&#3656;&#3609; &#3611;&#3619;&#3640;&#3591;&#3604;&#3657;&#3623;&#3618;&#3619;&#3626;&#3650;&#3594;&#3618;&#3640;&#3649;&#3621;&#3632;&#3604;&#3634;&#3594;&#3636; &#3648;&#3609;&#3639;&#3657;&#3629;&#3609;&#3640;&#3656;&#3617; &#3629;&#3641;&#3617;&#3634;&#3617;&#3636;&#3648;&#3586;&#3657;&#3617; &#3648;&#3627;&#3617;&#3634;&#3632;&#3585;&#3633;&#3610;&#3617;&#3639;&#3657;&#3629;&#3611;&#3619;&#3632;&#3592;&#3635;&#3623;&#3633;&#3609;&#3627;&#3619;&#3639;&#3629;&#3591;&#3634;&#3609;&#3623;&#3656;&#3634;&#3591;&#3591;&#3656;&#3634;&#3618;&#3654;<br>\n<br>\n&#3648;&#3611;&#3636;&#3604;&#3595;&#3629;&#3591;&#3619;&#3633;&#3610;&#3611;&#3619;&#3632;&#3607;&#3634;&#3609;&#3652;&#3604;&#3657;&#3607;&#3633;&#3609;&#3607;&#3637; &#3627;&#3619;&#3639;&#3629;&#3650;&#3619;&#3618;&#3610;&#3609;&#3586;&#3657;&#3634;&#3623;&#3619;&#3657;&#3629;&#3609; &#3648;&#3627;&#3617;&#3634;&#3632;&#3585;&#3633;&#3610;&#3650;&#3629;&#3594;&#3634;&#3595;&#3640;&#3648;&#3585;&#3632; &#3586;&#3657;&#3634;&#3623;&#3611;&#3633;&#3657;&#3609; &#3626;&#3621;&#3633;&#3604; &#3648;&#3605;&#3657;&#3634;&#3627;&#3641;&#3657;&#3648;&#3618;&#3655;&#3609; &#3614;&#3634;&#3626;&#3605;&#3634; &#3627;&#3619;&#3639;&#3629;&#3586;&#3629;&#3591;&#3623;&#3656;&#3634;&#3591;&#3588;&#3641;&#3656;&#3648;&#3588;&#3619;&#3639;&#3656;&#3629;&#3591;&#3604;&#3639;&#3656;&#3617;<br>\n<br>\n<b>&#3609;&#3657;&#3635;&#3627;&#3609;&#3633;&#3585;&#3626;&#3640;&#3607;&#3608;&#3636;: 50 &#3585;&#3619;&#3633;&#3617;&#3605;&#3656;&#3629;&#3595;&#3629;&#3591;</b><br>\n<b>&#3585;&#3634;&#3619;&#3648;&#3585;&#3655;&#3610;&#3619;&#3633;&#3585;&#3625;&#3634;: &#3629;&#3640;&#3603;&#3627;&#3616;&#3641;&#3617;&#3636;&#3627;&#3657;&#3629;&#3591; &#3627;&#3621;&#3637;&#3585;&#3648;&#3621;&#3637;&#3656;&#3618;&#3591;&#3649;&#3626;&#3591;&#3649;&#3604;&#3604; &#3588;&#3623;&#3634;&#3617;&#3619;&#3657;&#3629;&#3609; &#3649;&#3621;&#3632;&#3588;&#3623;&#3634;&#3617;&#3594;&#3639;&#3657;&#3609; &#3627;&#3621;&#3633;&#3591;&#3648;&#3611;&#3636;&#3604;&#3649;&#3621;&#3657;&#3623;&#3651;&#3627;&#3657;&#3611;&#3636;&#3604;&#3626;&#3609;&#3636;&#3607;&#3649;&#3621;&#3632;&#3619;&#3633;&#3610;&#3611;&#3619;&#3632;&#3607;&#3634;&#3609;&#3650;&#3604;&#3618;&#3648;&#3619;&#3655;&#3623;</b><br>\n<br>\n<b>&#3586;&#3657;&#3629;&#3617;&#3641;&#3621;&#3626;&#3635;&#3588;&#3633;&#3597;: &#3626;&#3656;&#3623;&#3609;&#3611;&#3619;&#3632;&#3585;&#3629;&#3610; &#3626;&#3634;&#3619;&#3585;&#3656;&#3629;&#3616;&#3641;&#3617;&#3636;&#3649;&#3614;&#3657; &#3623;&#3633;&#3609;&#3627;&#3617;&#3604;&#3629;&#3634;&#3618;&#3640; &#3649;&#3621;&#3632;&#3612;&#3641;&#3657;&#3612;&#3621;&#3636;&#3605; &#3651;&#3627;&#3657;&#3618;&#3638;&#3604;&#3593;&#3621;&#3634;&#3585;&#3616;&#3634;&#3625;&#3634;&#3597;&#3637;&#3656;&#3611;&#3640;&#3656;&#3609;&#3610;&#3609;&#3610;&#3619;&#3619;&#3592;&#3640;&#3616;&#3633;&#3603;&#3601;&#3660;&#3592;&#3619;&#3636;&#3591;&#3648;&#3611;&#3655;&#3609;&#3627;&#3621;&#3633;&#3585;</b><br>\n<br>\n<hr>\n<a name=\"lang-vi\"></a>\n<b>Tieng Viet｜Ca ngu van Nhat Ban tam vi, dung ngay</b><br>\n<br>\nSan pham thuc pham Nhat Ban lam tu ca ngu van cua Nhat, tam vi nuoc tuong va dashi. Thit mem, umami dam, hop cho bua an hang ngay hoac mon an kem de dang.<br>\n<br>\nMo goi va dung ngay, hoac rac len com nong. Hop voi ochazuke, com nam, salad, dau phu lanh, mi, hoac mon nham voi do uong.<br>\n<br>\n<b>Khoi luong tinh:</b> 50 g / goi<br>\n<b>Bao quan:</b> Nhiet do phong, tranh nang, nong va am. Sau khi mo, dong kin va dung som.<br>\n<br>\n<b>Luu y:</b> Thanh phan, di ung, HSD va nha san xuat: lay theo nhan tieng Nhat tren bao bi that.<br>\n\n</td></tr>\n</table>\n",
  additional2: "<table width=\"600\" border=\"0\" align=\"center\" cellpadding=\"0\" cellspacing=\"0\">\n<tr><td align=\"center\">\n<a name=\"lang-pt\"></a>\n<b>Portugues｜Bonito japones temperado, pronto para consumir</b><br>\n<br>\nAlimento japones feito com bonito do Japao, finalizado com molho de soja e sabor dashi. Textura macia e umami intenso, facil de incluir nas refeicoes do dia a dia.<br>\n<br>\nAbra a embalagem e consuma puro, ou sirva sobre arroz quente. Combina com ochazuke, onigiri, salada, tofu gelado, massa ou como petisco.<br>\n<br>\n<b>Peso liquido:</b> 50 g / embalagem<br>\n<b>Conservacao:</b> Temperatura ambiente, longe de sol, calor e umidade. Apos abrir, feche bem e consuma logo.<br>\n<br>\n<b>Importante:</b> Ingredientes, alergenicos, validade e fabricante: o rotulo em japones da embalagem real e a informacao oficial.<br>\n<br>\n<hr>\n<a name=\"lang-tl\"></a>\n<b>Filipino｜Tinimplahang bonito mula Japan, handa nang kainin</b><br>\n<br>\nGawa sa bonito mula Japan, may lasang toyo at dashi. Malambot, mayamang umami — madaling idagdag sa pang-araw-araw na pagkain.<br>\n<br>\nBuksan at kainin agad, o ilagay sa mainit na kanin. Bagay sa ochazuke, onigiri, salad, chilled tofu, pasta, o bilang pulutan.<br>\n<br>\n<b>Net weight:</b> 50 g / pack<br>\n<b>Storage:</b> Room temperature; iwasan ang araw, init, at halumigmig. Pagkabukas, isara nang mabuti at ubusin agad.<br>\n<br>\n<b>Important:</b> Ingredients, allergens, best-before, at manufacturer: ang Japanese label sa aktuwal na pakete ang opisyal.<br>\n\n</td></tr>\n</table>\n",
  sp_additional: "<div align=\"center\">\n<a name=\"top\"></a>\n<table width=\"100%\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\">\n  <tr>\n    <td align=\"center\">\n      <a href=\"#sets\">\n        <img src=\"https://shopping.c.yimg.jp/lib/octas/cta_gourmet_sets.jpg\" width=\"100%\" alt=\"お得なセットを見る\">\n      </a>\n    </td>\n  </tr>\n  <tr>\n    <td align=\"center\">\n      <br>\n      <a href=\"https://store.shopping.yahoo.co.jp/octas/\">\n        <img src=\"https://shopping.c.yimg.jp/lib/octas/cta_gourmet_bestsellers.jpg\" width=\"100%\" alt=\"当店の売れ筋を見る\">\n      </a>\n    </td>\n  </tr>\n</table>\n<br>\n<b>忙しい日のもう一品に。</b><br>\n<b>袋を開けて、そのまま食卓へ。</b><br>\n<br>\n国産かつおを、だし香る醤油仕立てで仕上げた「食べるおだし」。<br>\nふわりとほどける食感と、噛むほどに広がる旨味が特長です。<br>\n<br>\nご飯・お茶漬け・サラダ・冷奴・パスタやおつまみに。<br>\n\n<br>\n<img src=\"https://shopping.c.yimg.jp/lib/octas/katsuo_body_t03_sizzle.jpg\" width=\"100%\" alt=\"シズル・食感\"><br>\n<br>\n<img src=\"https://shopping.c.yimg.jp/lib/octas/katsuo_body_t04_aji.jpg\" width=\"100%\" alt=\"味の正しいアピール\"><br>\n<br>\n<img src=\"https://shopping.c.yimg.jp/lib/octas/katsuo_body_t11_arrange.jpg\" width=\"100%\" alt=\"アレンジレシピ\"><br>\n<br>\n<b>おすすめの食べ方</b><br>\n01　そのまま、おつまみに<br>\n02　温かいご飯や卵かけご飯に<br>\n03　お茶漬け、サラダ、冷奴に<br>\n\n<br>\n<a name=\"sets\"></a>\n<hr>\n<b>【セットコーナー】</b><br>\n袋数を選ぶと、各セットの商品ページへ移動します。<br>\n<br>\n<table width=\"100%\" border=\"0\" cellpadding=\"8\" cellspacing=\"0\">\n<tr>\n<td width=\"32%\" bgcolor=\"#f5efe4\" align=\"center\"><a href=\"https://store.shopping.yahoo.co.jp/octas/ishi-4906283045119-64s18.html\"><b><br>2袋セット<br>　</b></a></td>\n<td width=\"2%\">&nbsp;</td>\n<td width=\"32%\" bgcolor=\"#f5efe4\" align=\"center\"><a href=\"https://store.shopping.yahoo.co.jp/octas/ishi-4906283045119-97s12.html\"><b><br>3袋セット<br>　</b></a></td>\n<td width=\"2%\">&nbsp;</td>\n<td width=\"32%\" bgcolor=\"#f5efe4\" align=\"center\"><a href=\"https://store.shopping.yahoo.co.jp/octas/ishi-4906283045119-12s196.html\"><b><br>4袋セット<br>　</b></a></td>\n</tr>\n<tr><td colspan=\"5\" height=\"8\"></td></tr>\n<tr>\n<td width=\"32%\" bgcolor=\"#f5efe4\" align=\"center\"><a href=\"https://store.shopping.yahoo.co.jp/octas/ishi-4906283045119-16s120.html\"><b><br>5袋セット<br>　</b></a></td>\n<td width=\"2%\">&nbsp;</td>\n<td width=\"32%\" bgcolor=\"#f5efe4\" align=\"center\"><a href=\"https://store.shopping.yahoo.co.jp/octas/ishi-4906283045119-19s144.html\"><b><br>6袋セット<br>　</b></a></td>\n<td width=\"2%\">&nbsp;</td>\n<td width=\"32%\" bgcolor=\"#f5efe4\" align=\"center\"><a href=\"https://store.shopping.yahoo.co.jp/octas/ishi-4906283045119-25s192.html\"><b><br>8袋セット<br>　</b></a></td>\n</tr>\n<tr><td colspan=\"5\" height=\"8\"></td></tr>\n<tr>\n<td width=\"32%\" bgcolor=\"#f5efe4\" align=\"center\"><a href=\"https://store.shopping.yahoo.co.jp/octas/ishi-4906283045119-32s140.html\"><b><br>10袋セット<br>　</b></a></td>\n<td width=\"2%\">&nbsp;</td>\n<td width=\"32%\" bgcolor=\"#f5efe4\" align=\"center\"><a href=\"https://store.shopping.yahoo.co.jp/octas/ishi-4906283045119-38s188.html\"><b><br>12袋セット<br>　</b></a></td>\n<td width=\"2%\">&nbsp;</td>\n<td width=\"32%\" bgcolor=\"#f5efe4\" align=\"center\"><a href=\"https://store.shopping.yahoo.co.jp/octas/ishi-4906283045119-48s160.html\"><b><br>15袋セット<br>　</b></a></td>\n</tr>\n</table>\n<br>\n<hr>\n<b>【商品について】</b><br>\n名称：食べるおだし（かつお）<br>\n内容量：1袋あたり50g<br>\n主原料：国産かつお<br>\n保存：直射日光・高温多湿を避けて常温<br>\n<br>\n※原材料・アレルゲン・賞味期限はパッケージ表示が正式です。<br>\n※クーポンはカート画面でもご確認ください。<br>\n\n<a name=\"international\"></a>\n<hr>\n<b>【For international customers / 海外のお客様へ】</b><br>\n<a href=\"#lang-en\">English</a> ｜\n<a href=\"#lang-zh\">&#31616;&#20307;&#20013;&#25991;</a> ｜\n<a href=\"#lang-ko\">&#54620;&#44397;&#50612;</a> ｜\n<a href=\"#lang-th\">&#3616;&#3634;&#3625;&#3634;&#3652;&#3607;&#3618;</a> ｜\n<a href=\"#lang-vi\">Tieng Viet</a><br>\n<br>\n<a name=\"lang-en\"></a>\n<b>English｜Seasoned Bonito, Ready to Enjoy</b><br>\n<br>\nThis Japanese food product is made with bonito from Japan and finished with a savory soy sauce and dashi-style flavor. Soft, flaky texture and rich umami — easy for everyday meals.<br>\n<br>\nEnjoy as-is or over warm rice. Also great with ochazuke, rice balls, salad, chilled tofu, pasta, or as a snack with drinks.<br>\n<br>\n<b>Net weight:</b> 50 g per bag<br>\n<b>Storage:</b> Room temperature, away from sun, heat, and humidity. After opening, seal well and consume soon.<br>\n<br>\n<b>Important:</b> The Japanese package label is the official source for ingredients, allergens, best-before date, and manufacturer information.<br>\n<br>\n<hr>\n<a name=\"lang-zh\"></a>\n<b>&#31616;&#20307;&#20013;&#25991;&#65372;&#24320;&#34955;&#21363;&#39135;&#30340;&#26085;&#24335;&#35843;&#21619;&#40099;&#40060;</b><br>\n<br>\n&#26412;&#20135;&#21697;&#20351;&#29992;&#26085;&#26412;&#22269;&#20135;&#40099;&#40060;&#65292;&#25645;&#37197;&#37233;&#27833;&#21644;&#26085;&#24335;&#39640;&#27748;&#39118;&#21619;&#12290;&#21475;&#24863;&#26494;&#36719;&#12289;&#40092;&#21619;&#27987;&#37057;&#65292;&#36866;&#21512;&#26085;&#24120;&#29992;&#39184;&#25110;&#20316;&#20026;&#26041;&#20415;&#30340;&#23567;&#33756;&#12290;<br>\n<br>\n&#24320;&#34955;&#21518;&#21487;&#30452;&#25509;&#39135;&#29992;&#65292;&#20063;&#21487;&#38138;&#22312;&#28909;&#31859;&#39277;&#19978;&#12290;&#36824;&#36866;&#21512;&#33590;&#27873;&#39277;&#12289;&#39277;&#22242;&#12289;&#27801;&#25289;&#12289;&#20919;&#35910;&#33104;&#12289;&#24847;&#22823;&#21033;&#38754;&#65292;&#25110;&#20316;&#20026;&#20304;&#37202;&#23567;&#39135;&#12290;<br>\n<br>\n&#20928;&#21547;&#37327;&#65306;&#27599;&#34955;50&#20811;&#12290;&#20445;&#23384;&#65306;&#24120;&#28201;&#65292;&#36991;&#20809;&#38450;&#28526;&#12290;&#24320;&#23553;&#21518;&#35831;&#23494;&#23553;&#24182;&#23613;&#24555;&#39135;&#29992;&#12290;&#37325;&#35201;&#65306;&#37197;&#26009;&#12539;&#36807;&#25935;&#21407;&#12539;&#20445;&#36136;&#26399;&#20197;&#21253;&#35013;&#26085;&#25991;&#26631;&#31614;&#20026;&#20934;&#12290;<br>\n<br>\n<hr>\n<a name=\"lang-ko\"></a>\n<b>&#54620;&#44397;&#50612;&#65372;&#48148;&#47196; &#51600;&#44592;&#45716; &#51068;&#48376;&#49885; &#44032;&#45796;&#46993;&#50612; &#51312;&#48120;&#49885;&#54408;</b><br>\n<br>\n&#51068;&#48376;&#49328; &#44032;&#45796;&#46993;&#50612;&#50640; &#44036;&#51109;&#44284; &#45796;&#49884; &#54413;&#48120;. &#48512;&#46300;&#47084;&#50868; &#49885;&#44048;&#44284; &#51652;&#54620; &#44048;&#52832;&#47579;&#51004;&#47196; &#51068;&#49345; &#49885;&#49324;&#45208; &#48145;&#48152;&#52268;&#51004;&#47196; &#51339;&#49845;&#45768;&#45796;.<br>\n<br>\n&#44536;&#45824;&#47196; &#46300;&#49884;&#44144;&#45208; &#46384;&#46907;&#54620; &#48165; &#50948;&#50640;. &#50724;&#52264;&#51592;&#52992;, &#51452;&#47673;&#48165;, &#49360;&#47084;&#46300;, &#45257;&#46160;&#48512;, &#54028;&#49828;&#53440;, &#50504;&#51452;&#47196;&#46020; &#51096; &#50612;&#50872;&#47549;&#45768;&#45796;.<br>\n<br>\n&#45236;&#50857;&#47049;: 1&#48393; 50g. &#48372;&#44288;: &#49892;&#50728;, &#51649;&#49324;&#44305;&#49440;&#183;&#44256;&#50728;&#45796;&#49845; &#54588;&#54632;. &#44060;&#48393; &#54980; &#48128;&#48393;&#183;&#48744;&#47532; &#49453;&#52712;. &#51473;&#50836;: &#50896;&#51116;&#47308;&#183;&#50508;&#47112;&#47476;&#44592;&#183;&#49548;&#48708;&#44592;&#54620;&#51008; &#54252;&#51109; &#51068;&#48376;&#50612; &#54364;&#49884;&#44032; &#51221;&#49885;&#51077;&#45768;&#45796;.<br>\n<br>\n<hr>\n<a name=\"lang-th\"></a>\n<b>&#3616;&#3634;&#3625;&#3634;&#3652;&#3607;&#3618;&#65372;&#3611;&#3621;&#3634;&#3650;&#3629;&#3597;&#3637;&#3656;&#3611;&#3640;&#3656;&#3609;&#3611;&#3619;&#3640;&#3591;&#3619;&#3626; &#3614;&#3619;&#3657;&#3629;&#3617;&#3619;&#3633;&#3610;&#3611;&#3619;&#3632;&#3607;&#3634;&#3609;</b><br>\n<br>\n&#3611;&#3621;&#3634;&#3650;&#3629;&#3592;&#3634;&#3585;&#3597;&#3637;&#3656;&#3611;&#3640;&#3656;&#3609; &#3619;&#3626;&#3650;&#3594;&#3618;&#3640;&#3649;&#3621;&#3632;&#3604;&#3634;&#3594;&#3636; &#3648;&#3609;&#3639;&#3657;&#3629;&#3609;&#3640;&#3656;&#3617; &#3629;&#3641;&#3617;&#3634;&#3617;&#3636;&#3648;&#3586;&#3657;&#3617; &#3648;&#3627;&#3617;&#3634;&#3632;&#3585;&#3633;&#3610;&#3617;&#3639;&#3657;&#3629;&#3611;&#3619;&#3632;&#3592;&#3635;&#3623;&#3633;&#3609;<br>\n<br>\n&#3648;&#3611;&#3636;&#3604;&#3595;&#3629;&#3591;&#3607;&#3634;&#3609;&#3652;&#3604;&#3657;&#3607;&#3633;&#3609;&#3607;&#3637; &#3627;&#3619;&#3639;&#3629;&#3650;&#3619;&#3618;&#3610;&#3609;&#3586;&#3657;&#3634;&#3623; &#3626;&#3621;&#3633;&#3604; &#3648;&#3605;&#3657;&#3634;&#3627;&#3641;&#3657;&#3648;&#3618;&#3655;&#3609; &#3614;&#3634;&#3626;&#3605;&#3634;<br>\n<br>\n50 &#3585;&#3619;&#3633;&#3617;/&#3595;&#3629;&#3591;. &#3648;&#3585;&#3655;&#3610;&#3629;&#3640;&#3603;&#3627;&#3616;&#3641;&#3617;&#3636;&#3627;&#3657;&#3629;&#3591;. &#3586;&#3657;&#3629;&#3617;&#3641;&#3621;&#3626;&#3635;&#3588;&#3633;&#3597;: &#3604;&#3641;&#3593;&#3621;&#3634;&#3585;&#3597;&#3637;&#3656;&#3611;&#3640;&#3656;&#3609;&#3610;&#3609;&#3610;&#3619;&#3619;&#3592;&#3640;&#3616;&#3633;&#3603;&#3601;&#3660;<br>\n<br>\n<hr>\n<a name=\"lang-vi\"></a>\n<b>Tieng Viet｜Ca ngu van Nhat Ban tam vi, dung ngay</b><br>\n<br>\nSan pham thuc pham Nhat Ban lam tu ca ngu van cua Nhat, tam vi nuoc tuong va dashi. Thit mem, umami dam, hop cho bua an hang ngay.<br>\n<br>\nMo goi dung ngay, hoac rac len com nong. Hop voi ochazuke, com nam, salad, dau phu lanh, mi, hoac mon nham.<br>\n<br>\n<b>Khoi luong tinh:</b> 50 g / goi<br>\n<b>Bao quan:</b> Nhiet do phong, tranh nang, nong va am. Sau khi mo, dong kin va dung som.<br>\n<br>\n<b>Luu y:</b> Thanh phan, di ung, HSD: lay theo nhan tieng Nhat tren bao bi that.<br>\n\n</div>\n"
};

/**
 * テンプレID → HTML4欄。未知IDは null。
 * @param {string} templateId
 * @param {Object} [vars] img1..3 / fishJa / netWeight / setItems 等
 * @return {{caption: string, additional1: string, additional2: string, sp_additional: string}|null}
 */
function yahooCaptionRichBuildHtml_(templateId, vars) {
  var id = String(templateId || '').trim();
  if (!id) return null;
  // 互換: katsuo_odashi＝おだし共通骨格。odashi_gourmet も同一。
  if (id === 'odashi_gourmet') id = 'katsuo_odashi';
  var base = null;
  if (id === 'katsuo_odashi') {
    base = YAHOO_CAPTION_RICH_TMPL_KATSUO_ODASHI_;
  }
  if (!base) return null;
  vars = vars || {};
  var out = {
    caption: String(base.caption || ''),
    additional1: String(base.additional1 || ''),
    additional2: String(base.additional2 || ''),
    sp_additional: String(base.sp_additional || '')
  };
  out = yahooCaptionRichApplyOdashiVars_(out, vars);
  out = yahooCaptionRichApplyBodyImages_(out, vars, ['katsuo_body_t03_sizzle.jpg', 'katsuo_body_t04_aji.jpg', 'katsuo_body_t11_arrange.jpg']);
  if (yahooCaptionRichP6cEnabled_()) {
    var ranked = vars.rankedThemes || yahooCaptionRichLookupP6cRanked_(vars.jan, vars.slug, vars.parentSku);
    if (ranked && ranked.length) {
      out = yahooCaptionRichApplyRankedThemes_(out, ranked);
      console.log('[YAHOO_CAPTION_RICH] p6c_themes n=' + ranked.length +
        ' jan=' + String(vars.jan || '') +
        ' parent=' + String(vars.parentSku || '') +
        ' slug=' + String(vars.slug || ''));
    }
  }
  // 案C: 【商品について】枠廃止。注記は残す。原料は caption のみ。
  var stripKeys = ['caption', 'additional1', 'additional2', 'sp_additional'];
  var matJa = String(vars.materialJa || '').trim();
  var showMat = !!vars.showMaterial && !!matJa;
  for (var si = 0; si < stripKeys.length; si++) {
    var sk = stripKeys[si];
    out[sk] = yahooCaptionRichStripContentQtyLines_(out[sk]);
    out[sk] = yahooCaptionRichApplyProductAboutC_(
      out[sk], matJa, showMat && sk === 'caption');
  }
  return out;
}

function yahooCaptionRichP6cEnabled_() {
  if (typeof getBoolScriptProperty_ === 'function') {
    return getBoolScriptProperty_(YAHOO_CAPTION_RICH_P6C_PROP_, false);
  }
  return false;
}

/** 親SKU → slug → JAN（同一JAN複数親は親キー必須） */
function yahooCaptionRichLookupP6cRanked_(jan, slug, parentSku) {
  var pack = (typeof YAHOO_CAPTION_P6C_PACK_ !== 'undefined' && YAHOO_CAPTION_P6C_PACK_) || {};
  var rec =
    pack[String(parentSku || '')] ||
    pack[String(slug || '')] ||
    pack[String(jan || '')];
  if (!rec) return [];
  return rec.ranked || [];
}

function yahooCaptionRichHtmlEscape_(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

var YAHOO_CAPTION_RICH_CAPTION_LIMIT_ = 5000;
var YAHOO_CAPTION_RICH_SP_LIMIT_ = 10000;

function yahooCaptionRichStripNetQty_(s) {
  s = String(s || '');
  s = s.replace(/[０-９]/g, function (ch) {
    return String.fromCharCode(ch.charCodeAt(0) - 0xFF10 + 48);
  });
  s = s.replace(/[ｇＧ]/g, 'g');
  s = s.replace(/1\s*袋\s*(?:あたり|当たり)?\s*[\/／]?\s*\d+(?:\.\d+)?\s*(?:kg|mg|ml|mL|cc|g|l|L|グラム|リットル)?/gi, '');
  s = s.replace(/(?:の)?\s*\d+(?:\.\d+)?\s*(?:kg|mg|ml|mL|cc|g|l|L|グラム|キログラム|ミリリットル|リットル)\s*(?:入り)?の?/gi, '');
  s = s.replace(/\s{2,}/g, ' ').replace(/のの/g, 'の').replace(/。+/g, '。');
  return s.replace(/^\s+|\s+$/g, '');
}

function yahooCaptionRichThemeOverview_(r, isSp) {
  var raw = '';
  if (isSp && Object.prototype.hasOwnProperty.call(r, 'overviewSp')) {
    raw = r.overviewSp || '';
    if (!String(raw).trim()) raw = r.overview || '';
  } else {
    raw = r.overview || '';
  }
  return yahooCaptionRichStripNetQty_(raw);
}

function yahooCaptionRichThemeSub_(r, isSp) {
  var list = (isSp && r.subsSp) ? r.subsSp : (r.subs || []);
  return (list && list[0]) || r.salesTalk || '';
}

function yahooCaptionRichBuildThemeSlotsHtml_(ranked, imgW, mode) {
  mode = mode || 'full';
  var isSp = String(imgW) === '100%';
  var parts = [];
  var i;
  for (i = 0; i < ranked.length && i < 3; i++) {
    var r = ranked[i] || {};
    var catchT = yahooCaptionRichHtmlEscape_(r.catch || '');
    var ov = '';
    var sub = '';
    if (mode !== 'catch_img') {
      if (mode !== 'no_ov') ov = yahooCaptionRichHtmlEscape_(yahooCaptionRichThemeOverview_(r, isSp));
      if (mode !== 'no_sub') sub = yahooCaptionRichHtmlEscape_(yahooCaptionRichThemeSub_(r, isSp));
    }
    var fn = String(r.libFile || '').trim();
    if (!fn) continue;
    var url = /^https?:\/\//i.test(fn) ? fn : (YAHOO_CAPTION_RICH_LIB_BASE + fn);
    if (!yahooCaptionRichIsYahooSafeImgUrl_(url)) continue;
    parts.push('<b>' + catchT + '</b><br>');
    if (sub) parts.push(sub + '<br>');
    if (ov) parts.push('<br>' + ov + '<br>');
    parts.push('<br><img src="' + url + '" width="' + imgW + '" alt="' + catchT + '"><br><br>');
  }
  return parts.join('\n');
}

function yahooCaptionRichReplaceThemeBlock_(html, re, ranked, imgW, limit, fieldName) {
  if (!html || !re.test(html)) return html;
  var modes = ['full', 'no_sub', 'no_ov', 'catch_img'];
  var i;
  var last = html;
  for (i = 0; i < modes.length; i++) {
    var block = yahooCaptionRichBuildThemeSlotsHtml_(ranked, imgW, modes[i]);
    if (!block) return html;
    var out = html.replace(re, block + '\n');
    last = out;
    if (out.length <= limit) {
      console.log('[YAHOO_CAPTION_RICH] p6c_fit field=' + fieldName +
        ' mode=' + modes[i] + ' len=' + out.length + ' limit=' + limit);
      return out;
    }
  }
  console.log('[YAHOO_CAPTION_RICH] p6c_skip_over field=' + fieldName +
    ' len=' + last.length + ' limit=' + limit);
  return html;
}

function yahooCaptionRichApplyRankedThemes_(htmlParts, ranked) {
  if (!ranked || !ranked.length) return htmlParts;
  var re = /<b>忙しい日のもう一品に。<\/b>[\s\S]*?(?=<a name="sets"><\/a>)/;
  htmlParts.caption = yahooCaptionRichReplaceThemeBlock_(
    htmlParts.caption, re, ranked, '600', YAHOO_CAPTION_RICH_CAPTION_LIMIT_, 'caption');
  htmlParts.sp_additional = yahooCaptionRichReplaceThemeBlock_(
    htmlParts.sp_additional, re, ranked, '100%', YAHOO_CAPTION_RICH_SP_LIMIT_, 'sp_additional');
  return htmlParts;
}

/**
 * 内容量行を説明HTMLから除去（§0.1・ミス防止）
 */
function yahooCaptionRichStripContentQtyLines_(html) {
  var s = String(html || '');
  s = s.replace(/内容量：[^<]*(?:<br\s*\/?>)?\s*/g, '');
  s = s.replace(/<b>Net weight:<\/b>[^<]*(?:<br\s*\/?>)?\s*/gi, '');
  s = s.replace(/<b>Peso liquido:<\/b>[^<]*(?:<br\s*\/?>)?\s*/gi, '');
  s = s.replace(/<b>Khoi luong tinh:<\/b>[^<]*(?:<br\s*\/?>)?\s*/gi, '');
  s = s.replace(/<b>[^<]*<\/b>\s*1&#48393;\s*\d+\s*g(?:<br\s*\/?>)?\s*/gi, '');
  s = s.replace(/<b>[^<]*<\/b>\s*&#27599;&#34955;\d+&#20811;(?:<br\s*\/?>)?\s*/gi, '');
  s = s.replace(/<b>[^<]*<\/b>\s*\d+\s*&#3585;&#3619;&#3633;&#3617;[^<]*(?:<br\s*\/?>)?\s*/gi, '');
  return s;
}

/**
 * 原料1成分の表記ゆれを整える。
 * 「かつお：かつお（国産）」→「かつお（国産）」。「【かつお】かつお（国産）」→「かつお（国産）」。
 */
function yahooCaptionRichNormalizeMaterialPart_(part) {
  var p = String(part || '').trim();
  if (!p) return '';
  var mBr = p.match(/^【\s*([^】]+?)\s*】\s*(.*)$/);
  if (mBr) {
    var labB = String(mBr[1] || '').replace(/\s+/g, '');
    var bodyB = String(mBr[2] || '').trim();
    if (!bodyB) return labB;
    var headB = bodyB.replace(/\s+/g, '');
    if (labB && headB.indexOf(labB) === 0) return bodyB;
    return bodyB;
  }
  var m = p.match(/^([^：:]+)\s*[：:]\s*(.+)$/);
  if (m) {
    var lab = String(m[1] || '').replace(/\s+/g, '');
    var body = String(m[2] || '').trim();
    if (!body) return lab;
    var head = body.replace(/\s+/g, '');
    if (lab && head.indexOf(lab) === 0) return body;
    return body;
  }
  return p;
}

/**
 * 複数成分を /・／ または 【ラベル】 で分割。1件なら空配列（呼び出し側で単一部として扱う）。
 */
function yahooCaptionRichSplitMaterialParts_(materialJa) {
  var mat = String(materialJa || '').trim();
  if (!mat) return [];
  var parts = [];
  if (/[\/／]/.test(mat)) {
    var raw = mat.split(/\s*[\/／]\s*/);
    for (var i = 0; i < raw.length; i++) {
      var p = String(raw[i] || '').trim();
      if (p) parts.push(p);
    }
  } else if (/【[^】]+】/.test(mat)) {
    var re = /【[^】]+】[^【]*/g;
    var m;
    while ((m = re.exec(mat)) !== null) {
      var chunk = String(m[0] || '').trim();
      if (chunk) parts.push(chunk);
    }
  }
  if (parts.length <= 1) return [];
  return parts;
}

/**
 * ミックス等の複数成分は改行＋1行空け。ラベル重複は除去。カテゴリ専用の決め文句は置かない。
 */
function yahooCaptionRichFormatMaterialHtml_(materialJa) {
  var mat = String(materialJa || '').trim();
  if (!mat) return '';
  var parts = yahooCaptionRichSplitMaterialParts_(mat);
  if (!parts.length) {
    return yahooCaptionRichNormalizeMaterialPart_(mat) || mat;
  }
  var lines = [];
  for (var i = 0; i < parts.length; i++) {
    var line = yahooCaptionRichNormalizeMaterialPart_(parts[i]);
    if (line) lines.push(line);
  }
  if (!lines.length) return mat;
  if (lines.length === 1) return lines[0];
  return lines.join('<br><br>');
}

/**
 * 原料行（互換）。案Cでは yahooCaptionRichApplyProductAboutC_ を使う。
 * 既存の原料ブロックを消してから、保存：または※の直前へ差し込む。
 */
function yahooCaptionRichApplyMaterialLine_(html, materialJa, showMaterial) {
  var s = String(html || '');
  var mat = String(materialJa || '').trim();
  s = s.replace(/(?:主)?原料：(?:(?!保存：|※)[\s\S])*?(?=(?:保存：|※))/g, '');
  if (showMaterial && mat) {
    var matHtml = yahooCaptionRichFormatMaterialHtml_(mat);
    var matBlock = matHtml.indexOf('<br>') >= 0
      ? ('原料：<br>' + matHtml)
      : ('原料：' + matHtml);
    if (/保存：/.test(s)) {
      s = s.replace(/保存：/, function () {
        return matBlock + '<br>保存：';
      });
    } else if (/※/.test(s)) {
      s = s.replace(/※/, function () {
        return matBlock + '<br><br>※';
      });
    } else {
      s = matBlock + '<br>' + s;
    }
  }
  return s;
}

/**
 * 案C（2026-10-27）: 【商品について】枠・名称・保存を除去。※注記は残す。
 * 原料は showMaterial のときだけ（呼び出し側で caption のみ true）。
 */
function yahooCaptionRichApplyProductAboutC_(html, materialJa, showMaterial) {
  var s = String(html || '');
  s = s.replace(/<hr>\s*<b>【商品について】<\/b>(?:<br\s*\/?>)?\s*/gi, '');
  s = s.replace(/<b>【商品について】<\/b>(?:<br\s*\/?>)?\s*/gi, '');
  s = s.replace(/名称：[^<]*(?:<br\s*\/?>)?\s*/g, '');
  s = s.replace(/内容量：[^<]*(?:<br\s*\/?>)?\s*/g, '');
  s = s.replace(/保存：[^<]*(?:<br\s*\/?>)?\s*/g, '');
  // 既存の原料／主原料ブロックを除去（※注記の手前まで）
  s = s.replace(/(?:主)?原料：(?:(?!※)[\s\S])*?(?=(?:※|<a\s|$))/g, '');
  s = s.replace(/\n{3,}/g, '\n\n');
  var mat = String(materialJa || '').trim();
  if (showMaterial && mat) {
    var matHtml = yahooCaptionRichFormatMaterialHtml_(mat);
    var matBlock = matHtml.indexOf('<br>') >= 0
      ? ('原料：<br>' + matHtml)
      : ('原料：' + matHtml);
    if (/※/.test(s)) {
      s = s.replace(/※/, function () {
        return matBlock + '<br><br>※';
      });
    } else {
      s = s + '<br>' + matBlock + '<br>';
    }
  }
  return s;
}

/**
 * マスタ行が食品か（カテゴリー見出しに「カテゴリー」と「食品」、セルに「食品」）。
 */
function yahooCaptionRichIsFoodFromRow_(row) {
  if (!row) return false;
  var keys = Object.keys(row);
  for (var i = 0; i < keys.length; i++) {
    var k = String(keys[i] || '');
    if (k.indexOf('カテゴリー') >= 0 && k.indexOf('食品') >= 0) {
      return String(row[k] || '').indexOf('食品') >= 0;
    }
  }
  var cat = String(row['amazon カテゴリー'] || row['カテゴリー'] || '');
  return cat.indexOf('食品') >= 0;
}

/**
 * おだし骨格へ魚種・セットリンクを埋める（P5）。
 * 原料／【商品について】は BuildHtml 末尾の案C（captionのみ原料）で処理。
 */
function yahooCaptionRichApplyOdashiVars_(htmlParts, vars) {
  var fishJa = String(vars.fishJa || '').trim() || 'かつお';
  var fishEn = String(vars.fishEn || '').trim() || yahooCaptionRichFishEnFromJa_(fishJa);
  var keys = ['caption', 'additional1', 'additional2', 'sp_additional'];
  for (var ki = 0; ki < keys.length; ki++) {
    var k = keys[ki];
    var s = htmlParts[k];
    s = s.split('食べるおだし（かつお）').join('食べるおだし（' + fishJa + '）');
    s = s.split('国産かつおを').join('国産' + fishJa + 'を');
    s = yahooCaptionRichStripContentQtyLines_(s);
    s = s.split('Seasoned Bonito').join('Seasoned ' + fishEn);
    s = s.split('bonito from Japan').join(fishEn.toLowerCase() + ' from Japan');
    s = s.split('Bonito japones').join(fishEn + ' japones');
    s = s.split('bonito do Japao').join(fishEn.toLowerCase() + ' do Japao');
    s = s.split('Tinimplahang bonito').join('Tinimplahang ' + fishEn.toLowerCase());
    s = s.split('bonito mula Japan').join(fishEn.toLowerCase() + ' mula Japan');
    // 兄弟が取れても取れなくても置換する（骨格の単品JANリンクを残さない）
    s = yahooCaptionRichReplaceSetsBlock_(
      s, vars.setItems || [], k === 'sp_additional' ? '100%' : '600');
    htmlParts[k] = s;
  }
  return htmlParts;
}

function yahooCaptionRichFishEnFromJa_(fishJa) {
  var s = String(fishJa || '');
  var parts = s.split(/[・･\/／,、]/).map(function (p) { return p.replace(/\s+/g, ''); }).filter(Boolean);
  var map = { 'かつお': 'Bonito', 'まぐろ': 'Tuna', 'ぶり': 'Yellowtail' };
  var out = [];
  for (var i = 0; i < parts.length; i++) {
    out.push(map[parts[i]] || parts[i]);
  }
  if (!out.length) return 'Bonito';
  if (out.length === 1) return out[0];
  if (out.length === 2) return out[0] + ' and ' + out[1];
  return out.slice(0, -1).join(', ') + ', and ' + out[out.length - 1];
}

function yahooCaptionRichBuildSetsTable_(setItems, tableW) {
  var items = (setItems || []).slice().sort(function (a, b) { return a.bags - b.bags; });
  var rows = [];
  for (var i = 0; i < items.length; i += 3) {
    var cells = [];
    for (var j = 0; j < 3; j++) {
      if (j > 0) cells.push('<td width="2%">&nbsp;</td>');
      if (i + j < items.length) {
        var it = items[i + j];
        var unitLabel = (it.unit === '個') ? '個セット' : '袋セット';
        cells.push(
          '<td width="32%" bgcolor="#f5efe4" align="center">' +
          '<a href="' + YAHOO_CAPTION_RICH_STORE_ITEM_BASE_ + '/' + it.code + '.html">' +
          '<b><br>' + it.bags + unitLabel + '<br>　</b></a></td>'
        );
      } else {
        cells.push('<td width="32%" align="center">&nbsp;</td>');
      }
    }
    rows.push('<tr>\n' + cells.join('\n') + '\n</tr>');
    if (i + 3 < items.length) rows.push('<tr><td colspan="5" height="8"></td></tr>');
  }
  return (
    '<table width="' + tableW + '" border="0" cellpadding="8" cellspacing="0">\n' +
    rows.join('\n') +
    '\n</table>'
  );
}

function yahooCaptionRichReplaceSetsBlock_(html, setItems, tableW) {
  var items = setItems || [];
  var block;
  if (!items.length) {
    // 骨格の単品JANリンクを残さない（空のときはコーナー自体を外す）
    block = '<a name="sets"></a>\n';
  } else {
    var table = yahooCaptionRichBuildSetsTable_(items, tableW);
    block =
      '<a name="sets"></a>\n<hr>\n<b>【セットコーナー】</b><br>\n' +
      'セットを選ぶと、各セットの商品ページへ移動します。<br>\n<br>\n' +
      table + '\n';
  }
  // 案C後も置換できるよう、【商品について】以外の終端も見る
  var re = /<a name="sets"><\/a>[\s\S]*?(?=(?:<hr>\s*<b>【商品について】<\/b>)|(?:原料：)|(?:※原材料)|(?:<a name="international")|(?:<hr>\s*<b>【For international))/;
  if (!re.test(html)) return html;
  return html.replace(re, block);
}

/**
 * 袋／個のセット数。ミックスは「2個/2種」「3袋/3種」など。先頭の数量を取る。
 * @return {{n:number, unit:string}}
 */
function yahooCaptionRichParseSetCount_(text) {
  var s = String(text || '');
  var m = s.match(/(\d+)\s*袋/);
  if (m) return { n: parseInt(m[1], 10), unit: '袋' };
  m = s.match(/(\d+)\s*個で\s*1\s*セット/);
  if (m) return { n: parseInt(m[1], 10), unit: '個' };
  m = s.match(/(\d+)\s*個\s*[\/／]/);
  if (m) return { n: parseInt(m[1], 10), unit: '個' };
  m = s.match(/(\d+)\s*個/);
  if (m) return { n: parseInt(m[1], 10), unit: '個' };
  return { n: 0, unit: '' };
}

/** @deprecated ParseSetCount_ を使う */
function yahooCaptionRichParseBagCount_(text) {
  return yahooCaptionRichParseSetCount_(text).n;
}

function yahooCaptionRichBuildSetItemsFromSiblings_(siblings) {
  var byBags = {};
  (siblings || []).forEach(function (sib) {
    var code = String(sib && sib.code || '').trim();
    if (!code) return;
    var parsed = yahooCaptionRichParseSetCount_(sib.variation);
    if (!parsed.n) parsed = yahooCaptionRichParseSetCount_(sib.requestName);
    if (parsed.n < 2) return;
    if (!byBags[parsed.n]) {
      byBags[parsed.n] = { bags: parsed.n, code: code, unit: parsed.unit || '袋' };
    }
  });
  return Object.keys(byBags).map(Number).sort(function (a, b) { return a - b; }).map(function (b) {
    return byBags[b];
  });
}

/**
 * 商品名からおだしの括弧内（魚種ラベル）を取る
 */
function yahooCaptionRichExtractOdashiFishJa_(text) {
  var s = String(text || '');
  var m = s.match(/食べるおだし\s*[（(]\s*([^）)]+?)[）)]/);
  if (!m) return '';
  return m[1].replace(/\s+/g, '').replace(/[･]/g, '・');
}

function yahooCaptionRichExtractNetWeight_(text, fallbackCell) {
  var s = String(text || '');
  var m = s.match(/1\s*袋\s*[/／]\s*(\d+)\s*[gｇG]/);
  if (m) return { netWeight: m[1] + 'g', netWeightNum: m[1] };
  m = s.match(/(\d+)\s*[gｇG]\s*$/);
  if (m && /食べるおだし/.test(s)) return { netWeight: m[1] + 'g', netWeightNum: m[1] };
  var cell = String(fallbackCell || '').trim();
  var cm = cell.match(/(\d+)/);
  if (cm) return { netWeight: cm[1] + 'g', netWeightNum: cm[1] };
  return { netWeight: '', netWeightNum: '' };
}

function yahooCaptionRichLooksLikeOdashiProduct_(row, product) {
  var jan = String((row && (row['JANコード'] || row['JAN'] || row['jan'])) || '');
  var code = String((product && product.code) || (row && row['子SKU']) || '');
  var name = String(
    (row && (row['商品名ベース'] || row['依頼日'] || row['商品名Yahoo!'] || row['商品名'])) ||
    (product && product.name) || ''
  );
  if (/チーズかつお/.test(name)) return false;
  for (var i = 0; i < YAHOO_CAPTION_ODASHI_JANS_.length; i++) {
    var j = YAHOO_CAPTION_ODASHI_JANS_[i];
    if (jan.indexOf(j) !== -1 || code.indexOf(j) !== -1) return true;
  }
  if (/食べるおだし\s*[（(]/.test(name)) return true;
  return false;
}

/**
 * 本採用サブ画像を差し替え。
 * vars.imgUrl1..3 があれば絶対URL優先。なければ img1..3 を lib ファイル名として使う。
 */
function yahooCaptionRichApplyBodyImages_(htmlParts, vars, defaultNames) {
  vars = vars || {};
  var toUrls = [];
  for (var i = 0; i < defaultNames.length; i++) {
    var full = String(vars['imgUrl' + (i + 1)] || '').trim();
    if (full) {
      toUrls.push(full);
      continue;
    }
    var name = String(vars['img' + (i + 1)] || '').trim() || defaultNames[i];
    if (/^https?:\/\//i.test(name)) {
      toUrls.push(name);
    } else {
      toUrls.push(YAHOO_CAPTION_RICH_LIB_BASE + name);
    }
  }
  var keys = ['caption', 'additional1', 'additional2', 'sp_additional'];
  for (var ki = 0; ki < keys.length; ki++) {
    var k = keys[ki];
    var s = htmlParts[k];
    for (var j = 0; j < defaultNames.length; j++) {
      var fromUrl = YAHOO_CAPTION_RICH_LIB_BASE + defaultNames[j];
      if (fromUrl !== toUrls[j]) {
        s = s.split(fromUrl).join(toUrls[j]);
      }
    }
    htmlParts[k] = s;
  }
  return htmlParts;
}

function yahooCaptionRichP6ImgEnabled_() {
  if (typeof getBoolScriptProperty_ === 'function') {
    return getBoolScriptProperty_(YAHOO_CAPTION_RICH_P6_IMG_PROP_, true);
  }
  return true;
}

/**
 * Yahoo説明HTMLに載せてよい画像URLか（他モールCDNはプレビューで欠ける実測あり）
 */
function yahooCaptionRichIsYahooSafeImgUrl_(url) {
  var u = String(url || '').trim();
  if (!u) return false;
  if (!/^https?:\/\//i.test(u)) return false;
  // 追加画像 lib / 商品画像CDN 等
  return /(^|\.)yimg\.jp\//i.test(u) || /shopping\.yahoo\.co\.jp\//i.test(u);
}

/**
 * 相対パスを絶対URLに（Yahoo安全URLのみ後段で採用）
 */
function yahooCaptionRichAbsImageUrl_(url) {
  var u = String(url || '').trim();
  if (!u || u === 'undefined') return '';
  if (/^https?:\/\//i.test(u)) return u;
  if (u.charAt(0) === '/') return YAHOO_CAPTION_RICH_RAKUTEN_CABINET_ + u;
  return u;
}

/**
 * 魚種ラベル → lib ファイル用スラッグ（先頭1種）
 */
function yahooCaptionRichFishSlugFromJa_(fishJa) {
  var s = String(fishJa || '').replace(/\s+/g, '');
  var first = s.split(/[・･\/／]/)[0] || '';
  var map = { 'かつお': 'katsuo', 'まぐろ': 'maguro', 'ぶり': 'buri' };
  return map[first] || 'katsuo';
}

/**
 * おだし本文3枚の lib ファイル名
 */
function yahooCaptionRichOdashiLibBodyNames_(slug) {
  var s = String(slug || 'katsuo');
  return [
    s + '_body_t03_sizzle.jpg',
    s + '_body_t04_aji.jpg',
    s + '_body_t11_arrange.jpg'
  ];
}

/** 実行中キャッシュ: slug → lib t03 が 200 か */
var YAHOO_CAPTION_RICH_LIB_SLUG_OK_ = {};

/**
 * 追加画像libに魚種3枚がUP済みか（t03で代表チェック）
 */
function yahooCaptionRichLibSlugReady_(slug) {
  var s = String(slug || '');
  if (!s) return false;
  if (s === 'katsuo') return true; // 本採用済
  if (YAHOO_CAPTION_RICH_LIB_SLUG_OK_[s] !== undefined) return YAHOO_CAPTION_RICH_LIB_SLUG_OK_[s];
  var ok = false;
  try {
    var url = YAHOO_CAPTION_RICH_LIB_BASE + yahooCaptionRichOdashiLibBodyNames_(s)[0];
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    ok = res.getResponseCode() === 200;
  } catch (e) {
    ok = false;
  }
  YAHOO_CAPTION_RICH_LIB_SLUG_OK_[s] = ok;
  console.log('[YAHOO_CAPTION_RICH] lib_slug_check slug=' + s + ' ready=' + ok);
  return ok;
}

/**
 * P6: 説明画像列 → 魚種lib（UP済） → かつおlib。楽天CDN直載せはしない。
 * @param {function(string):string} cellFn
 * @param {string} [fishJa]
 */
function yahooCaptionRichResolveBodyImages_(cellFn, fishJa) {
  var col1 = cellFn(YAHOO_CAPTION_IMAGE1_COL);
  var col2 = cellFn(YAHOO_CAPTION_IMAGE2_COL);
  var col3 = cellFn(YAHOO_CAPTION_IMAGE3_COL);
  if (col1 || col2 || col3) {
    var u1 = /^https?:\/\//i.test(col1) ? col1 : '';
    var u2 = /^https?:\/\//i.test(col2) ? col2 : '';
    var u3 = /^https?:\/\//i.test(col3) ? col3 : '';
    if (u1 || u2 || u3) {
      var safe1 = yahooCaptionRichIsYahooSafeImgUrl_(u1) ? u1 : '';
      var safe2 = yahooCaptionRichIsYahooSafeImgUrl_(u2) ? u2 : '';
      var safe3 = yahooCaptionRichIsYahooSafeImgUrl_(u3) ? u3 : '';
      if (safe1 || safe2 || safe3) {
        return {
          img1: '', img2: '', img3: '',
          imgUrl1: safe1, imgUrl2: safe2, imgUrl3: safe3,
          imgSrc: 'master_col_url'
        };
      }
    }
    return {
      img1: u1 ? '' : col1,
      img2: u2 ? '' : col2,
      img3: u3 ? '' : col3,
      imgUrl1: '', imgUrl2: '', imgUrl3: '',
      imgSrc: 'master_col'
    };
  }
  if (!yahooCaptionRichP6ImgEnabled_()) {
    return { img1: '', img2: '', img3: '', imgUrl1: '', imgUrl2: '', imgUrl3: '', imgSrc: 'katsuo_lib_default' };
  }

  // P6b: 魚種専用 lib（Pro追加画像UP済みなら自動採用）
  var slug = yahooCaptionRichFishSlugFromJa_(fishJa || 'かつお');
  if (slug !== 'katsuo' && yahooCaptionRichLibSlugReady_(slug)) {
    var names = yahooCaptionRichOdashiLibBodyNames_(slug);
    return {
      img1: names[0],
      img2: names[1],
      img3: names[2],
      imgUrl1: '', imgUrl2: '', imgUrl3: '',
      imgSrc: 'fish_lib_' + slug
    };
  }
  if (slug !== 'katsuo') {
    var subN = 0;
    if (cellFn('楽天サブ画像1')) subN++;
    if (cellFn('楽天サブ画像2')) subN++;
    if (cellFn('楽天サブ画像3')) subN++;
    if (subN) {
      console.log('[YAHOO_CAPTION_RICH] img fish_lib_pending slug=' + slug +
        ' (追加画像未UP。かつおlib暫定。楽天サブ直載せはしない)');
    }
  }
  return { img1: '', img2: '', img3: '', imgUrl1: '', imgUrl2: '', imgUrl3: '', imgSrc: 'katsuo_lib_default' };
}

/**
 * 親行からテンプレIDと変数を読む（子は親を継承）
 */
function yahooCaptionRichReadVarsFromProduct_(product) {
  // originalRow は親強制マージ済み（テンプレ列含む）
  var row = (product && product.originalRow) || {};
  function cell(name) {
    if (row[name] !== undefined && row[name] !== null && String(row[name]).trim() !== '') {
      return String(row[name]).trim();
    }
    return '';
  }
  var templateId = cell(YAHOO_CAPTION_TEMPLATE_ID_COL);
  if (templateId === 'odashi_gourmet') templateId = 'katsuo_odashi';
  // P5: 列空でもおだし系なら共通骨格を解決（手製第2ID不要）
  if (!templateId && yahooCaptionRichLooksLikeOdashiProduct_(row, product)) {
    templateId = 'katsuo_odashi';
  }
  // 初号フォールバック（互換）
  if (!templateId) {
    var jan = cell('JANコード') || cell('JAN') || cell('jan');
    var code = String((product && product.code) || cell('子SKU') || '');
    if (jan.indexOf('4906283045119') !== -1 || code.indexOf('4906283045119') !== -1) {
      templateId = 'katsuo_odashi';
    }
  }

  var nameBlob = [cell('商品名ベース'), cell('依頼日'), cell('商品名Yahoo!'), cell('商品名'), product && product.name]
    .filter(Boolean).join(' ');
  var fishJa = yahooCaptionRichExtractOdashiFishJa_(nameBlob) || 'かつお';
  var setItems = yahooCaptionRichBuildSetItemsFromSiblings_(product && product.captionSiblings);
  var isFood = yahooCaptionRichIsFoodFromRow_(row);
  var materialJa = cell('原料');
  var showMaterial = isFood && !!materialJa;
  var imgs = yahooCaptionRichResolveBodyImages_(cell, fishJa);
  var jan = cell('JANコード') || cell('JAN') || cell('jan') || '';
  var parentSku = cell('親SKU') || '';
  // P6cパックの親別slug（同一JANの誤当て防止）。未設定時は従来の魚種slug。
  var slug = '';
  if (parentSku.indexOf('-oya3') !== -1 && jan.indexOf('4906283045119') !== -1) {
    slug = '5119_oya3';
  } else if (parentSku.indexOf('-oya2') !== -1 && jan.indexOf('4906283045119') !== -1) {
    slug = '5119_oya2';
  } else {
    slug = yahooCaptionRichFishSlugFromJa_(fishJa);
  }

  console.log('[YAHOO_CAPTION_RICH] vars tmpl=' + templateId +
    ' fish=' + fishJa +
    ' food=' + isFood +
    ' material=' + (showMaterial ? materialJa : '(hidden)') +
    ' net=stripped' +
    ' sets=' + setItems.length +
    ' imgSrc=' + imgs.imgSrc +
    ' p6c=' + yahooCaptionRichP6cEnabled_() +
    ' parent=' + parentSku +
    ' item_code=' + String((product && product.code) || cell('子SKU') || ''));

  return {
    templateId: templateId,
    jan: jan,
    parentSku: parentSku,
    slug: slug,
    img1: imgs.img1,
    img2: imgs.img2,
    img3: imgs.img3,
    imgUrl1: imgs.imgUrl1,
    imgUrl2: imgs.imgUrl2,
    imgUrl3: imgs.imgUrl3,
    imgSrc: imgs.imgSrc,
    fishJa: fishJa,
    fishEn: yahooCaptionRichFishEnFromJa_(fishJa),
    netWeight: '',
    netWeightNum: '',
    materialJa: materialJa,
    showMaterial: showMaterial,
    isFood: isFood,
    setItems: setItems
  };
}

/**
 * editItem: 装飾4欄＋Yahoo必須項目（path/name/product_category/price）
 * ※ caption のみ送ると it-01011 等で 400 になる（2026-09-14 実測）
 * @param {YahooApiClient} client
 * @param {Object} product YahooDataBuilder の商品
 * @param {{caption: string, additional1: string, additional2: string, sp_additional: string}} htmlParts
 */
function yahooCaptionRichEditItemOnly_(client, product, htmlParts) {
  var itemCode = product.code || (product.originalRow && product.originalRow['子SKU']) || '';
  var row = product.originalRow || {};
  var name = client._getValue('name', product) || product.name || row['商品名Yahoo!'] || row['商品名'] || '';
  var path = client._getValue('path', product);
  var productCategory = client._getValue('product-category', product);
  if (!productCategory) productCategory = client._getValue('product_category', product);
  var price = client._getValue('price', product);
  if (price === '' || price === null || price === undefined) {
    price = row['販売価格'] || row['商品の販売価格'] || row['Yahoo!価格設定'] || '';
  }
  var payload = {
    item_code: itemCode,
    name: name,
    path: path,
    product_category: productCategory,
    price: price,
    caption: htmlParts.caption,
    additional1: htmlParts.additional1,
    additional2: htmlParts.additional2,
    sp_additional: htmlParts.sp_additional
  };
  // 扉OFF時の最小投入でも、マッピングにあれば他欄を落としにくくする
  var keepKeys = ['headline', 'abstract', 'explanation', 'brand_code', 'jan', 'delivery', 'condition', 'taxable'];
  for (var ki = 0; ki < keepKeys.length; ki++) {
    var kk = keepKeys[ki];
    var kv = client._getValue(kk, product);
    if (kk === 'taxable' && kv !== '' && kv != null) {
      payload.taxable = String(client._coerceTaxable(kv));
      continue;
    }
    if (kv !== '' && kv != null && kv !== undefined) payload[kk] = kv;
  }
  console.log('[YAHOO_CAPTION_RICH] editItem item_code=' + itemCode +
    ' name_len=' + String(name).length +
    ' path=' + String(path).substring(0, 40) +
    ' cat=' + String(productCategory).substring(0, 20) +
    ' price=' + price +
    ' caption_bytes=' + yahooCaptionRichByteLen_(htmlParts.caption) +
    ' a1=' + yahooCaptionRichByteLen_(htmlParts.additional1) +
    ' a2=' + yahooCaptionRichByteLen_(htmlParts.additional2) +
    ' sp=' + yahooCaptionRichByteLen_(htmlParts.sp_additional));
  if (!name || !path || !productCategory || price === '' || price === null || price === undefined) {
    throw new Error('必須不足 item_code=' + itemCode +
      ' name=' + (!!name) + ' path=' + (!!path) + ' product_category=' + (!!productCategory) + ' price=' + (price !== '' && price != null));
  }
  return client._post(YAHOO_ITEM_API_URL, payload);
}

function yahooCaptionRichByteLen_(s) {
  try {
    return Utilities.newBlob(String(s || ''), 'text/html', 'x').getBytes().length;
  } catch (e) {
    return String(s || '').length;
  }
}

/**
 * ed-00006（反映／アップロード中）等の一時失敗か
 */
function yahooCaptionRichIsTransientEditError_(err) {
  var s = String(err && err.message ? err.message : err || '');
  return /ed-00006|反映また|アップロード中|更新ができません/.test(s);
}

/**
 * editItem→submitItem。ed-00006 は待機リトライ。
 * Property ON 時は updateItem（マッピング全項目＋テンプレ注入）を使い、
 * headline/abstract/explanation 等の省略＝空上書きを防ぐ。
 */
function yahooCaptionRichApplyOneProduct_(client, product, htmlParts, tmplId, runId) {
  var itemCode = product.code || (product.originalRow && product.originalRow['子SKU']) || '';
  var maxAttempts = 4;
  var lastErr = null;
  for (var attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      if (yahooCaptionRichLockEnabled_()) {
        client.updateItem(product);
      } else {
        yahooCaptionRichEditItemOnly_(client, product, htmlParts);
      }
      logYahooEvent('API', 'captionRich.editItem', itemCode, 'OK tmpl=' + tmplId + (attempt > 1 ? ' retry=' + attempt : ''));
      Utilities.sleep(1200);
      var submitRes = client.submitItem(product);
      if (!submitRes || !submitRes.ok) {
        var body = (submitRes && submitRes.body) || '';
        if (/it-07004|新規ページは個別反映できません/.test(body)) {
          logYahooEvent('API', 'captionRich.submitItem', itemCode, '手動反映が必要');
        } else if (/ed-00006|反映また|アップロード中/.test(body)) {
          throw new Error('submitItem ed-00006: ' + body.substring(0, 120));
        } else {
          throw new Error('submitItem HTTP ' + ((submitRes && submitRes.code) || '') + ' ' + String(body).substring(0, 160));
        }
      } else {
        logYahooEvent('API', 'captionRich.submitItem', itemCode, 'OK');
      }
      console.log('[YAHOO_CAPTION_RICH] OK item_code=' + itemCode + ' tmpl=' + tmplId + ' runId=' + runId +
        (attempt > 1 ? ' attempt=' + attempt : ''));
      return true;
    } catch (e) {
      lastErr = e;
      if (yahooCaptionRichIsTransientEditError_(e) && attempt < maxAttempts) {
        var waitMs = 3000 * attempt;
        console.warn('[YAHOO_CAPTION_RICH] transient item_code=' + itemCode +
          ' attempt=' + attempt + '/' + maxAttempts + ' waitMs=' + waitMs + ' ' + String(e && e.message).substring(0, 120));
        Utilities.sleep(waitMs);
        continue;
      }
      throw e;
    }
  }
  throw lastErr || new Error('apply failed');
}

/**
 * メニュー: 出品CK行にテンプレHTMLをAPI投入→submitItem
 * ※開始時の確認ダイアログは出さない（都度OKが手間なため）。完了はトーストのみ。
 */
function menuYahooCaptionRichApply() {
  var ui = SpreadsheetApp.getUi();
  var ans = ui.alert(
    'Yahoo説明の反映',
    'レ点の商品の説明をYahooへ送ります。\nキャンセルなら送りません。',
    ui.ButtonSet.OK_CANCEL
  );
  if (ans !== ui.Button.OK) {
    Logger.log('[menuYahooCaptionRichApply] state=CANCEL');
    return;
  }
  Logger.log('[menuYahooCaptionRichApply] state=OK');
  runYahooCaptionRichApply_();
}

/**
 * 本処理（メニュー／検証から呼び出し可）
 */
function runYahooCaptionRichApply_() {
  console.log('=== Yahoo説明 API反映 開始 ===');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName(SHEET_NAME_MASTER);
  var mapSheet = ss.getSheetByName(SHEET_NAME_YAHOO_MAP);
  if (!masterSheet || !mapSheet) throw new Error('必要なシートが見つかりません。');

  var builder = new YahooDataBuilder(masterSheet, mapSheet);
  var products = builder.buildProductData();
  if (!products.length) {
    SpreadsheetApp.getUi().alert('出品CK付きの商品がありません。');
    return;
  }

  var accessToken = getYahooAccessToken(mapSheet);
  var sellerId = mapSheet.getRange('B15').getValue();
  var client = new YahooApiClient(accessToken, sellerId, mapSheet);

  var ok = 0;
  var ng = 0;
  var skip = 0;
  var runId = 'ycr_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');

  for (var i = 0; i < products.length; i++) {
    var product = products[i];
    var itemCode = product.code || (product.originalRow && product.originalRow['子SKU']) || '';
    var vars = yahooCaptionRichReadVarsFromProduct_(product);
    if (!vars.templateId) {
      skip++;
      console.log('[YAHOO_CAPTION_RICH] skip no_template item_code=' + itemCode + ' runId=' + runId);
      continue;
    }
    var htmlParts = yahooCaptionRichBuildHtml_(vars.templateId, vars);
    if (!htmlParts) {
      skip++;
      console.log('[YAHOO_CAPTION_RICH] skip unknown_template id=' + vars.templateId + ' item_code=' + itemCode + ' runId=' + runId);
      continue;
    }
    try {
      yahooCaptionRichApplyOneProduct_(client, product, htmlParts, vars.templateId, runId);
      ok++;
    } catch (e) {
      ng++;
      console.error('[YAHOO_CAPTION_RICH] FAIL item_code=' + itemCode + ' ' + (e && e.message));
      logYahooEvent('API', 'captionRich.editItem', itemCode, 'FAIL ' + String(e && e.message || e).substring(0, 160));
    }
    // 連続 submitItem の反映ロック（ed-00006）回避
    if (i < products.length - 1) Utilities.sleep(2000);
  }

  logYahooEvent('SUCCESS', 'BATCH', 'Yahoo説明API反映', '成功:' + ok + ', 失敗:' + ng + ', スキップ:' + skip + ', runId=' + runId);
  console.log('=== Yahoo説明 API反映 終了 ok=' + ok + ' ng=' + ng + ' skip=' + skip + ' ===');
  var msg = '成功:' + ok + ' 失敗:' + ng + ' スキップ:' + skip + ' runId:' + runId;
  try {
    ss.toast(msg, 'Yahoo説明 API反映', 8);
  } catch (te) {
    // toast不可時のみダイアログ
    SpreadsheetApp.getUi().alert('Yahoo説明 API反映', msg.replace(/ /g, '\n'), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
