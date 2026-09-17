# -*- coding: utf-8 -*-
"""再訪不要案件の要因分類ロジック（備考テキスト＋結果区分）"""
import re
import pandas as pd

# 優先順に評価。最初に一致した要因を採用する。
RULES = [
    ("他社受託：システム抱き合わせ",
     r"行政研|ぎょうせい|TKC|MJS|アチカ|オーレンス|システム.*(保守|範囲|使用|利用)|保守(料|範囲)"),
    ("他社受託：入札実績要件で参入不可",
     r"実績要件"),
    ("他社受託：随契・委託継続",
     r"他社|随契|随意契約|委託済|へ委託|に委託|にて委託|委託にて|日本総研|辻本郷|建設コンサル|冨士商事|サーベイ|２１世紀|21世紀|三菱UFJ|アドバイザー"),
    ("予算がつかない・否決",
     r"予算(が)?(つかず|付かず|否決|取れ)|予算否決|予算もらい|付きづらく"),
    ("他計画へ一体化・包含され単独案件消滅",
     r"一体(で)?(策定|作成)|内包|統合|包含|として一体"),
    ("策定済み・改定時期が先",
     r"策定済|作成済|改訂?したばかり|R1[01]|計画終了|期限切れ|中間見直し|改定時期に提案|R9|R10"),
    ("必要性・意欲が低い",
     r"熱心でない|罰則|メリット(を)?感じ|必要性を感じ|考えていな|予定はな|予定な(し|い)|実施しない|温度感低|意思(も)?な(し|い)|切替の意思"),
    ("自前作成方針（庁内対応）",
     r"自前|自力|自庁|長年|担当者経歴|置き換え|自作"),
    ("庁内事情・外部要因による遅延",
     r"負担|時間を要|調整|めど"),
    ("担当不在・接触不可",
     r"不在|会えな"),
]


def classify(note: str) -> str:
    """備考テキストから要因を1つ返す。該当なしは空文字。"""
    if not isinstance(note, str) or not note.strip():
        return ""
    for label, pattern in RULES:
        if re.search(pattern, note):
            return label
    return ""


def classify_row(note: str, result: str) -> str:
    """備考で分類できない場合は結果区分から補完する。"""
    hit = classify(note)
    if hit:
        return hit
    fallback = {
        "自力": "自前作成方針（庁内対応）",
        "他社随契": "他社受託：随契・委託継続",
        "不在": "担当不在・接触不可",
    }
    return fallback.get(result, "要因不明（備考・結果とも未記入）")


def add_factors(df: pd.DataFrame) -> pd.DataFrame:
    """再訪不要行に要因列を付与した DataFrame を返す。"""
    out = df.copy()
    notes = out["備考"].fillna("").astype(str).str.replace("\n", "／", regex=False)
    out["要因"] = [
        classify_row(nt, rs) if act == "再訪不要" else ""
        for nt, rs, act in zip(notes, out["結果"], out["次回アクション"])
    ]
    return out
