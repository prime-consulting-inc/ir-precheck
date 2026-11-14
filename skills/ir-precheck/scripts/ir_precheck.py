#!/usr/bin/env python3
"""IR開示資料の提出前チェッカー。

PDF / Word / PowerPoint / Excel を検査し、プロパティ・コメント・
プレースホルダ・黒塗りの残存、伏せ表記の下に残ったテキスト、覆われた画像、社外秘マーク、レビュー文言、作成日時などを日本語で報告する。
検査対象ファイルは変更しない（--clean-metadata は公開用のコピーを新規作成）。
外部との通信はしない。黒塗り・伏せ字・白文字などの下にある文字列は出力しない（場所と件数だけ）。

終了コード: 0 指摘なしまたは情報のみ / 1 要確認あり / 2 要修正あり / 3 検査できなかったファイルや引数の誤りあり
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import unicodedata
import zipfile
from dataclasses import asdict, dataclass, field
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET

try:
    from zoneinfo import ZoneInfo
    JST = ZoneInfo("Asia/Tokyo")
except Exception:
    JST = timezone(timedelta(hours=9))

__version__ = "1.0.0"

SEVERITY_ORDER = {"未検査": 0, "要修正": 1, "要確認": 2, "情報": 3}
SEVERITIES = ("未検査", "要修正", "要確認", "情報")

INSTALL_HINT = "pip install -r requirements.txt（スキルのフォルダで実行）"

# プロパティの画面上の名前
FIELD_LABELS = {
    "Title": "タイトル（Title）",
    "Author": "作成者（Author）",
    "Subject": "サブタイトル（Subject）",
    "Keywords": "キーワード（Keywords）",
    "Creator": "作成アプリケーション（Creator）",
    "Producer": "PDF変換（Producer）",
    "CreationDate": "作成日時（CreationDate）",
    "ModDate": "更新日時（ModDate）",
    "LastModifiedBy": "最終更新者（LastModifiedBy）",
}


def _lab(key: str) -> str:
    return FIELD_LABELS.get(key, key)


def _nfkc(s) -> str:
    """日本語PDFの抽出で出る康熙部首（⽉⽇⾦）や互換漢字、全角英数をそろえてから照合する。"""
    if not s:
        return ""
    return unicodedata.normalize("NFKC", str(s))

# 「議案」「提案」など通常の開示用語を除外した要注意語
NG_PATTERNS: list[tuple[str, re.Pattern]] = [
    ("内部稿・案", re.compile(r"（案）|\(案\)|【案】|開示文ドラフト|チェック依頼|IR案|開示案|お知らせ案|説明会?資料案")),
    ("ドラフト", re.compile(r"ドラフト|(?<![A-Za-z])[Dd]raft(?![A-Za-z])")),
    ("バージョン痕跡", re.compile(r"(?i)(?<![A-Za-z0-9])v\d{1,2}(?![A-Za-z0-9._-])|_v\d+|v\d+_clean|\bclean\b|final\d")),
    ("Word変換名", re.compile(r"Microsoft Word\s*-")),
    ("プレースホルダ", re.compile(r"\bTODO\b|\bFIXME\b|\bTBD\b|ＴＢＤ|仮置き|要確認事項")),
    ("伏字プレースホルダ", re.compile(r"○○株式会社|〇〇株式会社|ＸＸ株式会社|XX株式会社|××株式会社")),
    ("リーク等の社内メモ", re.compile(r"リークプレス|リーク対応|提出断念|リスクリリース|週明け報道")),
]

WEEKDAYS = "月火水木金土日"
WEEKDAY_RE = re.compile(
    r"(?P<y>20\d{2}|令和\s*[元0-9０-９]+)\s*年\s*"
    r"(?P<m>\d{1,2})\s*月\s*(?P<d>\d{1,2})\s*日\s*[（(]\s*(?P<w>[月火水木金土日])"
)
WAREKI_OFFSET = 2018  # 令和元年 = 2019
CODE_RE = re.compile(r"(?<![A-Z0-9])(\d{4}|[0-9]{3}[A-Z]|[0-9]{4}[A-Z])(?![A-Z0-9])")
COMPANY_RE = re.compile(
    r"(?:株式会社[\s\u3000]*([一-龥ァ-ンーA-Za-z0-9Ａ-Ｚａ-ｚ０-９・]{2,20})|"
    r"([一-龥ァ-ンーA-Za-z0-9Ａ-Ｚａ-ｚ０-９・]{2,20})[\s\u3000]*株式会社)"
)
_COMPANY_STOP = {"会社名", "名称", "商号", "代表者", "代表者名", "問合せ先", "問い合わせ先", "住所", "氏名", "各位", "受注者", "発注者", "委託者"}

# グラフ等で競合名を伏せるときのラベル
ANON_LABEL_RE = re.compile(r"^[A-Ea-eＡ-Ｅ]社$|^[○●□■×ＸX某甲乙]社$|^〇{1,2}社$|^○{1,2}社$|^〇〇社$")


# 社外秘など。通常の開示文に出にくいものは要修正。
SECRET_STRONG = [
    ("社外秘", re.compile(r"社外秘")),
    ("取扱注意", re.compile(r"取扱注意")),
    ("関係者限り", re.compile(r"関係者限り")),
    ("社内限り", re.compile(r"社内限り")),
    ("内部限", re.compile(r"内部限(?:り)?")),
    ("Internal Use", re.compile(r"(?i)for\s+internal\s+use(?:\s+only)?|internal\s+use\s+only")),
    ("STRICTLY CONFIDENTIAL", re.compile(r"(?i)strictly\s+confidential")),
    ("開示委員会資料", re.compile(r"開示委員会資料")),
]
CONFIDENTIAL_RE = re.compile(r"(?i)(?<![A-Za-z])confidential(?![A-Za-z])")
CONFIDENTIAL_IN_SENTENCE = re.compile(
    r"(?i)(confidential\s+information|non-?confidential|confidentiality)"
)

def _not_officer_label(text: str, m: re.Match) -> bool:
    """役員一覧の「取締役（社内）」「監査等委員（社内）」は除外する。"""
    before = text[max(0, m.start() - 4): m.start()]
    return not re.search(r"[役員]\s*$", before)


# (表示名, パターン, 追加の判定。Falseなら指摘しない)
REVIEW_STRONG = [
    ("差し替え指示", re.compile(r"差し替えて(?:ください|下さい)|差し替え(?:お願い|願います)"), None),
    ("数字は後で", re.compile(r"数字(?:は|を)?後で|後で(?:数字|入れる|入れます|記載)"), None),
    ("修正指示", re.compile(r"※\s*修正|※\s*要修正|ここを直(?:して)?|直してください|直して下さい"), None),
    ("消し忘れ言及", re.compile(r"消し忘れて"), None),
    ("社内向け", re.compile(r"\(社内\)|【社内】"), _not_officer_label),
]
REVIEW_SOFT = [
    ("要確認の残", re.compile(r"※\s*要確認"), None),
    ("確認中メモ", re.compile(r"(?:数字は|まだ|※)\s*確認中"), None),
    ("統一する指示", re.compile(r"(?:表記|表現|言い回し|用語).{0,10}統一する|統一する.{0,10}(?:表記|表現)"), None),
    ("謝罪の文", re.compile(r"申し訳(?:ございません|ありません)"), None),
]
# 訂正・お詫びの開示では謝罪の文は通常の表現
APOLOGY_CONTEXT = re.compile(r"訂正|お詫び|おわび")

# 本文で「要修正」にする区分（それ以外は要確認）
BODY_STRONG_LABELS = ("内部稿・案", "リーク等の社内メモ", "プレースホルダ", "伏字プレースホルダ")


def parse_pdf_date(raw: str) -> datetime | None:
    if not raw:
        return None
    s = str(raw).strip()
    if s.startswith("D:"):
        s = s[2:]
    m = re.match(
        r"(?P<y>\d{4})(?P<mo>\d{2})(?P<d>\d{2})"
        r"(?P<h>\d{2})?(?P<mi>\d{2})?(?P<se>\d{2})?"
        r"(?P<tz>.*)?",
        s,
    )
    if not m:
        try:
            iso = str(raw).replace("Z", "+00:00")
            dt = datetime.fromisoformat(iso)
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=JST)
            return dt.astimezone(JST)
        except Exception:
            return None
    try:
        y, mo, d = int(m.group("y")), int(m.group("mo")), int(m.group("d"))
        hh = int(m.group("h") or 0)
        mi = int(m.group("mi") or 0)
        se = int(m.group("se") or 0)
        dt = datetime(y, mo, d, hh, mi, se)
    except Exception:
        return None
    tz = (m.group("tz") or "").strip()
    if not tz:
        return dt.replace(tzinfo=JST)
    if tz.startswith("Z"):
        return dt.replace(tzinfo=timezone.utc).astimezone(JST)
    zm = re.match(r"([+-])(\d{2})'?(\d{2})'?", tz)
    if zm:
        sign = 1 if zm.group(1) == "+" else -1
        offset = timedelta(hours=sign * int(zm.group(2)), minutes=sign * int(zm.group(3)))
        return dt.replace(tzinfo=timezone(offset)).astimezone(JST)
    return dt.replace(tzinfo=JST)


def parse_user_date(raw: str | None) -> date | None:
    if not raw:
        return None
    s = raw.strip()
    m = re.match(r"(\d{4})[-/\.年]\s*(\d{1,2})[-/\.月]\s*(\d{1,2})", s)
    if m:
        try:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except Exception:
            return None
    m = re.match(r"(\d{4})(\d{2})(\d{2})$", s)
    if m:
        try:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except Exception:
            return None
    return None


def format_dt(dt: datetime) -> str:
    return dt.astimezone(JST).strftime("%Y-%m-%d %H:%M（JST）")


def _overlap_area(a, b) -> float:
    ix0, iy0 = max(a[0], b[0]), max(a[1], b[1])
    ix1, iy1 = min(a[2], b[2]), min(a[3], b[3])
    if ix1 <= ix0 or iy1 <= iy0:
        return 0.0
    return (ix1 - ix0) * (iy1 - iy0)


def _is_simple_rect_items(items) -> bool:
    return bool(items) and len(items) == 1 and items[0][0] == "re"


def _is_simple_ellipse_items(items) -> bool:
    return bool(items) and len(items) == 4 and all(it[0] == "c" for it in items)


def _is_near_gray_or_bw(fill) -> bool:
    if not fill or not isinstance(fill, (list, tuple)) or len(fill) < 3:
        return False
    r, g, b = float(fill[0]), float(fill[1]), float(fill[2])
    return (max(r, g, b) - min(r, g, b)) <= 0.15


def _drawing_matches_rect(d, rect, tol=3.0) -> bool:
    r = d.get("rect")
    if r is None:
        return False
    crec = (float(r.x0), float(r.y0), float(r.x1), float(r.y1))
    return all(abs(crec[i] - rect[i]) <= tol for i in range(4))


def _line_is_short_request(line: str, phrase: str) -> bool:
    s = line.strip()
    return s == phrase or s == f"「{phrase}」" or (len(s) <= 20 and phrase in s)

def check_pdf_dates(rep: Report, fields: dict, opt: dict) -> None:
    """作成日時の不自然さ。詳細には準備時期が推測され得る旨を書く。"""
    created = parse_pdf_date(fields.get("CreationDate") or "")
    modified = parse_pdf_date(fields.get("ModDate") or "")
    disc = opt.get("disclosure_date")  # date | None
    today = datetime.now(JST).date()
    horizon = disc or today
    if created:
        rep.info["creation_dt"] = created.isoformat()
    if modified:
        rep.info["mod_dt"] = modified.isoformat()

    def _future(dt: datetime, label: str) -> None:
        if dt.date() > horizon + timedelta(days=1):
            when = f"開示日 {horizon.isoformat()} より後" if disc else "現在より後"
            rep.add(
                "要確認",
                "作成日時",
                label,
                f"{label}が{when}の日付になっています。作成日時から社内の準備時期が推測され得るため、提出前に整えてください",
                format_dt(dt),
            )

    def _too_old(dt: datetime, label: str) -> None:
        if disc and dt.date() < disc - timedelta(days=90):
            rep.add(
                "要確認",
                "作成日時",
                label,
                f"{label}が開示日（{disc.isoformat()}）より90日以上前です。作成日時から社内の準備時期が推測され得るため、提出前に整えてください",
                format_dt(dt),
            )

    if created:
        _future(created, "CreationDate")
        _too_old(created, "CreationDate")
    if modified:
        _future(modified, "ModDate")
        _too_old(modified, "ModDate")
    if created and modified and modified < created - timedelta(minutes=2):
        rep.add(
            "要確認",
            "作成日時",
            "ModDate / CreationDate",
            "更新日時が作成日時より前です。作成日時から社内の準備時期が推測され得るため、提出前に整えてください",
            f"CreationDate={format_dt(created)}、ModDate={format_dt(modified)}",
        )


def check_multi_pdf_dates(reports: list) -> None:
    """指定順を開示順とみなし、後のファイルの作成日時が先のファイルより早い場合に要確認。"""
    stamped = []
    for r in reports:
        raw = (r.info.get("metadata") or {}).get("CreationDate") or ""
        dt = parse_pdf_date(raw)
        if dt:
            stamped.append((r, dt))
    for i in range(1, len(stamped)):
        prev_r, prev_dt = stamped[i - 1]
        cur_r, cur_dt = stamped[i]
        if cur_dt < prev_dt - timedelta(minutes=1):
            cur_r.add(
                "要確認",
                "作成日時",
                "複数ファイル",
                "後から開示したとみなしたファイルの作成日時が、先のファイルより早くなっています。作成日時から社内の準備時期が推測され得るため、提出前に整えてください（指定順を開示順として比較）",
                f"{Path(prev_r.path).name}: {format_dt(prev_dt)} → {Path(cur_r.path).name}: {format_dt(cur_dt)}",
            )


def check_marks_and_review(rep: Report, body: str, page_texts: list[str], labels: list[str] | None = None) -> None:
    """社外秘マークとレビュー文言。通常の開示文に出る語は要確認に留める。

    labels はページ・スライド・セルなど、page_texts と同じ並びの位置の名前。
    """
    body = _nfkc(body)
    if not body:
        return
    apology_ok = bool(APOLOGY_CONTEXT.search(body))
    already = {(f.category, f.message, f.location) for f in rep.findings}

    def _add(sev, cat, loc, msg, detail=""):
        key = (cat, msg, loc)
        if key in already:
            return
        already.add(key)
        rep.add(sev, cat, loc, msg, detail)

    def _match(pat, flt, text) -> bool:
        for m in pat.finditer(text):
            if flt is None or flt(text, m):
                return True
        return False

    pages = [_nfkc(t) for t in (page_texts or [body])]
    for i, text in enumerate(pages, 1):
        if labels and i - 1 < len(labels):
            loc = labels[i - 1]
        else:
            loc = f"{i}ページ" if page_texts else "本文"
        for label, pat in SECRET_STRONG:
            if pat.search(text):
                _add("要修正", "社外秘マーク", loc, f"「{label}」などの提出前に残したくない表記があります", label)
        for m in CONFIDENTIAL_RE.finditer(text):
            ctx = text[max(0, m.start() - 40): m.end() + 40]
            if CONFIDENTIAL_IN_SENTENCE.search(ctx):
                continue
            # お知らせをご確認... とは無関係。文中の legal は上で除外
            # 欄外・単独・大文字見出しは要修正、それ以外は要確認
            line = text[:m.start()].rsplit("\n", 1)[-1] + text[m.start():].split("\n", 1)[0]
            standalone = len(line.strip()) <= 40
            if standalone or line.strip().isupper() or "CONFIDENTIAL" in line:
                _add("要修正", "社外秘マーク", loc, "「Confidential」などの提出前に残したくない表記があります", "Confidential")
            else:
                _add("要確認", "社外秘マーク", loc, "「Confidential」という語があります。透かしや社内向け表記でないか確認してください", "Confidential")
        for label, pat, flt in REVIEW_STRONG:
            if _match(pat, flt, text):
                _add("要修正", "レビュー文言", loc, f"社内向けの指示やメモ（{label}）が残っています", label)
        for label, pat, flt in REVIEW_SOFT:
            if label == "謝罪の文" and apology_ok:
                continue
            if _match(pat, flt, text):
                _add("要確認", "レビュー文言", loc, f"社内向けの指示やメモの可能性があります（{label}）", label)
        # 「ご確認ください」は「〜をご確認ください」という通常の案内を除外し、単独行だけ
        for line in text.splitlines():
            s = line.strip()
            if s in {"ご確認ください", "ご確認下さい", "「ご確認ください」", "ご確認ください。", "ご確認下さい。"}:
                _add("要確認", "レビュー文言", loc, "社内向けの指示やメモの可能性があります（ご確認ください）", "ご確認ください")


def check_span_watermarks(page, page_no: int, rep: Report) -> None:
    """回転や大きな文字の社外秘・Confidential 透かし。"""
    try:
        td = page.get_text("dict")
    except Exception:
        return
    for block in td.get("blocks", []):
        for line in block.get("lines", []):
            direction = line.get("dir") or (1, 0)
            rotated = abs(float(direction[1])) > 0.3
            for span in line.get("spans", []):
                s = _nfkc(span.get("text") or "")
                size = span.get("size") or 0
                if not s.strip():
                    continue
                strong = any(pat.search(s) for _, pat in SECRET_STRONG)
                conf = bool(CONFIDENTIAL_RE.search(s)) and not CONFIDENTIAL_IN_SENTENCE.search(s)
                if not (strong or conf):
                    continue
                if rotated or size >= 24:
                    label = "社外秘" if strong else "Confidential"
                    rep.add(
                        "要修正",
                        "社外秘マーク",
                        f"{page_no}ページ",
                        f"透かしまたは大きな文字で「{label}」などの表記があります",
                        label,
                    )


def check_hidden_images(page, page_no: int, rep: Report) -> None:
    """描画順で、後から重ねた不透明な図形や別画像が大部分を覆っている画像を検知する。"""
    try:
        log = list(page.get_bboxlog() or [])
    except Exception:
        return
    try:
        drawings = list(page.get_drawings() or [])
    except Exception:
        drawings = []
    try:
        infos = list(page.get_image_info(xrefs=True) or [])
    except Exception:
        infos = []

    page_rect = page.rect
    page_area = max(float(page_rect.width * page_rect.height), 1.0)
    remaining_infos = list(infos)
    remaining_drawings = [d for d in drawings if (d.get("type") or "") in {"f", "fs"}]
    events = []
    for item in log:
        typ, r = item[0], item[1]
        rect = (float(r[0]), float(r[1]), float(r[2]), float(r[3]))
        info = None
        drawing = None
        if typ == "fill-image" and remaining_infos:
            best_i, best_d = None, 10**9
            for i, inf in enumerate(remaining_infos):
                b = inf.get("bbox")
                if not b:
                    continue
                d = sum(abs(float(b[j]) - float(rect[j])) for j in range(4))
                if d < best_d:
                    best_d, best_i = d, i
            if best_i is not None and best_d <= 6:
                info = remaining_infos.pop(best_i)
        elif typ == "fill-path" and remaining_drawings:
            for i, d in enumerate(remaining_drawings):
                if _drawing_matches_rect(d, rect, tol=4.0):
                    drawing = remaining_drawings.pop(i)
                    break
        events.append((typ, rect, info, drawing))

    hits = 0
    for i, (typ, irec, info, _) in enumerate(events):
        if typ != "fill-image":
            continue
        w, h = irec[2] - irec[0], irec[3] - irec[1]
        area = max(w * h, 1.0)
        if w < 40 or h < 40 or area / page_area > 0.8:
            continue
        flagged = False
        for typ2, crec, later, drawing in events[i + 1 :]:
            if typ2 not in {"fill-path", "fill-image"}:
                continue
            c_area = max((crec[2] - crec[0]) * (crec[3] - crec[1]), 1.0)
            if c_area / page_area > 0.85:
                continue
            inter = _overlap_area(irec, crec)
            frac = inter / area
            if typ2 == "fill-path":
                if not drawing:
                    continue
                opac = drawing.get("fill_opacity")
                if opac is not None and float(opac) < 0.9:
                    continue
                items = drawing.get("items") or []
                # ほぼ同じ大きさの不透明な矩形で画像を隠す型
                # 小さな円で顔を隠す型は、地図の点やバッジと区別しにくいため見送り
                if _is_simple_rect_items(items):
                    if frac >= 0.85 and c_area <= area * 1.3:
                        flagged = True
                        break
                continue
            if typ2 == "fill-image":
                if info and later:
                    if info.get("xref") and info.get("xref") == later.get("xref"):
                        continue
                    if info.get("digest") and info.get("digest") == later.get("digest"):
                        continue
                iou = inter / (area + c_area - inter)
                iw, ih = max(irec[2] - irec[0], 1.0), max(irec[3] - irec[1], 1.0)
                cw, ch = max(crec[2] - crec[0], 1.0), max(crec[3] - crec[1], 1.0)
                wr, hr = cw / iw, ch / ih
                # ほぼ同じ位置・同じ大きさの別画像で隠す型。影や少しずれた重ねは除外
                if iou >= 0.95 and 0.97 <= wr <= 1.05 and 0.97 <= hr <= 1.05:
                    flagged = True
                    break
        if flagged:
            hits += 1
    if hits:
        rep.add(
            "要修正",
            "覆われた画像",
            f"{page_no}ページ",
            "後から重ねた不透明な図形または別の画像の下に、元の画像オブジェクトが残っています。選択やコピーで取り出せる状態です",
            f"{hits}箇所",
        )



def _norm_name(s: str) -> str:
    if not s:
        return ""
    table = str.maketrans({chr(ord("Ａ") + i): chr(ord("A") + i) for i in range(26)} |
                          {chr(ord("ａ") + i): chr(ord("a") + i) for i in range(26)} |
                          {chr(ord("０") + i): chr(ord("0") + i) for i in range(10)})
    return s.translate(table).replace(" ", "").replace("\u3000", "").replace("㈱", "株式会社")


@dataclass
class Finding:
    severity: str
    category: str
    location: str
    message: str
    detail: str = ""

    def to_dict(self) -> dict:
        return asdict(self)


@dataclass
class Report:
    path: str
    findings: list[Finding] = field(default_factory=list)
    info: dict = field(default_factory=dict)

    def add(self, *args, **kwargs) -> None:
        self.findings.append(Finding(*args, **kwargs))

    @property
    def worst(self) -> str | None:
        if not self.findings:
            return None
        return min(self.findings, key=lambda f: SEVERITY_ORDER.get(f.severity, 9)).severity


def _read_ng_list(value: str | None) -> list[str]:
    if not value:
        return []
    p = Path(value)
    if p.is_file():
        return [ln.strip() for ln in p.read_text(encoding="utf-8").splitlines() if ln.strip()]
    return [x.strip() for x in value.split(",") if x.strip()]


def _hit_ng(text: str, extra: Iterable[str] = ()) -> list[str]:
    if not text:
        return []
    text = _nfkc(text)
    hits: list[str] = []
    for label, pat in NG_PATTERNS:
        if pat.search(text):
            hits.append(label)
    for w in extra:
        w = _nfkc(w)
        if not w:
            continue
        if len(w) == 1:
            continue  # 1文字は議案等への誤ヒットが多い
        if w in text:
            hits.append(f"指定NG語「{w}」")
    # 重複除去（順維持）
    out: list[str] = []
    for h in hits:
        if h not in out:
            out.append(h)
    return out


_SOFTWARE_AUTHORS = {
    "anonymous", "unknown", "unspecified", "user", "admin", "administrator",
    "adobe", "microsoft", "reportlab", "writer", "owner", "nobody", "n/a",
    "windows user", "windows ユーザー", "microsoft office user", "microsoft office ユーザー",
    "office user", "ユーザー", "ユーザ", "管理者",
}
# 部署名（経営企画部、IR室など）は人名として扱わない
_DEPT_SUFFIX = re.compile(r"(部|室|課|局|係|本部|グループ|チーム|センター|事務局|委員会|部門)$")


def _looks_like_person_name(author: str) -> bool:
    a = _nfkc(author).strip()
    if not a:
        return False
    if a.lower() in _SOFTWARE_AUTHORS:
        return False
    if _DEPT_SUFFIX.search(a):
        return False
    if any(k in a for k in ("株式会社", "有限会社", "合同会社", "(株)", "Inc", "LLC", "Corp", "Ltd", "自社", "IR")):
        return False
    # 姓名の区切りがある場合（山田 太郎 / Taro Yamada / 山本・俊）
    if re.fullmatch(r"[一-龥ァ-ンA-Za-z]{1,8}[　\s・]+[一-龥ァ-ンA-Za-z]{1,8}", a) and len(a) <= 20:
        return True
    # 区切りなしの日本語人名（3〜8文字、会社名などを除外）
    if re.fullmatch(r"[一-龥ァ-ン]{3,8}", a):
        return True
    # ログイン名風（s-tanaka, j.inoue）
    if re.fullmatch(r"[a-z][a-z]*[\.\-_][a-z]+", a, re.I):
        return True
    return False


def _parse_reiwa(ytext: str) -> int | None:
    ytext = ytext.replace(" ", "")
    if ytext.startswith("20"):
        return int(ytext)
    m = re.search(r"令和\s*([元0-9０-９]+)", ytext)
    if not m:
        return None
    n = m.group(1)
    n = "1" if n == "元" else n.translate(str.maketrans("０１２３４５６７８９", "0123456789"))
    return 2018 + int(n)


# ---------------------------------------------------------------------------
# PDF
# ---------------------------------------------------------------------------
def inspect_pdf(path: Path, opt: dict) -> Report:
    import pymupdf

    rep = Report(str(path))
    try:
        doc = pymupdf.open(path)
    except Exception as e:
        msg = str(e)
        if any(w in msg.lower() for w in ("password", "encrypt", "passwd")):
            rep.add(
                "要修正",
                "セキュリティ",
                "ファイル",
                "パスワード付きのPDFです。パスワードなしでは開けません。提出用にはパスワードや権限制限を外してください",
                msg[:200],
            )
        else:
            rep.add("未検査", "ファイル", "全体", "PDFを開けないため検査していません。壊れていないか確かめてください", type(e).__name__)
        return rep

    # ユーザーパスワード付きで開けない場合は、以降の検査をせず要修正で返す
    if getattr(doc, "needs_pass", False):
        ok = False
        try:
            ok = bool(doc.authenticate(""))
        except Exception:
            ok = False
        if not ok:
            rep.add(
                "要修正",
                "セキュリティ",
                "ファイル",
                "パスワード付きのPDFです。パスワードなしでは開けません。提出用にはパスワードや権限制限を外してください",
            )
            try:
                doc.close()
            except Exception:
                pass
            return rep

    # 暗号化・権限（空パスワードで開ける場合も含む）
    encryption = (doc.metadata or {}).get("encryption")
    if encryption:
        perm = int(getattr(doc, "permissions", -1))  # 符号付きの値。ビットで判定する
        bad = []
        if not perm & pymupdf.PDF_PERM_COPY:
            bad.append("コピー（抽出）禁止")
        if not perm & pymupdf.PDF_PERM_PRINT:
            bad.append("印刷禁止")
        if bad:
            rep.add(
                "要修正",
                "セキュリティ",
                "ファイル",
                "TDnetでは印刷・検索・コピー禁止の設定をしないよう案内されています。権限の制限を外して保存し直してください",
                "、".join(bad),
            )
        else:
            rep.add(
                "情報",
                "セキュリティ",
                "ファイル",
                "暗号化の設定があります（印刷・コピーは許可されています）",
                str(encryption),
            )

    meta = doc.metadata or {}
    xmp = ""
    try:
        raw = doc.get_xml_metadata()
        if isinstance(raw, str):
            xmp = raw
        elif raw:
            xmp = str(raw)
    except Exception:
        xmp = ""

    fields = {
        "Title": meta.get("title") or "",
        "Author": meta.get("author") or "",
        "Subject": meta.get("subject") or "",
        "Keywords": meta.get("keywords") or "",
        "Creator": meta.get("creator") or "",
        "Producer": meta.get("producer") or "",
        "CreationDate": meta.get("creationDate") or "",
        "ModDate": meta.get("modDate") or "",
    }
    rep.info["metadata"] = {k: v for k, v in fields.items() if v}
    rep.info["pages"] = doc.page_count
    check_pdf_dates(rep, fields, opt)

    extra_ng = opt.get("ng_words") or []
    company = (opt.get("company") or "").strip()
    code = _norm_name((opt.get("code") or "").strip()).upper()

    for key, val in fields.items():
        if not val:
            continue
        hits = _hit_ng(val, extra_ng)
        if hits:
            sev = "要修正" if key in ("Title", "Author", "Subject", "Keywords") else "要確認"
            rep.add(sev, "メタデータ", _lab(key), f"{_lab(key)}に公開したくない表現があります（{('、').join(hits)}）", val)
        if key in ("Author", "LastModifiedBy") and _looks_like_person_name(val):
            rep.add("要修正", "メタデータ", _lab(key), "作成者が個人名またはログイン名に見えます", val)
        if company and key in ("Title", "Author", "Subject", "Keywords") and re.search(r"株式会社|有限会社|合同会社", val):
            # 自社名を含まない組織名
            nv, nc = _norm_name(val), _norm_name(company)
            if nc and nc not in nv and _norm_name(company.replace("株式会社", "")) not in nv:
                sev = "要修正" if key == "Author" else "要確認"
                rep.add(sev, "メタデータ", _lab(key), f"{_lab(key)}に自社以外の会社名がある可能性があります", val)
        if key in ("Title", "Subject", "Keywords", "Author") and code:
            for m in CODE_RE.finditer(val):
                other = m.group(1).upper()
                if other == code or other in {code.lstrip("0")}:
                    continue
                if re.fullmatch(r"20\d{2}", other):
                    continue
                # 時刻（1107=11:07）や日付の一部、v8 の直前は除外
                if re.fullmatch(r"([01]\d|2[0-3])[0-5]\d", other):
                    continue
                ctx = val[max(0, m.start() - 6): m.end() + 4]
                if re.search(r"\d" + re.escape(m.group(1)) + r"\d", val):
                    continue
                if re.search(re.escape(m.group(1)) + r"v\d", val, re.I):
                    continue
                if re.search(r"(コード|証券)", ctx):
                    rep.add("要確認", "メタデータ", _lab(key), f"{_lab(key)}に自社以外の証券コードらしき値があります", other)

    if xmp:
        hits = _hit_ng(xmp, extra_ng)
        # メタデータ項目と重複しがちなので、XMPにしかない場合のみ
        already = " ".join(f.detail or "" for f in rep.findings if f.category == "メタデータ")
        leftover = [h for h in hits if h not in already]
        if leftover and any(p.search(xmp) for _, p in NG_PATTERNS if _ not in ("ドラフト",)):
            # Title 等ですでに拾っているなら追加しない
            if not any(f.category == "メタデータ" and f.severity == "要修正" for f in rep.findings):
                rep.add("要確認", "メタデータ", "XMPメタデータ", "XMPメタデータに要注意語があります", "、".join(hits[:5]))

    # JavaScript
    try:
        if "/Names" in doc.trailer_xml:  # 使えない場合あり
            pass
    except Exception:
        pass
    try:
        # catalog JS
        for xref in range(1, doc.xref_length()):
            try:
                s = doc.xref_object(xref)
            except Exception:
                continue
            if s and "/JavaScript" in s or (s and "/JS" in s and "S /JavaScript" in s):
                rep.add("要確認", "JavaScript", "ファイル", "PDF内にJavaScriptがあります")
                break
    except Exception:
        pass

    # 埋め込みファイル
    try:
        n_emb = doc.embfile_count()
        if n_emb:
            names = []
            for i in range(n_emb):
                try:
                    info = doc.embfile_info(i)
                    names.append(info.get("filename") or info.get("name") or str(i))
                except Exception:
                    names.append(str(i))
            rep.add("要修正", "埋め込みファイル", "ファイル", f"埋め込みファイルが {n_emb} 件あります", ", ".join(names))
    except Exception:
        pass

    # しおり
    try:
        for item in doc.get_toc() or []:
            title = item[1] if len(item) > 1 else ""
            hits = _hit_ng(title, extra_ng)
            if hits:
                rep.add("要修正", "しおり", title, "しおりに要注意語があります", "、".join(hits))
    except Exception:
        pass

    # 非表示レイヤー (OCG)
    try:
        ocgs = doc.get_ocgs() or {}
        if ocgs:
            hidden = []
            for name, info in ocgs.items():
                on = info.get("on", True) if isinstance(info, dict) else True
                label = info.get("name", name) if isinstance(info, dict) else str(name)
                if on is False:
                    hidden.append(str(label))
            if hidden:
                rep.add("要修正", "非表示レイヤー", "レイヤー", "非表示のオプションコンテンツ（レイヤー）があります", "、".join(hidden))
            else:
                rep.add("情報", "非表示レイヤー", "レイヤー", f"オプションコンテンツ（レイヤー）が {len(ocgs)} 件あります。表示状態を目視確認してください")
    except Exception:
        pass

    # 注釈
    annot_n = 0
    for i, page in enumerate(doc, 1):
        try:
            annots = list(page.annots() or [])
        except Exception:
            annots = []
        for a in annots:
            annot_n += 1
            info = a.info or {}
            content = (info.get("content") or info.get("subject") or a.type[1] if a.type else "") or ""
            typ = a.type[1] if a.type else "Annot"
            # リンクとフォーム部品は対象外
            if typ in {"Link", "Widget"}:
                continue
            if typ == "Redact":
                rep.add(
                    "要修正",
                    "墨消し未適用",
                    f"{i}ページ",
                    "墨消しの範囲を指定しただけで、適用されていません。下の文字は残っています。Acrobatで「墨消しを適用」を実行して保存し直してください",
                )
                continue
            sev = "要修正" if typ in {"Text", "FreeText", "Highlight", "Caret", "StrikeOut", "Underline", "Popup", "Note"} else "要確認"
            rep.add(sev, "注釈", f"{i}ページ / {typ}", "注釈またはコメントが残っています", str(content)[:200])
    if annot_n:
        rep.info["annotations"] = annot_n

    # フォント未埋め込み（pdffonts の emb 列を優先。pymupdf のフラグは Type0 で誤ることがある）
    unemb: set[str] = set()
    try:
        import subprocess
        pf = subprocess.check_output(["pdffonts", str(path)], text=True, errors="replace", stderr=subprocess.DEVNULL, timeout=60)
        plines = pf.splitlines()
        if len(plines) >= 3 and "emb" in plines[0]:
            col = plines[0].find("emb")
            for line in plines[2:]:
                if len(line) <= col + 2:
                    continue
                name = line[:col].strip()
                emb = line[col:col + 3].strip().lower()
                if emb == "no" and name:
                    unemb.add(name)
    except Exception:
        unemb = set()
    BASE14_PFX = ("helvetica", "times", "courier", "symbol", "zapfdingbats", "times-roman")
    def _is_base14(name: str) -> bool:
        token = name.split()[0].split("+")[-1].lower().replace(",", "")
        return any(token.startswith(p) for p in BASE14_PFX)
    unemb_jp = {n.split()[0] for n in unemb if not _is_base14(n)}
    unemb_base = {n.split()[0] for n in unemb if _is_base14(n)}
    if unemb_jp:
        rep.add(
            "要確認",
            "フォント",
            "ファイル",
            "埋め込みされていないフォントがあります。他環境で字形が崩れることがあります",
            "、".join(sorted(unemb_jp)[:8]),
        )
    elif unemb_base:
        rep.add(
            "情報",
            "フォント",
            "ファイル",
            "標準フォント（Helvetica等）が未埋め込みです。日本語フォントでなければ通常は問題になりません",
            "、".join(sorted(unemb_base)[:8]),
        )

    # 本文・描画オブジェクト
    redact_words = opt.get("redact_words") or []
    body_parts: list[str] = []
    page_hashes: dict[str, list[int]] = {}
    printed_pages: list[tuple[int, int]] = []  # (physical, detected number)
    white_hits: list[tuple[int, int]] = []  # (ページ, 文字数)。文字列そのものは持たない
    tiny_hits: list[tuple[int, int, float]] = []

    def _overlap(a, b, ratio=0.3) -> bool:
        ax0, ay0, ax1, ay1 = a
        bx0, by0, bx1, by1 = b
        ix0, iy0 = max(ax0, bx0), max(ay0, by0)
        ix1, iy1 = min(ax1, bx1), min(ay1, by1)
        if ix1 <= ix0 or iy1 <= iy0:
            return False
        inter = (ix1 - ix0) * (iy1 - iy0)
        area = max((ax1 - ax0) * (ay1 - ay0), 1)
        return inter / area >= ratio

    def _table_band_rects(rects: list[tuple]) -> set[tuple]:
        """同じ幅で縦に並ぶ塗りつぶしは表の行背景とみなす。"""
        skip: set[tuple] = set()
        for i, a in enumerate(rects):
            siblings = [a]
            for b in rects:
                if a is b:
                    continue
                if abs(a[0] - b[0]) <= 3 and abs((a[2] - a[0]) - (b[2] - b[0])) <= 3:
                    # 同程度の高さで縦にずれている
                    if abs((a[3] - a[1]) - (b[3] - b[1])) <= 8:
                        siblings.append(b)
            if len(siblings) >= 3:
                skip.update(siblings)
        return skip

    def _is_black(color) -> bool:
        if color is None:
            return False
        if isinstance(color, (int, float)):
            return color <= 0.08
        if isinstance(color, (list, tuple)) and len(color) >= 3:
            return all(c <= 0.08 for c in color[:3])
        return False

    def _is_white(color) -> bool:
        if color is None:
            return False
        if isinstance(color, (int, float)):
            return color >= 0.95
        if isinstance(color, (list, tuple)) and len(color) >= 3:
            return all(c >= 0.95 for c in color[:3])
        return False


    def _overlap_fracs(a, b) -> tuple[float, float]:
        ax0, ay0, ax1, ay1 = a
        bx0, by0, bx1, by1 = b
        ix0, iy0 = max(ax0, bx0), max(ay0, by0)
        ix1, iy1 = min(ax1, bx1), min(ay1, by1)
        if ix1 <= ix0 or iy1 <= iy0:
            return 0.0, 0.0
        inter = (ix1 - ix0) * (iy1 - iy0)
        area_a = max((ax1 - ax0) * (ay1 - ay0), 1.0)
        area_b = max((bx1 - bx0) * (by1 - by0), 1.0)
        return inter / area_a, inter / area_b

    def _smaller_cover(a, b) -> float:
        ra, rb = _overlap_fracs(a, b)
        area_a = max((a[2] - a[0]) * (a[3] - a[1]), 1.0)
        area_b = max((b[2] - b[0]) * (b[3] - b[1]), 1.0)
        return ra if area_a <= area_b else rb

    def _norm_span(s: str) -> str:
        return re.sub(r"\s+", "", s)

    period = (opt.get("period") or "").strip()

    for i, page in enumerate(doc, 1):
        rect = page.rect
        text = _nfkc(page.get_text("text") or "")
        body_parts.append(text)
        h = hashlib.sha1(re.sub(r"\s+", "", text).encode("utf-8", "ignore")).hexdigest()
        page_hashes.setdefault(h, []).append(i)

        # ページ番号らしきもの（フッター）
        m = re.search(r"(?:^|\n)\s*(?:[-–—]?\s*)(\d{1,3})\s*(?:/|／)\s*(\d{1,3})\s*(?:[-–—])?\s*$", text)
        if m:
            printed_pages.append((i, int(m.group(1))))
        else:
            # 末尾の単独数字
            lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
            if lines and re.fullmatch(r"\d{1,3}", lines[-1]) and 1 <= int(lines[-1]) <= doc.page_count + 5:
                printed_pages.append((i, int(lines[-1])))

        # 描画（黒塗り矩形）
        black_rects = []
        try:
            for d in page.get_drawings():
                fill = d.get("fill")
                typ = d.get("type") or ""
                if typ in {"f", "fs", "fclip"} or (fill is not None and typ != "s"):
                    if _is_black(fill):
                        r = d.get("rect")
                        if r and r.width > 4 and r.height > 4:
                            black_rects.append((r.x0, r.y0, r.x1, r.y1))
        except Exception:
            black_rects = []

        table_rects = _table_band_rects(black_rects)
        redaction_rects = [r for r in black_rects if r not in table_rects]

        try:
            td = page.get_text("dict")
        except Exception:
            td = {"blocks": []}


        pix = None
        page_spans: list[dict] = []
        offpage_n = 0
        def _light_background(bbox) -> bool:
            nonlocal pix
            try:
                if pix is None:
                    pix = page.get_pixmap(matrix=pymupdf.Matrix(1, 1), alpha=False)
                x = int(min(max((bbox[0] + bbox[2]) / 2, 0), pix.width - 1))
                y = int(min(max((bbox[1] + bbox[3]) / 2, 0), pix.height - 1))
                # 文字の中心はグリフ自体が白いので、少し上の背景を見る
                y2 = int(min(max(bbox[1] - 2, 0), pix.height - 1))
                samples = [pix.pixel(x, y2)]
                # 左右
                x3 = int(min(max(bbox[0] - 2, 0), pix.width - 1))
                samples.append(pix.pixel(x3, int(min(max((bbox[1]+bbox[3])/2, 0), pix.height-1))))
                def _lum(rgb):
                    return sum(rgb[:3]) / 3
                return sum(_lum(s) for s in samples) / len(samples) >= 230
            except Exception:
                return False

        for block in td.get("blocks", []):
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    s = _nfkc(span.get("text") or "").strip()
                    if not s:
                        continue
                    bbox = span.get("bbox") or (0, 0, 0, 0)
                    size = span.get("size") or 0
                    color = span.get("color")
                    # color は 整数 0xRRGGBB のことがある
                    rgb = None
                    if isinstance(color, int):
                        rgb = ((color >> 16) & 255) / 255, ((color >> 8) & 255) / 255, (color & 255) / 255
                    elif isinstance(color, (list, tuple)):
                        rgb = tuple(color[:3])
                    page_spans.append({"text": s, "bbox": bbox, "size": size, "rgb": rgb})

                    # ページ外
                    margin = 2
                    if bbox[2] < rect.x0 - margin or bbox[0] > rect.x1 + margin or bbox[3] < rect.y0 - margin or bbox[1] > rect.y1 + margin:
                        offpage_n += 1
                    # 極小
                    if 0 < size < 3:
                        tiny_hits.append((i, len(s), size))
                    # 白文字（白地に白だけを要確認。色付き背景上の白文字はデザイン）
                    if rgb and _is_white(rgb) and not any(_overlap(bbox, br, 0.4) for br in black_rects):
                        if _light_background(bbox):
                            white_hits.append((i, len(s)))
        if offpage_n:
            rep.add("要確認", "ページ外テキスト", f"{i}ページ", "ページ範囲の外にテキストがあります", f"{offpage_n}箇所")

        # 黒塗りの下の文字。行の一部だけを塗った場合も拾えるよう、文字ごとに中心が黒い矩形の内側かを見る
        covered_chars = 0
        if redaction_rects:
            try:
                raw = page.get_text("rawdict")
            except Exception:
                raw = {"blocks": []}
            for block in raw.get("blocks", []):
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        color = span.get("color")
                        if isinstance(color, int):
                            lum = (((color >> 16) & 255) + ((color >> 8) & 255) + (color & 255)) / 765
                        else:
                            lum = 0.0
                        if lum >= 0.5:
                            continue  # 黒地の白抜き文字はデザイン
                        for ch in span.get("chars", []):
                            if not (ch.get("c") or "").strip():
                                continue
                            cb = ch.get("bbox") or (0, 0, 0, 0)
                            cx, cy = (cb[0] + cb[2]) / 2, (cb[1] + cb[3]) / 2
                            if any(r[0] <= cx <= r[2] and r[1] <= cy <= r[3] for r in redaction_rects):
                                covered_chars += 1
        if covered_chars:
            rep.add(
                "要修正",
                "黒塗りの残存",
                f"{i}ページ",
                "黒い図形の下に文字データが残っています（図形を重ねただけの黒塗り）。選択やコピーで読める状態です",
                f"{covered_chars}文字分",
            )

        # 伏せ表記・見えにくい文字の重ね残り、色を問わない覆いの下の文字
        hidden_n = 0
        visible_labels: list[str] = []
        for j, a in enumerate(page_spans):
            ta = _norm_span(a["text"])
            a_low = bool(a["rgb"] and _is_white(a["rgb"]) and _light_background(a["bbox"]))
            a_anon = bool(ANON_LABEL_RE.match(a["text"]))
            for b in page_spans[j + 1 :]:
                tb = _norm_span(b["text"])
                if not ta or not tb or ta == tb or ta in tb or tb in ta:
                    continue
                if len(ta) == 1 and len(tb) == 1:
                    continue
                smaller = _smaller_cover(a["bbox"], b["bbox"])
                if smaller < 0.4:
                    continue
                sa, sb = a["size"] or 0, b["size"] or 0
                ratio = max(sa, sb) / max(min(sa, sb), 0.1)
                if ratio > 2.5:
                    continue
                b_low = bool(b["rgb"] and _is_white(b["rgb"]) and _light_background(b["bbox"]))
                b_anon = bool(ANON_LABEL_RE.match(b["text"]))
                contrast_diff = a_low != b_low
                if (a_anon or b_anon) or contrast_diff:
                    hidden_n += 1
                    for lab, flag in ((a["text"], a_anon), (b["text"], b_anon)):
                        if flag and lab not in visible_labels:
                            visible_labels.append(lab)
        if hidden_n:
            label_bit = f"（見える側の例: {'、'.join(visible_labels[:6])}）" if visible_labels else ""
            rep.add(
                "要修正",
                "伏せテキストの残存",
                f"{i}ページ",
                "伏せ表記や別ラベルと同じ位置に、見た目では見えない別のテキストが残っています（選択やコピーで読める状態です）",
                f"{hidden_n}箇所{label_bit}",
            )
            # 伏せ文字そのものを白文字の詳細に出さない
            white_hits[:] = [(pg, n) for pg, n in white_hits if pg != i]
            tiny_hits[:] = [(pg, n, sz) for pg, n, sz in tiny_hits if pg != i]

        check_hidden_images(page, i, rep)
        check_span_watermarks(page, i, rep)

        if len(black_rects) >= 8:
            # 表の塗りつぶし見出しであることが多い。墨消し語が指定されていなければ情報に留める
            if not redact_words:
                rep.add(
                    "情報",
                    "黒塗りの残存",
                    f"{i}ページ",
                    "黒または濃色の塗りつぶし図形が多数あります。表の見出しか、図形を重ねただけの黒塗りでないか、目視で確認してください",
                )

        if redact_words:
            for w in redact_words:
                w = _nfkc(w)
                if w and w in text:
                    rep.add("要修正", "墨消し検証", f"{i}ページ", f"伏せたはずの語「{w}」が本文から抽出できました")

    body = "\n".join(body_parts)
    rep.info["chars"] = len(body)
    check_marks_and_review(rep, body, body_parts)
    if white_hits:
        pages = sorted({pg for pg, _ in white_hits})
        if len(white_hits) <= 5:
            for pg in pages:
                n = sum(1 for p_, _ in white_hits if p_ == pg)
                rep.add("要確認", "白文字", f"{pg}ページ", "白地に白い（または極めて薄い）文字があります。非表示の意図がないか確認してください", f"{n}箇所")
        else:
            rep.add(
                "情報",
                "白文字",
                f"{len(pages)}ページに計{len(white_hits)}箇所",
                "白文字が多数あります。色付き背景の白抜き見出しであることが多いですが、白地への隠し文字がないか代表箇所を目視確認してください",
                "該当ページ: " + "、".join(str(p_) for p_ in pages[:10]),
            )
    if tiny_hits:
        pages = sorted({pg for pg, _, _ in tiny_hits})
        if len(tiny_hits) <= 5:
            for pg, _, size in tiny_hits:
                rep.add("要確認", "極小文字", f"{pg}ページ", f"{size:.1f}pt の極小テキストがあります", "")
        else:
            rep.add(
                "情報",
                "極小文字",
                f"{len(pages)}ページに計{len(tiny_hits)}箇所",
                "極小テキストが多数あります。グラフの軸目盛や装飾であることが多いですが、隠し文字がないか代表箇所を目視確認してください",
                "該当ページ: " + "、".join(str(p_) for p_ in pages[:10]),
            )

    hits = _hit_ng(body, extra_ng)
    # メタデータで既出でも本文に残っていれば再掲（本文は公開文面）
    if hits:
        # 「議案」は NG に入れてない。本文の（案）は要修正
        sev = "要修正" if any(h in BODY_STRONG_LABELS for h in hits) else "要確認"
        rep.add(sev, "本文", "全体", f"本文に要注意表現があります（{'、'.join(hits)}）")

    # 日付と曜日
    for m in WEEKDAY_RE.finditer(body):
        y = _parse_reiwa(m.group("y"))
        try:
            dt = date(y, int(m.group("m")), int(m.group("d")))
        except Exception:
            continue
        actual = WEEKDAYS[dt.weekday()]
        written = m.group("w")
        if actual != written:
            rep.add(
                "要修正",
                "日付と曜日",
                next((f"{n}ページ" for n, t in enumerate(body_parts, 1) if m.group(0) in t), "本文"),
                f"日付と曜日が一致しません（{dt.year}年{dt.month}月{dt.day}日は{actual}曜日）",
                m.group(0) + ")",
            )

    # 自社以外のコード（本文は誤検知が多いので、表紙＝1ページ目に限定して要確認）
    if code and body_parts:
        first = body_parts[0][:800]
        for m in CODE_RE.finditer(first):
            other = _norm_name(m.group(1)).upper()
            if other == code or re.fullmatch(r"20\d{2}", other):
                continue
            # 電話番号の一部などを避けるため、コード番号の文脈だけ
            if re.search(r"(コード|証券|東証|名証|札証|福証)", first):
                # ページ内に自社コードと別コード
                if other != code:
                    # 市場の「スタンダード」近くの数字は除外しにくい。コード表記の直後だけ
                    window = first[max(0, m.start() - 12) : m.end() + 4]
                    if re.search(r"(コード|証券コード|コード番号|（コード)", window):
                        if other != code:
                            rep.add("要確認", "証券コード", "1ページ目", f"自社コード({code})以外の証券コードらしき値があります", f"{window.strip()} → {other}")

    # 決算期の食い違い（指定時）
    if period:
        # 例: 2026年3月期 が資料内にあるか、別の「20xx年x月期」がないか
        periods = set(re.findall(r"20\d{2}\s*年\s*\d{1,2}\s*月期", body))
        norm = re.sub(r"\s+", "", period)
        others = [p for p in periods if re.sub(r"\s+", "", p) != norm and not re.sub(r"\s+", "", p).startswith(norm[:6])]
        # 前期・予想で別の期が出るのは正常。多くありすぎる場合は情報に留める
        # 「お知らせ（案）」とは別。指定した期が本文に一度も無いとき要確認
        if not any(re.sub(r"\s+", "", p) == norm for p in periods) and norm not in re.sub(r"\s+", "", body):
            if "月期" in body:
                rep.add("要確認", "決算期", "本文", f"指定した決算期「{period}」が本文から見つかりません。表記ゆれまたは誤記の可能性があります")

    # 他社名（メタデータは上で実施。本文は要確認・上限）
    if company:
        seen = set()
        nc = _norm_name(company)
        for m in COMPANY_RE.finditer(body):
            raw = (m.group(1) or m.group(2) or "").strip()
            if raw in _COMPANY_STOP:
                continue
            name = _norm_name(m.group(0) or "")
            if not name or nc in name or name in nc or _norm_name(raw) in nc:
                continue
            if name in seen:
                continue
            seen.add(name)
            if len(seen) <= 8:
                # 取引先として出てくるのは正常なので要確認・まとめて後で
                pass
        if seen:
            rep.add(
                "情報",
                "他社名",
                "本文",
                "本文に自社以外の「株式会社」名称があります（取引先の記載のことが多いです。テンプレ流用でないかだけ確認してください）",
                "、".join(list(seen)[:8]),
            )

    # 重複ページ
    for h, pages in page_hashes.items():
        if len(pages) >= 2 and any(len(re.sub(r"\s+", "", body_parts[p - 1])) > 80 for p in pages):
            rep.add("要確認", "重複ページ", "、".join(str(p) for p in pages), "本文がほぼ同一のページが複数あります")

    # ページ番号の飛び（説明会資料は表紙・区分スライドで欠番が多いので、欠けが多い場合は情報）
    if len(printed_pages) >= 3:
        nums = [n for _, n in printed_pages]
        expected = list(range(nums[0], nums[0] + len(nums)))
        if nums != expected:
            missing = len(set(range(min(nums), max(nums) + 1))) - len(set(nums))
            sev = "情報" if missing >= 3 or (max(nums) - min(nums) + 1) >= 15 else "要確認"
            rep.add(sev, "ページ番号", "フッター", f"印刷されたページ番号が連続していない可能性があります: {nums[:20]}")

    doc.close()
    return rep


# ---------------------------------------------------------------------------
# Office
# ---------------------------------------------------------------------------
NSMAP = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
}


def _core_props_from_zip(z: zipfile.ZipFile) -> dict:
    out = {}
    try:
        xml = z.read("docProps/core.xml")
    except KeyError:
        return out
    root = ET.fromstring(xml)
    for tag, key in (("title", "Title"), ("subject", "Subject"), ("creator", "Author"), ("keywords", "Keywords"), ("description", "Subject")):
        el = root.find(f"{{http://purl.org/dc/elements/1.1/}}{tag}")
        if el is None:
            el = root.find(f"{{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}}{tag}")
        if el is not None and el.text:
            out[key] = el.text
    # lastModifiedBy
    el = root.find("{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}lastModifiedBy")
    if el is not None and el.text:
        out["LastModifiedBy"] = el.text
    return out


def _check_core(rep: Report, props: dict, opt: dict) -> None:
    extra = opt.get("ng_words") or []
    company = (opt.get("company") or "").strip()
    for key, val in props.items():
        hits = _hit_ng(val, extra)
        if hits:
            sev = "要修正" if key in ("Title", "Author", "Subject", "Keywords", "LastModifiedBy") else "要確認"
            rep.add(sev, "メタデータ", _lab(key), f"{_lab(key)}に公開したくない表現があります（{'、'.join(hits)}）", val)
        if key in ("Author", "LastModifiedBy") and _looks_like_person_name(val):
            rep.add("要修正", "メタデータ", _lab(key), "作成者または更新者が個人名に見えます", val)
        if key in ("Author", "LastModifiedBy") and company and company not in val and re.search(r"株式会社|Inc|Corp", val):
            rep.add("要修正", "メタデータ", _lab(key), "作成者が自社名ではなく、別の組織名に見えます", val)
    rep.info["metadata"] = props


def inspect_docx(path: Path, opt: dict) -> Report:
    from docx import Document

    rep = Report(str(path))
    with zipfile.ZipFile(path) as z:
        _check_core(rep, _core_props_from_zip(z), opt)
        names = set(z.namelist())
        if "word/comments.xml" in names:
            xml = z.read("word/comments.xml")
            comments = ET.fromstring(xml).findall(f"{{{NSMAP['w']}}}comment")
            for c in comments:
                texts = [t.text or "" for t in c.iter(f"{{{NSMAP['w']}}}t")]
                author = c.get("author") or ""
                rep.add("要修正", "コメント", f"コメント（{author}）" if author else "コメント", "Wordのコメントが残っています", "".join(texts)[:200])
        # 変更履歴
        doc_xml = z.read("word/document.xml")
        root = ET.fromstring(doc_xml)
        ins = list(root.iter(f"{{{NSMAP['w']}}}ins"))
        dels = list(root.iter(f"{{{NSMAP['w']}}}del"))
        if ins or dels:
            rep.add("要修正", "変更履歴", "本文", f"変更履歴が残っています（挿入{len(ins)}件・削除{len(dels)}件）。「校閲 ＞ すべての変更を承諾」のあと保存し直してください")
        # 隠し文字（文字列は出さない）
        hidden_chars = 0
        for r in root.iter(f"{{{NSMAP['w']}}}r"):
            rPr = r.find(f"{{{NSMAP['w']}}}rPr")
            if rPr is not None and rPr.find(f"{{{NSMAP['w']}}}vanish") is not None:
                hidden_chars += len("".join(t.text or "" for t in r.iter(f"{{{NSMAP['w']}}}t")))
        if hidden_chars:
            rep.add("要修正", "隠し文字", "本文", "隠し文字（書式の「隠し文字」）が残っています。PDFには出なくてもWordファイルには残ります", f"{hidden_chars}文字分")

    doc = Document(str(path))
    body = "\n".join(p.text for p in doc.paragraphs)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                body += "\n" + cell.text
    # ヘッダ・フッタも見る
    try:
        for section in doc.sections:
            for part in (section.header, section.footer, section.first_page_header, section.first_page_footer):
                try:
                    body += "\n" + "\n".join(p.text for p in part.paragraphs)
                except Exception:
                    pass
    except Exception:
        pass
    body = _nfkc(body)
    hits = _hit_ng(body, opt.get("ng_words") or [])
    if hits:
        sev = "要修正" if any(h in BODY_STRONG_LABELS for h in hits) else "要確認"
        rep.add(sev, "本文", "本文", f"本文に要注意表現があります（{'、'.join(hits)}）")
    check_marks_and_review(rep, body, [body], ["本文"])
    return rep


def _iter_pptx_shapes(shapes):
    """グループの中まで図形をたどる。"""
    for shape in shapes:
        yield shape
        if getattr(shape, "shape_type", None) is not None:
            try:
                from pptx.enum.shapes import MSO_SHAPE_TYPE

                if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                    yield from _iter_pptx_shapes(shape.shapes)
            except Exception:
                pass


def _pptx_shape_text(shape) -> str:
    parts = []
    if getattr(shape, "has_text_frame", False):
        parts.append(shape.text_frame.text)
    if getattr(shape, "has_table", False):
        for row in shape.table.rows:
            for cell in row.cells:
                parts.append(cell.text)
    return "\n".join(p for p in parts if p)


def _pptx_comments(z: zipfile.ZipFile) -> dict[str, list[str]]:
    """スライドのファイル名 → コメント本文の一覧。旧形式（comments/comment*.xml）と新形式（modernComment）の両方。"""
    names = z.namelist()
    part_to_slide: dict[str, str] = {}
    for n in names:
        m = re.fullmatch(r"ppt/slides/_rels/(slide\d+)\.xml\.rels", n)
        if not m:
            continue
        try:
            rels = ET.fromstring(z.read(n))
        except Exception:
            continue
        for rel in rels:
            target = rel.get("Target") or ""
            if "comment" in target.lower():
                part_to_slide["ppt/comments/" + target.rsplit("/", 1)[-1]] = m.group(1)
    out: dict[str, list[str]] = {}
    for n in names:
        if not (n.startswith("ppt/comments/") and n.endswith(".xml")):
            continue
        try:
            root = ET.fromstring(z.read(n))
        except Exception:
            continue
        texts = []
        # 1件ごとの本文（旧形式は p:text、新形式は a:t の連なり）
        for el in root.iter():
            local = el.tag.rsplit("}", 1)[-1]
            if local in {"cm"}:
                t = "".join((x.text or "") for x in el.iter() if x.tag.rsplit("}", 1)[-1] in {"text", "t"})
                if t.strip():
                    texts.append(t.strip())
        if not texts:
            t = "".join((x.text or "") for x in root.iter() if x.tag.rsplit("}", 1)[-1] in {"text", "t"})
            if t.strip():
                texts.append(t.strip())
        if texts:
            out.setdefault(part_to_slide.get(n, ""), []).extend(texts)
    return out


def inspect_pptx(path: Path, opt: dict) -> Report:
    from pptx import Presentation

    rep = Report(str(path))
    with zipfile.ZipFile(path) as z:
        _check_core(rep, _core_props_from_zip(z), opt)
        comments = _pptx_comments(z)

    prs = Presentation(str(path))
    sw, sh = prs.slide_width, prs.slide_height
    extra = opt.get("ng_words") or []
    company = (opt.get("company") or "").strip()
    body_all: list[str] = []
    labels: list[str] = []

    slide_file_to_no: dict[str, int] = {}
    for i, slide in enumerate(prs.slides, 1):
        try:
            slide_file_to_no[Path(slide.part.partname).stem] = i
        except Exception:
            pass

    for i, slide in enumerate(prs.slides, 1):
        loc = f"{i}枚目"
        if slide._element.get("show") == "0":
            rep.add("要確認", "非表示スライド", loc, "非表示スライドがあります。PDFに書き出すとき「非表示スライドを含める」をオフにしてください。PowerPointファイルを公開する場合は削除してください")

        if slide.has_notes_slide:
            try:
                notes = slide.notes_slide.notes_text_frame.text or ""
            except Exception:
                notes = ""
            if notes.strip():
                hits = _hit_ng(notes, extra)
                msg = "発表者ノートがあります。PDFはノートなしの形式で書き出してください。ノート付きで公開する場合は内容を確認してください"
                if hits:
                    msg += f"（ノートに要注意語: {'、'.join(hits)}）"
                rep.add("要確認", "発表者ノート", loc, msg, notes.strip()[:200])

        # スライドの外に置いた図形（トップレベルのみ）
        for shape in slide.shapes:
            try:
                l, t, w, h = shape.left, shape.top, shape.width, shape.height
            except Exception:
                continue
            if l is None or t is None:
                continue
            if (l + w <= 0) or (t + h <= 0) or (l >= sw) or (t >= sh):
                name = getattr(shape, "name", "") or "オブジェクト"
                rep.add("要確認", "スライド外オブジェクト", f"{loc} / {name}", "スライドの外に図形が置かれています。PowerPointファイルを開くと見えます")

        texts = [_pptx_shape_text(s) for s in _iter_pptx_shapes(slide.shapes)]
        body = _nfkc("\n".join(t for t in texts if t))
        body_all.append(body)
        labels.append(loc)
        hits = _hit_ng(body, extra)
        if hits:
            sev = "要修正" if any(h in BODY_STRONG_LABELS for h in hits) else "要確認"
            rep.add(sev, "本文", loc, f"スライドに要注意表現があります（{'、'.join(hits)}）")

    for slide_file, items in comments.items():
        no = slide_file_to_no.get(slide_file)
        loc = f"{no}枚目" if no else "スライド"
        for t in items:
            rep.add("要修正", "コメント", loc, "PowerPointのコメントが残っています。「校閲 ＞ コメント」から削除してください", t[:200])

    body = "\n".join(body_all)
    check_marks_and_review(rep, body, body_all, labels)
    if company:
        seen = set()
        for m in COMPANY_RE.finditer(body):
            name = (m.group(0) or "").replace(" ", "")
            if _nfkc(company).replace(" ", "") in name:
                continue
            seen.add(name)
        if seen:
            rep.add("情報", "他社名", "スライド", "自社以外の「株式会社」名称があります（テンプレ流用や他社ロゴ周辺の社名残がないか確認）", "、".join(list(seen)[:8]))
    return rep


def inspect_xlsx(path: Path, opt: dict) -> Report:
    from openpyxl import load_workbook

    rep = Report(str(path))
    with zipfile.ZipFile(path) as z:
        _check_core(rep, _core_props_from_zip(z), opt)

    wb = load_workbook(str(path), data_only=False)
    extra = opt.get("ng_words") or []
    body_parts: list[str] = []
    labels: list[str] = []
    for ws in wb.worksheets:
        if ws.sheet_state and ws.sheet_state != "visible":
            rep.add("要修正", "非表示シート", ws.title, f"シートが非表示です（state={ws.sheet_state}）")
        for row in ws.iter_rows():
            for cell in row:
                loc = f"{ws.title}!{cell.coordinate}"
                if cell.comment and cell.comment.text:
                    rep.add("要修正", "コメント", loc, "セルコメントが残っています", cell.comment.text[:200])
                if isinstance(cell.value, str):
                    v = _nfkc(cell.value)
                    hits = _hit_ng(v, extra)
                    if hits:
                        rep.add("要確認", "本文", loc, f"セルに要注意表現があります（{'、'.join(hits)}）", v[:80])
                    body_parts.append(v)
                    labels.append(loc)
    check_marks_and_review(rep, "\n".join(body_parts), body_parts, labels)
    return rep


# ---------------------------------------------------------------------------
# メタデータ清掃（コピー作成）
# ---------------------------------------------------------------------------
def _pdf_date_string(d: date) -> str:
    return f"D:{d.year:04d}{d.month:02d}{d.day:02d}000000+09'00'"


def clean_metadata_copy(src: Path, dst: Path, disclosure_date: date | None = None) -> None:
    """文書情報とXMPを消した公開用コピーを作る。原本は変更しない。"""
    import pymupdf

    if dst.resolve() == src.resolve():
        raise SystemExit("原本と同じパスには書き出せません。別の出力先を指定してください。")
    doc = pymupdf.open(src)
    meta = {k: "" for k in ("title", "author", "subject", "keywords", "creator", "producer", "creationDate", "modDate")}
    if disclosure_date:
        stamp = _pdf_date_string(disclosure_date)
        meta["creationDate"] = stamp
        meta["modDate"] = stamp
    doc.set_metadata(meta)
    try:
        doc.del_xml_metadata()
    except Exception:
        pass
    tmp = dst.with_suffix(dst.suffix + ".tmp")
    doc.save(tmp, garbage=4, deflate=True)
    doc.close()
    tmp.replace(dst)


# ---------------------------------------------------------------------------
# レポート出力
# ---------------------------------------------------------------------------
NO_FINDING_LABEL = "機械チェックでの指摘なし（目視確認は未了）"
DISCLAIMER = (
    "このチェックは機械で拾える項目だけを見ています。指摘がないことは、開示してよいことの確認にはなりません。\n"
    "グラフと数値の一致、他社ロゴ、中期経営計画との整合、画像化された文字は、manual_checklist.md で目視確認してください。"
)


def _overall(reports: list[Report]) -> str | None:
    worst = None
    for r in reports:
        if r.worst is not None and (worst is None or SEVERITY_ORDER[r.worst] < SEVERITY_ORDER[worst]):
            worst = r.worst
    return worst


def format_markdown(reports: list[Report]) -> str:
    lines = [f"# IR提出前チェック結果（ir-precheck {__version__}）", ""]
    worst = _overall(reports)
    if worst == "未検査":
        lines.append("- 総合判定: 検査できなかったファイルがあります（下の「未検査」を参照）")
    elif worst in ("要修正", "要確認"):
        lines.append(f"- 総合判定: {worst}あり")
    else:
        lines.append(f"- 総合判定: {NO_FINDING_LABEL}")
    lines.append("")
    for r in reports:
        lines.append(f"## {Path(r.path).name}")
        lines.append("")
        if r.info.get("pages") is not None:
            lines.append(f"- ページ数: {r.info['pages']}")
        if r.info.get("metadata"):
            lines.append("- 文書のプロパティ:")
            for k, v in r.info["metadata"].items():
                if v:
                    lines.append(f"  - {_lab(k)}: {v}")
        lines.append("")
        if not r.findings:
            lines.append(f"{NO_FINDING_LABEL}。")
            lines.append("")
            continue
        groups: dict[str, list[Finding]] = {}
        for f in r.findings:
            groups.setdefault(f.severity, []).append(f)
        for sev in SEVERITIES:
            items = groups.get(sev) or []
            if not items:
                continue
            lines.append(f"### {sev}（{len(items)}件）")
            lines.append("")
            for f in items:
                loc = f" {f.location}:" if f.location else ""
                lines.append(f"- [{f.category}]{loc} {f.message}")
                if f.detail:
                    detail = f.detail.replace("\n", " ").strip()
                    lines.append(f"  - 詳細: {detail}")
            lines.append("")
    lines.append("---")
    lines.append(DISCLAIMER)
    return "\n".join(lines) + "\n"


def _unchecked(path: Path, message: str, detail: str = "") -> Report:
    rep = Report(str(path))
    rep.add("未検査", "ファイル", "全体", message, detail)
    return rep


def inspect(path: Path, opt: dict) -> Report:
    """1ファイルを検査する。失敗しても例外を外に出さず「未検査」で返す。"""
    if not path.is_file():
        return _unchecked(path, "ファイルが見つからないため検査していません")
    suf = path.suffix.lower()
    funcs = {".pdf": inspect_pdf, ".docx": inspect_docx, ".pptx": inspect_pptx, ".pptm": inspect_pptx, ".xlsx": inspect_xlsx, ".xlsm": inspect_xlsx}
    func = funcs.get(suf)
    if func is None:
        hint = "（.doc・.ppt・.xls は .docx・.pptx・.xlsx で保存し直すと検査できます）" if suf in {".doc", ".ppt", ".xls"} else ""
        return _unchecked(path, f"未対応の形式（{suf or '拡張子なし'}）のため検査していません{hint}")
    try:
        return func(path, opt)
    except ImportError as e:
        return _unchecked(path, f"必要なライブラリがないため検査していません。{INSTALL_HINT} を実行してください", str(e))
    except zipfile.BadZipFile:
        return _unchecked(path, "パスワード付き、または壊れたOfficeファイルのため開けません。パスワードを外したファイルで実行し直してください")
    except Exception as e:  # noqa: BLE001
        return _unchecked(path, "検査中にエラーが起きたため、このファイルは検査を終えていません", f"{type(e).__name__}: {str(e)[:120]}")


class _Parser(argparse.ArgumentParser):
    """引数の誤りは終了コード3（argparse の既定の2は「要修正」と区別できないため）。"""

    def error(self, message: str):  # type: ignore[override]
        self.print_usage(sys.stderr)
        print(f"引数の誤り: {message}", file=sys.stderr)
        raise SystemExit(3)


def _utf8_stdio() -> None:
    """Windows の既定（cp1252・cp932）でも日本語の結果が途中で止まらないよう、出力をUTF-8にする。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
        except Exception:
            pass


def main(argv: list[str] | None = None) -> int:
    _utf8_stdio()
    p = _Parser(description=f"IR開示資料の提出前チェッカー（ir-precheck {__version__}）")
    p.add_argument("files", nargs="+", help="PDF / DOCX / PPTX / XLSX")
    p.add_argument("--company", help="自社名（例: 株式会社〇〇 / 〇〇株式会社）")
    p.add_argument("--code", help="自社の証券コード（例: 8783 / 462A）")
    p.add_argument("--period", help="対象決算期（例: 2027年3月期）")
    p.add_argument("--date", dest="disclosure_date", help="開示日（例: 2026-11-13 / 2026年11月13日）。作成日時との比較に使う")
    p.add_argument("--ng-words", dest="ng_words", help="追加NG語（カンマ区切り、またはファイルパス）")
    p.add_argument("--redact-words", dest="redact_words", help="墨消ししたはずの語（カンマ区切り、またはファイルパス）。自社ファイルにだけ使う")
    p.add_argument("--json", action="store_true", help="JSONで出力")
    p.add_argument("--clean-metadata", metavar="OUT.pdf", help="文書のプロパティを消した公開用コピーを作成（原本は変更しない）。--date があるときは作成・更新日時もその日にそろえる")
    p.add_argument("--version", action="version", version=f"ir-precheck {__version__}")
    args = p.parse_args(argv)

    disc = parse_user_date(args.disclosure_date) if args.disclosure_date else None
    if args.disclosure_date and not disc:
        print(f"開示日の形式が読めません: {args.disclosure_date}", file=sys.stderr)
        return 3

    opt = {
        "company": _nfkc(args.company or ""),
        "code": args.code or "",
        "period": _nfkc(args.period or ""),
        "ng_words": _read_ng_list(args.ng_words),
        "redact_words": _read_ng_list(args.redact_words),
        "disclosure_date": disc,
    }

    paths = [Path(f) for f in args.files]

    if args.clean_metadata:
        if len(paths) != 1 or paths[0].suffix.lower() != ".pdf" or not paths[0].is_file():
            print("--clean-metadata はPDFを1つだけ指定してください", file=sys.stderr)
            return 3
        out = Path(args.clean_metadata)
        clean_metadata_copy(paths[0], out, disclosure_date=disc)
        print(f"公開用コピーを作成しました: {out}（このファイルにも通常のチェックを実行してください）", file=sys.stderr)

    reports = [inspect(path, opt) for path in paths]
    pdf_reports = [r for r, path in zip(reports, paths) if path.suffix.lower() == ".pdf" and r.worst != "未検査"]
    if len(pdf_reports) >= 2:
        check_multi_pdf_dates(pdf_reports)

    worst = _overall(reports)
    if args.json:
        payload = {
            "tool": "ir-precheck",
            "version": __version__,
            "worst": worst,
            "unchecked": [r.path for r in reports if r.worst == "未検査"],
            "reports": [
                {"path": r.path, "info": r.info, "worst": r.worst, "findings": [f.to_dict() for f in r.findings]}
                for r in reports
            ],
            "note": DISCLAIMER,
        }
        json.dump(payload, sys.stdout, ensure_ascii=False, indent=2)
        sys.stdout.write("\n")
    else:
        sys.stdout.write(format_markdown(reports))

    return {"未検査": 3, "要修正": 2, "要確認": 1}.get(worst or "", 0)


if __name__ == "__main__":
    sys.exit(main())
