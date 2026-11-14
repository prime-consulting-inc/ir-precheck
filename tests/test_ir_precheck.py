"""ir_precheck.py の回帰テスト。フィクスチャはテストのたびに一時フォルダへ作る。

実行: python -m pytest -q
"""
from __future__ import annotations

import json
import subprocess
import sys
import zipfile
from pathlib import Path

import pymupdf
import pytest

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "skills" / "ir-precheck" / "scripts" / "ir_precheck.py"
sys.path.insert(0, str(SCRIPT.parent))
import ir_precheck as ip  # noqa: E402


def run(*args: str) -> tuple[int, dict]:
    proc = subprocess.run(
        [sys.executable, "-I", str(SCRIPT), "--json", *args],
        capture_output=True,
        text=True,
    )
    return proc.returncode, json.loads(proc.stdout) if proc.stdout.strip() else {}


def findings(payload: dict, i: int = 0) -> list[dict]:
    return payload["reports"][i]["findings"]


def cats(payload: dict, i: int = 0) -> dict[str, str]:
    return {f["category"]: f["severity"] for f in findings(payload, i)}


def make_pdf(path: Path, lines: list[str], meta: dict | None = None, extra=None) -> Path:
    doc = pymupdf.open()
    page = doc.new_page()
    y = 60
    for ln in lines:
        page.insert_text((50, y), ln, fontname="japan", fontsize=11)
        y += 20
    if extra:
        extra(page)
    m = {"title": "", "author": "", "creator": "", "producer": ""}
    m.update(meta or {})
    doc.set_metadata(m)
    doc.save(path)
    return path


# ---------------------------------------------------------------- 黒塗り・墨消し


def test_partial_black_box_is_detected_without_revealing_text(tmp_path):
    def blk(page):
        page.insert_text((50, 200), "取引先 竹中商事 との契約", fontname="japan", fontsize=11)
        page.draw_rect(pymupdf.Rect(90, 188, 150, 204), color=None, fill=(0, 0, 0))

    pdf = make_pdf(tmp_path / "a.pdf", ["2026年11月13日"], extra=blk)
    code, out = run(str(pdf))
    assert code == 2
    assert cats(out)["黒塗りの残存"] == "要修正"
    dumped = json.dumps(out, ensure_ascii=False)
    assert "竹中" not in dumped


def test_white_text_on_black_band_is_not_flagged(tmp_path):
    def band(page):
        page.draw_rect(pymupdf.Rect(40, 180, 300, 210), color=None, fill=(0, 0, 0))
        page.insert_text((50, 200), "業績ハイライト", fontname="japan", fontsize=11, color=(1, 1, 1))

    pdf = make_pdf(tmp_path / "b.pdf", ["2026年11月13日"], extra=band)
    _, out = run(str(pdf))
    assert "黒塗りの残存" not in cats(out)


def test_unapplied_redaction_annotation_is_must_fix(tmp_path):
    def red(page):
        page.insert_text((50, 200), "取引先 竹中商事 との契約", fontname="japan", fontsize=11)
        page.add_redact_annot(pymupdf.Rect(90, 188, 150, 204), fill=(0, 0, 0))

    pdf = make_pdf(tmp_path / "c.pdf", ["2026年11月13日"], extra=red)
    code, out = run(str(pdf))
    assert code == 2
    assert cats(out)["墨消し未適用"] == "要修正"


# ---------------------------------------------------------------- 普通の開示での誤検知


def test_correction_notice_is_not_flagged(tmp_path):
    pdf = make_pdf(
        tmp_path / "teisei.pdf",
        [
            "（訂正）「2027年3月期 第2四半期決算短信」の一部訂正について",
            "訂正箇所に誤りがありましたので、資料を差し替えております。",
            "ご迷惑をおかけし申し訳ございません。",
        ],
        {"author": "経営企画部"},
    )
    code, out = run("--period", "2027年3月期", str(pdf))
    assert code == 0, findings(out)


def test_officer_label_and_default_author_are_not_flagged(tmp_path):
    pdf = make_pdf(
        tmp_path / "yakuin.pdf",
        ["取締役（社内） 鈴木一郎", "監査等委員（社内） 田中花子", "新規出店は控え、既存店の改装を優先します。"],
        {"author": "Windows User"},
    )
    code, out = run(str(pdf))
    assert code == 0, findings(out)


def test_draft_in_body_is_review_not_must_fix(tmp_path):
    pdf = make_pdf(tmp_path / "beer.pdf", ["生ドラフトビールの販売が好調に推移しました。"])
    code, out = run(str(pdf))
    assert code == 1
    assert cats(out)["本文"] == "要確認"


def test_draft_in_title_is_must_fix(tmp_path):
    pdf = make_pdf(tmp_path / "t.pdf", ["2026年11月13日"], {"title": "決算説明資料（案）_v8"})
    code, out = run(str(pdf))
    assert code == 2
    assert cats(out)["メタデータ"] == "要修正"


def test_internal_instruction_is_still_must_fix(tmp_path):
    pdf = make_pdf(tmp_path / "r.pdf", ["この表は差し替えてください", "社外秘"])
    code, out = run(str(pdf))
    assert code == 2
    assert cats(out)["レビュー文言"] == "要修正"
    assert cats(out)["社外秘マーク"] == "要修正"


# ---------------------------------------------------------------- 日付・正規化


def test_weekday_mismatch(tmp_path):
    pdf = make_pdf(tmp_path / "w.pdf", ["説明会は2026年11月13日（木）に開催します。"])
    _, out = run(str(pdf))
    assert cats(out)["日付と曜日"] == "要修正"


def test_kangxi_radicals_are_normalized():
    text = ip._nfkc("2026年11⽉13⽇（⽊）")  # 康熙部首の月・日・木
    m = ip.WEEKDAY_RE.search(text)
    assert m and m.group("w") == "木"
    assert ip.SECRET_STRONG[0][1].search(ip._nfkc("社外⽐")) is None  # 別字は拾わない
    assert ip._hit_ng("決算短信（案）") == ip._hit_ng("決算短信(案)")


# ---------------------------------------------------------------- セキュリティ


def test_print_copy_restriction(tmp_path):
    doc = pymupdf.open()
    doc.new_page().insert_text((50, 80), "テスト", fontname="japan")
    path = tmp_path / "perm.pdf"
    doc.save(path, encryption=pymupdf.PDF_ENCRYPT_AES_256, owner_pw="o", permissions=pymupdf.PDF_PERM_ACCESSIBILITY)
    code, out = run(str(path))
    assert code == 2
    assert cats(out)["セキュリティ"] == "要修正"


def test_user_password_pdf(tmp_path):
    doc = pymupdf.open()
    doc.new_page().insert_text((50, 80), "テスト", fontname="japan")
    path = tmp_path / "pw.pdf"
    doc.save(path, encryption=pymupdf.PDF_ENCRYPT_AES_256, owner_pw="o", user_pw="u")
    code, out = run(str(path))
    assert code == 2
    assert cats(out)["セキュリティ"] == "要修正"


# ---------------------------------------------------------------- 未検査の扱い


def test_unreadable_office_file_is_unchecked_and_others_still_run(tmp_path):
    broken = tmp_path / "locked.docx"
    broken.write_bytes(b"\xd0\xcf\x11\xe0 not a zip")  # パスワード付きOfficeはCFB形式でzipではない
    pdf = make_pdf(tmp_path / "ok.pdf", ["社外秘"])
    code, out = run(str(broken), str(pdf))
    assert code == 3
    assert out["worst"] == "未検査"
    assert out["unchecked"] == [str(broken)]
    assert cats(out, 1)["社外秘マーク"] == "要修正"


def test_missing_and_unsupported_files_are_unchecked(tmp_path):
    old = tmp_path / "old.ppt"
    old.write_bytes(b"x")
    code, out = run(str(tmp_path / "nothing.pdf"), str(old))
    assert code == 3
    assert len(out["unchecked"]) == 2


def test_bad_arguments_exit_3(tmp_path):
    pdf = make_pdf(tmp_path / "x.pdf", ["テスト"])
    proc = subprocess.run([sys.executable, "-I", str(SCRIPT), "--date", "あした", str(pdf)], capture_output=True, text=True)
    assert proc.returncode == 3
    proc = subprocess.run([sys.executable, "-I", str(SCRIPT), "--no-such-option", str(pdf)], capture_output=True, text=True)
    assert proc.returncode == 3


# ---------------------------------------------------------------- Office


def test_pptx_table_group_comment_notes(tmp_path):
    from pptx import Presentation
    from pptx.util import Inches

    prs = Presentation()
    s = prs.slides.add_slide(prs.slide_layouts[6])
    tb = s.shapes.add_table(2, 2, Inches(1), Inches(1), Inches(4), Inches(1)).table
    tb.cell(0, 0).text = "社外秘"
    grp = s.shapes.add_group_shape()
    grp.shapes.add_textbox(Inches(1), Inches(3), Inches(3), Inches(1)).text_frame.text = "TODO 数字は後で"
    s.notes_slide.notes_text_frame.text = "ここは口頭で補足"
    prs.core_properties.last_modified_by = "Koichi I."
    src = tmp_path / "src.pptx"
    prs.save(src)

    # 旧形式のコメントを1枚目にひも付けて足す
    with zipfile.ZipFile(src) as z:
        items = {n: z.read(n) for n in z.namelist()}
    items["ppt/comments/comment1.xml"] = (
        '<?xml version="1.0" encoding="UTF-8"?><p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:cm authorId="0" idx="1"><p:pos x="10" y="10"/><p:text>ここの数字は未公表</p:text></p:cm></p:cmLst>'
    ).encode()
    rels = items["ppt/slides/_rels/slide1.xml.rels"].decode()
    rels = rels.replace(
        "</Relationships>",
        '<Relationship Id="rIdC1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comment1.xml"/></Relationships>',
    )
    items["ppt/slides/_rels/slide1.xml.rels"] = rels.encode()
    deck = tmp_path / "deck.pptx"
    with zipfile.ZipFile(deck, "w") as w:
        for n, b in items.items():
            w.writestr(n, b)

    code, out = run(str(deck))
    assert code == 2
    fs = findings(out)
    by = {(f["category"], f["severity"], f["location"]) for f in fs}
    assert ("社外秘マーク", "要修正", "1枚目") in by
    assert ("レビュー文言", "要修正", "1枚目") in by
    assert ("コメント", "要修正", "1枚目") in by
    assert ("発表者ノート", "要確認", "1枚目") in by


def test_xlsx_location_is_cell_address(tmp_path):
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "データ"
    for r in range(1, 40):
        ws.cell(r, 2, f"項目{r}")
    ws["B30"] = "社外秘"
    path = tmp_path / "data.xlsx"
    wb.save(path)
    _, out = run(str(path))
    locs = [f["location"] for f in findings(out) if f["category"] == "社外秘マーク"]
    assert locs == ["データ!B30"]


def test_docx_hidden_text_is_counted_not_shown(tmp_path):
    from docx import Document

    d = Document()
    p = d.add_paragraph("公表文。")
    run_ = p.add_run("未公表の取引先名")
    run_.font.hidden = True
    path = tmp_path / "h.docx"
    d.save(path)
    code, out = run(str(path))
    assert code == 2
    assert cats(out)["隠し文字"] == "要修正"
    assert "未公表の取引先名" not in json.dumps(out, ensure_ascii=False)


# ---------------------------------------------------------------- 公開用コピー


def test_clean_metadata_copy(tmp_path):
    src = make_pdf(tmp_path / "src.pdf", ["2026年11月13日"], {"title": "決算説明資料（案）", "author": "山田 太郎"})
    dst = tmp_path / "pub.pdf"
    subprocess.run([sys.executable, "-I", str(SCRIPT), "--clean-metadata", str(dst), str(src)], capture_output=True, check=False)
    assert dst.is_file()
    meta = pymupdf.open(dst).metadata
    assert not meta.get("title") and not meta.get("author")
    assert pymupdf.open(src).metadata["title"] == "決算説明資料（案）"  # 原本は変わらない
    code, out = run(str(dst))
    assert code == 0, findings(out)


if __name__ == "__main__":
    raise SystemExit(pytest.main([__file__, "-q"]))
