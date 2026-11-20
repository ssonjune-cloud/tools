# -*- coding: utf-8 -*-
"""
試験問題の docx の体裁を整えるスクリプト（ルビの有無を問わず使えます）

使い方：
    python mondai_shoshiki.py 試験問題.docx
    （複数ファイルも可。元のファイルは残し、「〜_整形.docx」を新しく作ります）
    ルビを振る場合は、ルビを振ったあとに通してください。

やること（最初の大問の見出しより後ろだけ。表紙・注意には触りません）：
  1. 行間を固定値23ptにそろえる（ルビが振ってある文書は28pt。図の入った段落は除く）
  2. 設問のまわりのインデントと間隔をそろえる
       問の文 …………… ぶら下げ（「問１　」の後ろで折り返しがそろうよう自動で合わせる）
       選択肢 …………… 字下げ2字＋ぶら下げ2字（「ア　」の後ろでそろう）
       ａｂｃ・Ⅰ〜Ⅳ … 選択肢と同じ
       小問ⅰ・ⅱ ……… 字下げ1字＋ぶら下げ2字
       空欄補充の文など … 字下げ2字
       間隔 ……………… 設問の中は0.2行、設問の最後の行の後は0.5行
     ※選択肢・項目・小問の行頭の全角スペースは消して、インデントに置き換えます
  3. 設問とその選択肢が、ページや段の途中で分かれないようにする

触らないもの：用紙・段組み・セクション、文字の大きさや太字、表・テキストボックスの中
"""
import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Pt, Twips
from docx.text.paragraph import Paragraph

# ===== 設定 =====
KIHON_PT = 23          # 基本の行間（固定値 pt）。変更しないなら None
RUBY_PT = 28           # ルビが振ってある文書の行間（固定値 pt）。自動で切り替えます

# インデント（字数）：(左の字下げ, ぶら下げ)
IND_TOI = "auto"       # 問の文：ぶら下げを「問１　」の幅に自動で合わせる（例：(0, 4) で固定）
IND_SENTAKU = (4, 2)   # 選択肢（1行目は2字下がり、折り返しは4字目から）
IND_KOUMOKU = (4, 2)   # ａｂｃ・Ⅰ〜Ⅳ などの項目
IND_SHOMON = (3, 2)    # 小問ⅰ・ⅱ（1行目は1字下がり）
IND_SONOTA = (2, 0)    # 空欄補充の文など

AIDA = 0.2             # 設問の中の段落の後の間隔（行）
OWARI = 0.5            # 設問の最後の行の後の間隔（行）

KUUHAKU_KESU = True    # 選択肢・項目・小問の行頭の空白を消す
MATAGANAI = True       # 設問と選択肢をページ・段の途中で分けない

MOJI_PT = 11           # 本文の文字の大きさ（字数→長さの換算に使う）
# ================

RE_DAIMON = re.compile(r"^[\s　]*[１-９0-9]{1,2}[\s　]*次")      # 大問の見出し
RE_TOI = re.compile(r"^[\s　]*問[\s　]*[０-９0-9]+")              # 問１、問１０
RE_TOI_LABEL = re.compile(r"^問[\s　]*[０-９0-9]+[\s　]*")        # 「問１　」の部分
RE_SENTAKU = re.compile(r"^[\s　]*[アイウエオカ][\s　]")          # 選択肢
RE_SHOMON = re.compile(r"^[\s　]*[ⅰⅱⅲⅳⅴ]")                     # 小問
RE_KOUMOKU = re.compile(r"^[\s　]*[ａ-ｅＡ-ＥⅠ-Ⅵ][\s　]")         # 項目

CHAR_TW = MOJI_PT * 20  # 1字の長さ（twip）


# ---------- 文字列の取り出し（ルビの読みは除く） ----------
def in_ruby_text(el):
    return any(a.tag == qn("w:rt") for a in el.iterancestors())


def ptext(par):
    return "".join(t.text or "" for t in par._p.iter(qn("w:t")) if not in_ruby_text(t))


def strip_leading(par):
    for t in par._p.iter(qn("w:t")):
        if in_ruby_text(t):
            continue
        s = (t.text or "").lstrip(" 　\t")
        t.text = s
        if s:
            break


# ---------- 書式 ----------
def has_drawing(par):
    p = par._p
    return any(True for _ in p.iter(qn("w:drawing"))) or any(True for _ in p.iter(qn("w:pict")))


def is_numbered(par):
    pPr = par._p.pPr
    return pPr is not None and pPr.find(qn("w:numPr")) is not None


def is_section_break(par):
    pPr = par._p.pPr
    return pPr is not None and pPr.find(qn("w:sectPr")) is not None


def has_ruby(body):
    """ルビ（ルビ機能・EQフィールドのどちらでも）があるか"""
    if any(True for _ in body.iter(qn("w:ruby"))):
        return True
    for it in body.iter(qn("w:instrText")):
        if "EQ" in (it.text or "").upper() and "\\*" in (it.text or ""):
            return True
    return False


def set_line(par, pt):
    if pt is None or has_drawing(par):
        return
    pf = par.paragraph_format
    pf.line_spacing = Pt(pt)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY


def set_after(par, lines):
    pf = par.paragraph_format
    pf.space_after = Pt(0)  # w:spacing を確実に作る
    sp = par._p.pPr.find(qn("w:spacing"))
    sp.set(qn("w:afterLines"), str(int(lines * 100)))
    sp.set(qn("w:after"), "0")


def drop_indent_style(par):
    """「リスト段落」のように字下げだけを決めているスタイルは外す。
    （スタイル側の字数指定が残ると、ソフトによっては字下げが効かないため）"""
    pPr = par._p.pPr
    ps = pPr.find(qn("w:pStyle")) if pPr is not None else None
    if ps is None:
        return
    st = par.part.styles.element.get_by_id(ps.get(qn("w:val")))
    if st is None:
        return
    spPr = st.find(qn("w:pPr"))
    only_ind = spPr is not None and all(c.tag == qn("w:ind") for c in spPr)
    if only_ind and st.find(qn("w:rPr")) is None:
        pPr.remove(ps)


def set_indent(par, left, hang):
    """left・hang は字数（11ptの文字1字分を単位とする。0.5字も可）"""
    drop_indent_style(par)
    pf = par.paragraph_format
    pf.left_indent = Twips(round(left * CHAR_TW))
    pf.first_line_indent = Twips(-round(hang * CHAR_TW)) if hang else Twips(0)
    ind = par._p.pPr.find(qn("w:ind"))
    for k in list(ind.attrib):
        if k not in (qn("w:left"), qn("w:hanging"), qn("w:firstLine")):
            del ind.attrib[k]
    # 字数指定（Chars）は 0 にして、長さ（twip）の指定を使わせる
    #  ※文書のグリッドの都合で「1字」が実際の文字より小さくなるのを避けるため
    ind.set(qn("w:leftChars"), "0")
    if hang:
        ind.set(qn("w:hangingChars"), "0")
    else:
        ind.set(qn("w:firstLineChars"), "0")
        ind.set(qn("w:firstLine"), "0")


def label_width(text, pattern):
    """行頭の「問１　」などの幅（字数）。半角は0.5字として数える"""
    m = pattern.match(text)
    if not m:
        return 0
    return sum(0.5 if ord(c) < 0x2000 else 1 for c in m.group(0))


# ---------- 設問のまとまりを見つける ----------
def find_blocks(pars):
    """設問ごとに [(段落, 種類), ...] のまとまりを返す。
    種類：toi / sentaku / koumoku / shomon / sonota"""
    blocks, cur = [], []

    def close():
        nonlocal cur
        if cur:
            # 番号の自動付け（行番号など）で始まるまとまりは、選択肢があるときだけ設問とみなす
            numbered = cur[0][1] == "toi_num"
            if not numbered or any(k == "sentaku" for _, k in cur):
                blocks.append(cur)
        cur = []

    def seen_choice():
        return any(k == "sentaku" for _, k in cur)

    started = False
    for par in pars:
        t = ptext(par)
        if RE_DAIMON.match(t):
            started = True
            close()
            continue
        if not started:
            continue
        if RE_TOI.match(t):
            close()
            cur = [(par, "toi")]
        elif is_numbered(par) and not RE_SENTAKU.match(t):
            close()
            cur = [(par, "toi_num")]
        elif not cur:
            pass  # 設問の外（本文など）
        elif t.strip() == "":
            close()
        elif RE_SENTAKU.match(t):
            cur.append((par, "sentaku"))
        elif RE_SHOMON.match(t):
            cur.append((par, "shomon"))
        elif seen_choice() and cur[-1][1] == "sentaku":
            close()  # 選択肢の後に続く、設問以外の段落
        elif RE_KOUMOKU.match(t):
            cur.append((par, "koumoku"))
        else:
            cur.append((par, "sonota"))
        if cur and is_section_break(par):
            close()
    close()
    return blocks


def keep_ranges(block):
    """ページ・段をまたがせない範囲（問または小問から、その最後の選択肢まで）"""
    ranges, start = [], 0
    for i, (_, k) in enumerate(block):
        if k == "shomon" and i > 0:
            # 問の文のすぐ後の小問は、問の文とつなげる
            if not (i == 1 and block[0][1].startswith("toi")):
                ranges.append(block[start:i])
                start = i
    ranges.append(block[start:])
    out = []
    for r in ranges:
        last = max((i for i, (_, k) in enumerate(r) if k == "sentaku"), default=-1)
        if last >= 1:
            out.append([p for p, _ in r[: last + 1]])
    return out


def format_block(block):
    for i, (par, kind) in enumerate(block):
        if kind in ("sentaku", "koumoku", "shomon", "toi") and KUUHAKU_KESU:
            strip_leading(par)
        if kind == "toi":
            if IND_TOI == "auto":
                w = label_width(ptext(par), RE_TOI_LABEL) or 3
                set_indent(par, w, w)
            else:
                set_indent(par, *IND_TOI)
        elif kind == "sentaku":
            set_indent(par, *IND_SENTAKU)
        elif kind == "koumoku":
            set_indent(par, *IND_KOUMOKU)
        elif kind == "shomon":
            set_indent(par, *IND_SHOMON)
        elif kind == "sonota":
            set_indent(par, *IND_SONOTA)
        # toi_num（自動番号の設問）はインデントを番号の設定に任せる
        last = i == len(block) - 1
        nxt_shomon = not last and block[i + 1][1] == "shomon" and i > 0
        set_after(par, OWARI if (last or nxt_shomon) else AIDA)


def apply_keep(rng):
    for i, par in enumerate(rng):
        pf = par.paragraph_format
        pf.keep_together = True
        pf.keep_with_next = i < len(rng) - 1


def process(path: Path, out: Path = None):
    doc = Document(path)
    body = doc.element.body
    pars = [Paragraph(p, doc) for p in body if p.tag == qn("w:p")]

    # 1. 行間（最初の大問の見出しから後ろ）
    ruby = has_ruby(body)
    pt = RUBY_PT if (ruby and KIHON_PT is not None) else KIHON_PT
    started = False
    for par in pars:
        if not started and RE_DAIMON.match(ptext(par)):
            started = True
        if started:
            set_line(par, pt)

    # 2・3. 設問のまわり
    blocks = find_blocks(pars)
    n_keep = 0
    for b in blocks:
        format_block(b)
        if MATAGANAI:
            for r in keep_ranges(b):
                apply_keep(r)
                n_keep += 1

    out = out or path.with_name(path.stem + "_整形.docx")
    doc.save(out)
    gyokan = f"行間 {pt}pt" + ("（ルビあり）" if ruby else "") if pt else "行間そのまま"
    print(f"{path.name} → {out.name}（{gyokan}、設問 {len(blocks)} か所、分けないまとまり {n_keep} か所）")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("docx ファイルを指定してください")
        sys.exit(1)
    for a in sys.argv[1:]:
        process(Path(a))
