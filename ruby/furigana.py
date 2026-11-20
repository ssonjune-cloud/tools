# -*- coding: utf-8 -*-
"""
定期試験問題 .docx に総ルビを振る（Word 標準のルビ機能と同じ形式で挿入）

使い方:
    python furigana.py 入力.docx                 -> 入力_ルビ.docx を出力
    python furigana.py 入力.docx 出力.docx

特徴:
  - 書式（太字・下線・サイズ）はそのまま保持
  - 送り仮名を除いて漢字部分にだけルビを振る（書く -> 書(か)く）
  - 〔ルビ停止〕〔ルビ再開〕でルビを振らない範囲を指定できる（漢字の読み書き問題など）
  - 読み辞書.json に固有名詞・熟字訓の読みを登録できる（次回以降も効く）
  - 辞書には前後の文字ごと登録できる（例："何が": "なにが" → 何(なに)が）
  - 〔古文〕〔古文ここまで〕で挟んだ範囲だけ、古文の辞書を使える（Colab版）
"""
import json, re, sys, os
from docx import Document
from docx.oxml.ns import qn
from janome.tokenizer import Tokenizer

STOP_MARK, RESUME_MARK = "〔ルビ停止〕", "〔ルビ再開〕"
KOBUN_MARK, KOBUN_END = "〔古文〕", "〔古文ここまで〕"
MARKS = (STOP_MARK, RESUME_MARK, KOBUN_MARK, KOBUN_END)

# ===== 設定（ここだけ書き換えれば動きが変わります）=====
# --- ページ数を減らす ---
GRID_OFF   = True    # 行グリッドに合わせるのをやめる（これが一番効きます）
TIGHTEN    = True    # 段落の前後の間隔をゼロにする
RUBY_RATIO = 0.5     # ルビの大きさ（本文に対する比率）0.45 や 0.4 にすると行が縮みます
MARGIN_MM  = 0       # 上下左右の余白をミリで指定（0 なら元のまま。例：15）
DEL_BLANK  = False   # 空行を削除する（元の余白設計が崩れることがあるので既定は False）

# --- ページまたぎ ---
KEEP_LINES    = True   # 一つの段落が途中で切れないようにする
KEEP_TOGETHER = False  # 設問と選択肢を必ず同じページに置く（ページ数は増えます）
# =====================================================
DICT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "読み辞書.json")

KANJI = re.compile(r'[\u4e00-\u9fff\u3005\u3007\u303b]')
HAS_KANJI = lambda s: bool(KANJI.search(s))
REAL_KANJI = re.compile(r'[\u4e00-\u9fff]')
YOMI = re.compile(r'^[\u3041-\u309f\u30fc]+$')

def kata2hira(s):
    return ''.join(chr(ord(c) - 0x60) if 'ァ' <= c <= 'ヶ' else c for c in s)

def load_dict():
    if os.path.exists(DICT_PATH):
        with open(DICT_PATH, encoding='utf-8') as f:
            return json.load(f)
    return {}

def make_matcher(yomi):
    """辞書の見出し（漢字を含むもの）を長い順に探す正規表現"""
    keys = sorted((k for k in yomi if HAS_KANJI(k)), key=len, reverse=True)
    return re.compile('|'.join(map(re.escape, keys))) if keys else None

def segments(text, yomi, matcher, tk):
    """(表層, 読み or None) を順に返す。辞書にある語を優先し、残りは形態素解析"""
    pos = 0
    pieces = []
    if matcher:
        for m in matcher.finditer(text):
            if m.start() > pos:
                pieces.append((text[pos:m.start()], None))
            pieces.append((m.group(), yomi[m.group()]))
            pos = m.end()
    if pos < len(text):
        pieces.append((text[pos:], None))
    for chunk, rd in pieces:
        if rd is not None:
            yield chunk, rd, True
            continue
        # janome は前後の空白（全角スペースの字下げなど）を落とすので、自分で残す
        core = chunk.strip()
        if not core:
            yield chunk, None, False
            continue
        lead = chunk[:len(chunk) - len(chunk.lstrip())]
        trail = chunk[len(chunk.rstrip()):]
        if lead:
            yield lead, None, False
        got = 0
        for token in tk.tokenize(core):
            got += len(token.surface)
            r = token.reading if token.reading and token.reading != '*' else None
            yield token.surface, (kata2hira(r) if r else None), False
        if got != len(core):   # 念のため：取りこぼしがあれば元の文字列に戻す
            raise ValueError('tokenize mismatch: ' + core[:20])
        if trail:
            yield trail, None, False

# ---------- 送り仮名の切り分け ----------
def split_okurigana(surface, reading):
    """表層形と読みから [(部分文字列, ルビ or None), ...] を返す"""
    parts = [m for m in re.finditer(r'[\u4e00-\u9fff\u3005\u3007\u303b]+|[^\u4e00-\u9fff\u3005\u3007\u303b]+', surface)]
    pattern = ''
    for m in parts:
        seg = m.group()
        if HAS_KANJI(seg):
            pattern += '(.+?)'
        else:
            pattern += re.escape(kata2hira(seg))
    mm = re.fullmatch(pattern, reading)
    if not mm:
        return [(surface, reading)] if HAS_KANJI(surface) else [(surface, None)]
    out, gi = [], 0
    for m in parts:
        seg = m.group()
        if HAS_KANJI(seg):
            gi += 1
            out.append((seg, mm.group(gi)))
        else:
            out.append((seg, None))
    return out

# ---------- ルビ XML ----------
def make_ruby(base_text, ruby_text, rPr):
    from copy import deepcopy
    from docx.oxml import OxmlElement
    sz = 21
    if rPr is not None:
        szel = rPr.find(qn('w:sz'))
        if szel is not None:
            try: sz = int(szel.get(qn('w:val')))
            except Exception: pass
    r = OxmlElement('w:r')
    if rPr is not None: r.append(deepcopy(rPr))
    ruby = OxmlElement('w:ruby')
    pr = OxmlElement('w:rubyPr')
    rsz = max(2, int(sz * RUBY_RATIO))
    for tag, val in (('w:rubyAlign', 'distributeSpace'), ('w:hps', str(rsz)),
                     ('w:hpsRaise', str(sz + 1)), ('w:hpsBaseText', str(sz)), ('w:lid', 'ja-JP')):
        e = OxmlElement(tag); e.set(qn('w:val'), val); pr.append(e)
    ruby.append(pr)
    for tag, txt, size in (('w:rt', ruby_text, rsz), ('w:rubyBase', base_text, sz)):
        holder = OxmlElement(tag)
        rr = OxmlElement('w:r')
        if rPr is not None:
            p2 = deepcopy(rPr)
            for t in ('w:sz', 'w:szCs'):
                for old in p2.findall(qn(t)): p2.remove(old)
            for t in ('w:sz', 'w:szCs'):
                e = OxmlElement(t); e.set(qn('w:val'), str(size)); p2.append(e)
            rr.append(p2)
        t = OxmlElement('w:t'); t.set(qn('xml:space'), 'preserve'); t.text = txt
        rr.append(t); holder.append(rr); ruby.append(holder)
    r.append(ruby)
    return r

def plain_run(text, rPr):
    from copy import deepcopy
    from docx.oxml import OxmlElement
    r = OxmlElement('w:r')
    if rPr is not None: r.append(deepcopy(rPr))
    t = OxmlElement('w:t'); t.set(qn('xml:space'), 'preserve'); t.text = text
    r.append(t)
    return r

# ---------- ページまたぎ対策 ----------
CHOICE = re.compile(r'^[\s\u3000]*[アイウエオカ][\s\u3000　]')
QUESTION = re.compile(r'^[\s\u3000]*(問|＊|条件|・)')

def get_pPr(p):
    from docx.oxml import OxmlElement
    pPr = p.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr'); p.insert(0, pPr)
    return pPr

def set_flag(p, tag):
    from docx.oxml import OxmlElement
    pPr = get_pPr(p)
    if pPr.find(qn(tag)) is None:
        pPr.append(OxmlElement(tag))

def fix_line_spacing(p):
    """行間が「固定値」だとルビが上で切れるので「最小値」に変える"""
    pPr = get_pPr(p)
    sp = pPr.find(qn('w:spacing'))
    if sp is not None and sp.get(qn('w:lineRule')) == 'exact':
        sp.set(qn('w:lineRule'), 'atLeast')


def base_text(p):
    """ルビの読み(w:rt)を除いた本文だけを取り出す"""
    out = []
    for t in p.iter(qn('w:t')):
        anc, skip = t.getparent(), False
        while anc is not None:
            if anc.tag == qn('w:rt'): skip = True; break
            anc = anc.getparent()
        if not skip: out.append(t.text or '')
    return ''.join(out)

def compact(paras):
    """ページ数を減らすための段落単位の調整"""
    from docx.oxml import OxmlElement
    for p in paras:
        pPr = get_pPr(p)
        if GRID_OFF and pPr.find(qn('w:snapToGrid')) is None:
            e = OxmlElement('w:snapToGrid'); e.set(qn('w:val'), '0'); pPr.append(e)
        if TIGHTEN:
            sp = pPr.find(qn('w:spacing'))
            if sp is None:
                sp = OxmlElement('w:spacing'); pPr.append(sp)
            sp.set(qn('w:before'), '0'); sp.set(qn('w:after'), '0')
            for k in ('w:beforeLines', 'w:afterLines'):
                if sp.get(qn(k)) is not None: sp.set(qn(k), '0')

def set_margins(doc):
    if not MARGIN_MM:
        return
    tw = str(int(MARGIN_MM * 56.7))
    for sec in doc.element.body.iter(qn('w:sectPr')):
        m = sec.find(qn('w:pgMar'))
        if m is not None:
            for side in ('w:top', 'w:bottom', 'w:left', 'w:right'):
                m.set(qn(side), tw)

def apply_keep(paras):
    """段落そのものを分割させない＋設問と選択肢の塊を離さない"""
    texts = [base_text(p) for p in paras]
    for i, p in enumerate(paras):
        if not texts[i].strip():
            continue
        if KEEP_LINES:
            set_flag(p, 'w:keepLines')      # 段落を分割しない
        if not KEEP_TOGETHER:
            continue
        nxt = texts[i + 1] if i + 1 < len(texts) else ''
        if CHOICE.match(nxt) and (CHOICE.match(texts[i]) or QUESTION.match(texts[i])):
            set_flag(p, 'w:keepNext')       # 次の段落と分離しない

# ---------- 本体 ----------
def merge_runs(p):
    """書式が同一の隣接ランを結合（単語がランで分断されるのを防ぐ）"""
    runs = p.findall(qn('w:r'))
    prev = None
    for r in runs:
        if r.find(qn('w:t')) is None or len(r.findall(qn('w:t'))) != len(list(r)) - (1 if r.find(qn('w:rPr')) is not None else 0):
            prev = None; continue
        rpr = r.find(qn('w:rPr'))
        key = rpr.xml if rpr is not None else ''
        if prev is not None and prev[1] == key:
            prev[0].find(qn('w:t')).text = (prev[0].find(qn('w:t')).text or '') + (r.find(qn('w:t')).text or '')
            prev[0].find(qn('w:t')).set(qn('xml:space'), 'preserve')
            p.remove(r)
        else:
            prev = (r, key)

def process(inpath, outpath, yomi=None, kobun=None, tk=None):
    """yomi：使う辞書（省略時は読み辞書.json）。kobun：〔古文〕の範囲だけで使う辞書。
    戻り値：(読めなかった語, 要確認の箇所, 振った読みの一覧)"""
    doc = Document(inpath)
    tk = tk or Tokenizer()
    if yomi is None:
        yomi = load_dict()
    kobun_yomi = dict(yomi); kobun_yomi.update(kobun or {})
    matchers = {False: make_matcher(yomi), True: make_matcher(kobun_yomi)}
    in_kobun = [False]
    readings = {}   # (語, 読み) -> {'n': 回数, 'ctx': [前後の文]}
    active = [True]
    unknown = set()
    risky = []   # ルビが答えを漏らす恐れのある設問の検出
    marker_paras = []
    all_paras = []
    RISK = re.compile(r'仮名に|仮名は|読み方|現代仮名遣い|歴史的仮名遣い|送り仮名|音・訓|音訓|読みの組み合わせ|ひらがなで')

    for p in doc.element.body.iter(qn('w:p')):
        all_paras.append(p)
        merge_runs(p)
        for r in list(p.findall(qn('w:r'))):
            tnode = r.find(qn('w:t'))
            if tnode is None or not tnode.text:
                continue
            text = tnode.text
            # 停止／再開・古文の範囲のマーカー
            if any(mk in text for mk in MARKS):
                for mark, state in ((STOP_MARK, False), (RESUME_MARK, True)):
                    if mark in text: active[0] = state
                if KOBUN_END in text: in_kobun[0] = False
                elif KOBUN_MARK in text: in_kobun[0] = True
                for mk in MARKS:
                    text = text.replace(mk, '')
                tnode.text = text
                if not ''.join(t.text or '' for t in p.iter(qn('w:t'))).strip():
                    marker_paras.append(p)   # 空になった目印の行は削除する
                continue
            if RISK.search(text) and active[0]:
                risky.append(text[:40])
            if not active[0] or not HAS_KANJI(text):
                continue
            rPr = r.find(qn('w:rPr'))
            newnodes, buf = [], ''
            cur = kobun_yomi if in_kobun[0] else yomi
            off = 0
            for s, rd, from_dict in segments(text, cur, matchers[in_kobun[0]], tk):
                start = off; off += len(s)
                if not HAS_KANJI(s):
                    buf += s; continue
                if not from_dict and s in cur:
                    rd = cur[s]
                if rd and not YOMI.match(rd):
                    rd = None              # 〇 など、読みが仮名で取れないもの
                    if not REAL_KANJI.search(s):
                        buf += s; continue
                if not rd:
                    unknown.add(s); buf += s; continue
                rec = readings.setdefault((s, rd), {'n': 0, 'ctx': []})
                rec['n'] += 1
                if len(rec['ctx']) < 3:
                    rec['ctx'].append(text[max(0, start - 8):start] + '【' + s + '】' + text[off:off + 8])
                for seg, ruby in split_okurigana(s, rd):
                    if ruby:
                        if buf: newnodes.append(plain_run(buf, rPr)); buf = ''
                        newnodes.append(make_ruby(seg, ruby, rPr))
                    else:
                        buf += seg
            if buf: newnodes.append(plain_run(buf, rPr))
            if newnodes:
                idx = list(p).index(r)
                for n in reversed(newnodes): p.insert(idx, n)
                p.remove(r)
                fix_line_spacing(p)

    for p in marker_paras:
        if p.getparent() is not None:
            p.getparent().remove(p)
        if p in all_paras: all_paras.remove(p)
    if DEL_BLANK:
        for p in list(all_paras):
            if not base_text(p).strip() and p.getparent() is not None:
                p.getparent().remove(p); all_paras.remove(p)
    compact(all_paras)
    set_margins(doc)
    apply_keep(all_paras)

    doc.save(outpath)
    return unknown, risky, readings

def pick_file():
    """引数がないときはファイル選択ダイアログを出す（ダブルクリック起動用）"""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk(); root.withdraw()
        path = filedialog.askopenfilename(
            title='ルビを振るWordファイルを選んでください',
            filetypes=[('Wordファイル', '*.docx')])
        root.destroy()
        return path
    except Exception:
        return input('ファイルのパスを入力してEnter: ').strip().strip('"')

if __name__ == '__main__':
    src = sys.argv[1] if len(sys.argv) > 1 else pick_file()
    if not src:
        sys.exit()
    dst = sys.argv[2] if len(sys.argv) > 2 else re.sub(r'\.docx$', '_ルビ.docx', src)
    if not os.path.exists(DICT_PATH):
        with open(DICT_PATH, 'w', encoding='utf-8') as f:
            f.write('{}')
    unk, risky, _ = process(src, dst)
    print(f'出力: {dst}')
    if unk:
        print('\n【読みが取れなかった語】読み辞書.json に追加してください：')
        print('  ' + '、'.join(sorted(unk)))
    if risky:
        print('\n【要確認】ルビが答えを漏らす可能性のある箇所です。')
        print('　　　　　 マスターに〔ルビ停止〕〔ルビ再開〕を入れて振り直してください：')
        for t in risky:
            print('  ・' + t)
    if len(sys.argv) <= 1:
        input('\nEnterキーで終了します。')
