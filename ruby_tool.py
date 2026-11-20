# -*- coding: utf-8 -*-
"""
Colab ノート（ルビ振り.ipynb）から呼ばれる本体。
共通の辞書・教科の辞書・今回の辞書を重ねて furigana.py を動かし、
必要なら mondai_shoshiki.py で体裁を整える。
"""
import json
import os
import re
import urllib.request
from pathlib import Path

BASE_URL = "https://raw.githubusercontent.com/ssonjune-cloud/tools/main/ruby/"
HERE = Path(__file__).resolve().parent
# 辞書のファイル名（文字化けを避けるため英字にしている）
DICT_FILE = {"共通": "common", "古文": "kobun",
             "国語": "kokugo", "社会": "shakai", "数学": "sugaku", "理科": "rika", "音楽": "ongaku",
             "美術": "bijutsu", "保健体育": "hotai", "技術・家庭": "gika", "英語": "eigo"}

KANJI = re.compile(r'[一-鿿々〇〻]')
YOMI_OK = re.compile(r'^[ぁ-ゟー]+$')   # ひらがな（ゐ・ゑ・ゝ・ゞを含む）と長音
SEP = re.compile(r'\s*(?:＝|=|：|:|→|⇒|->|\t|，|,|、)\s*')
PAREN = re.compile(r'^(.+?)\s*[（(]\s*([ぁ-ゟー]+)\s*[）)]$')


def _get(name):
    """辞書ファイルを読む。手元にあればそれを、なければ GitHub から取る"""
    local = HERE / name
    if local.exists():
        return local.read_text(encoding="utf-8")
    url = BASE_URL + name
    with urllib.request.urlopen(url, timeout=20) as r:
        return r.read().decode("utf-8")


def load_json_dict(name):
    try:
        return json.loads(_get(name) or "{}")
    except Exception as e:
        print(f"（{name} を読み込めませんでした：{e}）")
        return {}


def parse_dict_text(text):
    """AI が出した辞書を読み取る。JSON でも「語＝よみ」の一行ずつでもよい。
    戻り値：(辞書, 読み取れなかった行)"""
    text = (text or "").strip()
    if not text:
        return {}, []
    body = text.strip("`").strip()
    if body.startswith("json"):
        body = body[4:].strip()
    for cand in (body, "{" + body.strip().rstrip(",") + "}"):
        try:
            d = json.loads(cand)
            if isinstance(d, dict):
                out, bad = {}, []
                for k, v in d.items():
                    k, v = str(k).strip(), str(v).strip()
                    if KANJI.search(k) and YOMI_OK.match(v):
                        out[k] = v
                    else:
                        bad.append(f"{k}：{v}")
                return out, bad
        except Exception:
            pass
    out, bad = {}, []
    for line in text.splitlines():
        s = line.strip()
        if not s or s.startswith(("#", "```", "{", "}")) or s in ("json",):
            continue
        s = re.sub(r'^(?:[-・*•●○◆▶]|\d+[.)．）])\s*', '', s)
        s = s.rstrip(",，、 ").replace('"', '').replace("'", '').replace("「", "").replace("」", "")
        m = PAREN.match(s)
        if m:
            k, v = m.group(1), m.group(2)
        else:
            parts = SEP.split(s, maxsplit=1)
            if len(parts) != 2:
                parts = s.split(None, 1)
            if len(parts) != 2:
                bad.append(line.strip()); continue
            k, v = parts
        k, v = k.strip(), v.strip()
        v = v.split("（")[0].split("(")[0].strip()   # 「よみ（説明）」の説明は捨てる
        if KANJI.search(k) and YOMI_OK.match(v):
            out[k] = v
        else:
            bad.append(line.strip())
    return out, bad


def reading_list(readings, limit_ctx=1):
    """振った読みの一覧（AI に点検してもらう用）"""
    lines = []
    for (s, rd), rec in sorted(readings.items(), key=lambda x: (x[0][0], x[0][1])):
        n = 3 if len(s) == 1 else limit_ctx
        ctx = " ／ ".join(rec["ctx"][:n])
        lines.append(f"{s}＝{rd}　（{rec['n']}回）　{ctx}")
    return "\n".join(lines)


AI_PROMPT = """中学校の定期試験の問題に、ツールで自動的にふりがな（ルビ）を振りました。
下の一覧は「語＝ツールが振った読み （回数） 前後の文」で、【】の中がその語です。
中学生が読む試験問題として、読みが誤っているものだけを選び、正しい読みを次の形で出力してください。

・一行に一つ、「語＝よみ」の形で書く（例：一日中＝いちにちじゅう）
・よみはひらがなだけで書く
・前後の文によって読みが変わる語は、前後の文字を含めて書く。その文字は一覧の前後の文に実際にあるとおりに書く（例：何が＝なにが、長いこと家に＝ながいこといえに）
・古文の文章の中の語は歴史的仮名遣いで書く。現代文にも同じ語があるときは、前後の文字を含めて区別できるようにする
・人名・地名・教科の用語は、教科書での読みに合わせる
・誤りのない語は書かない。説明や見出しも書かない。誤りが一つもなければ「なし」とだけ書く
"""


def doc_text(path):
    """docx の本文（ルビの読みを除く）"""
    import zipfile
    d = zipfile.ZipFile(path).read("word/document.xml").decode("utf-8")
    d = re.sub(r"<w:rt>.*?</w:rt>", "", d, flags=re.S)
    return "".join(re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", d))


def ai_corrections(readings, text, ai_fn, kyoka):
    """AI に一覧を点検させ、直す読みの辞書を返す。使えなければ None"""
    lines = reading_list(readings).splitlines()
    head = AI_PROMPT + (f"\nこの問題の教科は{kyoka}です。\n" if kyoka in DICT_FILE else "")
    out = {}
    for i in range(0, len(lines), 250):          # 長い一覧は分けて頼む
        reply = ai_fn(head + "\n" + "\n".join(lines[i:i + 250]))
        d, _ = parse_dict_text(str(reply))
        out.update(d)
    have = {(s, r) for (s, r) in readings}
    # 本文にない語・いまと同じ読みは捨てる
    return {k: v for k, v in out.items() if k in text and (k, v) not in have}


def run(paths, kyoka="共通のみ", seikei=False, konkai="", out_dir=None, ai_fn=None):
    """paths の docx それぞれにルビを振り、出力ファイルのパスの一覧を返す"""
    import furigana
    from janome.tokenizer import Tokenizer

    yomi = load_json_dict("dict/common.json")
    if kyoka in DICT_FILE and kyoka != "共通":
        yomi.update(load_json_dict(f"dict/{DICT_FILE[kyoka]}.json"))
    kobun = load_json_dict("dict/kobun.json")
    extra, bad = parse_dict_text(konkai)
    yomi.update(extra)

    print(f"辞書：共通" + (f"＋{kyoka}" if kyoka != "共通のみ" else "")
          + (f"＋今回の辞書 {len(extra)}語" if extra else "") + f"（合わせて {len(yomi)}語）")
    if bad:
        print("\n【今回の辞書で読み取れなかった行】（読みはひらがなで書いてください）")
        for b in bad:
            print("  ・" + b)

    tk = Tokenizer()
    outs = []
    for src in paths:
        src = Path(src)
        dst = Path(out_dir or src.parent) / (src.stem + "_ルビ.docx")
        unk, risky, readings = furigana.process(str(src), str(dst), yomi=yomi, kobun=kobun, tk=tk)
        fixed = {}
        if ai_fn:
            print(f"\n{src.name}：AIが読みを点検しています（1〜2分かかることがあります）…")
            try:
                fixed = ai_corrections(readings, doc_text(str(src)), ai_fn, kyoka)
            except Exception as e:
                print(f"（AIの点検は使えませんでした：{e}）")
                print("  下の【振った読みの一覧】を、使い方にある指示文と一緒にAIに貼って点検できます。")
            fixed = {k: v for k, v in (fixed or {}).items() if k not in yomi}   # 辞書・今回の辞書が優先
            if fixed:
                y2 = dict(yomi); y2.update(fixed)
                unk, risky, readings = furigana.process(str(src), str(dst), yomi=y2, kobun=kobun, tk=tk)
        if seikei:
            import mondai_shoshiki
            mondai_shoshiki.process(dst, dst)
        outs.append(dst)
        print("\n" + "=" * 40)
        print(f"できたファイル：{dst.name}" + ("（書式も整えました）" if seikei else ""))
        if risky:
            print("\n【要確認】ルビで答えがわかってしまうかもしれない設問です。")
            print("  元のファイルで、その範囲を〔ルビ停止〕と〔ルビ再開〕で挟んで、もう一度実行してください。")
            for t in risky:
                print("  ・" + t)
        if ai_fn and not fixed:
            print("\n【AIの点検】直す読みはありませんでした。")
        if fixed:
            print("\n【AIが直した読み】念のため目で確かめてください。違っていたら、正しい読みを")
            print("  「今回の辞書」に「語＝よみ」で書いて、もう一度実行すると、そちらが優先されます。")
            for k, v in fixed.items():
                print(f"  ・{k} → {v}")
            print("  ※次からも同じ直しを使いたい語は、作成者に伝えて共通・教科の辞書に入れてもらいましょう。")
        if unk:
            print("\n【読みがわからずルビを振らなかった語】今回の辞書に「語＝よみ」で足してください：")
            print("  " + "、".join(sorted(unk)))
        print("\n【振った読みの一覧】↓ここから下を全部コピーして、AI に点検してもらえます")
        print("-" * 40)
        print(reading_list(readings))
        print("-" * 40)
    return outs
