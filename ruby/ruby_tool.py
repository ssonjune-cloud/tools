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

KANJI = re.compile(r'[一-鿿々〇〻]')
YOMI_OK = re.compile(r'^[ぁ-ゟー]+$')   # ひらがな（ゐ・ゑ・ゝ・ゞを含む）と長音
SEP = re.compile(r'\s*(?:＝|=|：|:|→|⇒|->|\t|，|,|、)\s*')
PAREN = re.compile(r'^(.+?)\s*[（(]\s*([ぁ-ゟー]+)\s*[）)]$')


def _get(name):
    """辞書ファイルを読む。手元にあればそれを、なければ GitHub から取る"""
    local = HERE / name
    if local.exists():
        return local.read_text(encoding="utf-8")
    url = BASE_URL + urllib.request.quote(name)
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


def run(paths, kyoka="共通のみ", seikei=False, konkai="", out_dir=None):
    """paths の docx それぞれにルビを振り、出力ファイルのパスの一覧を返す"""
    import furigana
    from janome.tokenizer import Tokenizer

    yomi = load_json_dict("辞書/共通.json")
    if kyoka and kyoka != "共通のみ":
        yomi.update(load_json_dict(f"辞書/{kyoka}.json"))
    kobun = load_json_dict("辞書/古文.json")
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
        if unk:
            print("\n【読みがわからずルビを振らなかった語】今回の辞書に「語＝よみ」で足してください：")
            print("  " + "、".join(sorted(unk)))
        print("\n【振った読みの一覧】↓ここから下を全部コピーして、AI に点検してもらえます")
        print("-" * 40)
        print(reading_list(readings))
        print("-" * 40)
    return outs
