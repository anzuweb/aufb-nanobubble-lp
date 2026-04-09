"""
AUFB キッチン蛇口用ナノバブル発生装置 LP ワイヤーフレーム生成スクリプト
原稿: AUFB様_キッチン蛇口原稿.txt
"""

import os
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

# ===== 設定 =====
SCOPES = ['https://www.googleapis.com/auth/presentations', 'https://www.googleapis.com/auth/drive']
CREDENTIALS_FILE = '/Users/yamadayuri/あんずウェブ/03_ツール/Google認証/credentials.json'
TOKEN_FILE = '/Users/yamadayuri/あんずウェブ/03_ツール/Google認証/token.json'
PRESENTATION_TITLE = 'AUFB キッチン蛇口用ナノバブル発生装置 LP ワイヤーフレーム'

# ===== スライドサイズ =====
SLIDE_W = 9144000   # 10インチ (EMU)
SLIDE_H = 5143500   # 5.63インチ 16:9 (EMU)

# ===== カラー =====
C_BG_SECTION  = {'red': 0.95, 'green': 0.97, 'blue': 1.0}
C_BG_GRAY     = {'red': 0.88, 'green': 0.88, 'blue': 0.88}
C_BG_DARK     = {'red': 0.1,  'green': 0.15, 'blue': 0.2}
C_BG_WHITE    = {'red': 1.0,  'green': 1.0,  'blue': 1.0}
C_ACCENT      = {'red': 0.05, 'green': 0.45, 'blue': 0.65}   # キッチン = ティール系
C_ORANGE      = {'red': 1.0,  'green': 0.6,  'blue': 0.0}
C_WHITE       = {'red': 1.0,  'green': 1.0,  'blue': 1.0}
C_DARK        = {'red': 0.1,  'green': 0.1,  'blue': 0.1}
C_GRAY        = {'red': 0.5,  'green': 0.5,  'blue': 0.5}
C_BORDER      = {'red': 0.7,  'green': 0.7,  'blue': 0.7}

def rgb(r, g, b): return {'red': r/255, 'green': g/255, 'blue': b/255}
def emu(inches):  return int(inches * 914400)
def pt(px):       return int(px * 9525)

def shape(sid, page_id, shape_type, x, y, w, h):
    return {'createShape': {
        'objectId': sid, 'shapeType': shape_type,
        'elementProperties': {
            'pageObjectId': page_id,
            'size': {'width': {'magnitude': w, 'unit': 'EMU'}, 'height': {'magnitude': h, 'unit': 'EMU'}},
            'transform': {'translateX': x, 'translateY': y, 'scaleX': 1, 'scaleY': 1, 'unit': 'EMU'}
        }
    }}

def text(sid, t):
    return {'insertText': {'objectId': sid, 'text': t, 'insertionIndex': 0}}

def fill_solid(sid, color, opacity=1.0):
    return {'updateShapeProperties': {
        'objectId': sid,
        'shapeProperties': {
            'shapeBackgroundFill': {'solidFill': {'color': {'rgbColor': color}, 'alpha': opacity}},
            'outline': {'propertyState': 'NOT_RENDERED'}
        },
        'fields': 'shapeBackgroundFill,outline'
    }}

def slide_bg(slide_id, color):
    return {'updatePageProperties': {
        'objectId': slide_id,
        'pageProperties': {
            'pageBackgroundFill': {'solidFill': {'color': {'rgbColor': color}}}
        },
        'fields': 'pageBackgroundFill'
    }}

def fill_border(sid, bg, border=None, bw=1.5):
    bc = border or C_BORDER
    return {'updateShapeProperties': {
        'objectId': sid,
        'shapeProperties': {
            'shapeBackgroundFill': {'solidFill': {'color': {'rgbColor': bg}}},
            'outline': {
                'outlineFill': {'solidFill': {'color': {'rgbColor': bc}}},
                'weight': {'magnitude': bw, 'unit': 'PT'}
            }
        },
        'fields': 'shapeBackgroundFill,outline'
    }}

def text_style(sid, s, e, size=None, bold=False, color=None, align=None):
    reqs = []
    style, fields = {}, []
    if size:  style['fontSize'] = {'magnitude': size, 'unit': 'PT'}; fields.append('fontSize')
    if bold:  style['bold'] = True; fields.append('bold')
    if color: style['foregroundColor'] = {'opaqueColor': {'rgbColor': color}}; fields.append('foregroundColor')
    if fields:
        reqs.append({'updateTextStyle': {
            'objectId': sid,
            'textRange': {'type': 'FIXED_RANGE', 'startIndex': s, 'endIndex': e},
            'style': style, 'fields': ','.join(fields)
        }})
    if align:
        reqs.append({'updateParagraphStyle': {
            'objectId': sid,
            'textRange': {'type': 'FIXED_RANGE', 'startIndex': s, 'endIndex': e},
            'style': {'alignment': align}, 'fields': 'alignment'
        }})
    return reqs

def placeholder(page, sid, label, x, y, w, h, bg=None):
    r = []
    r.append(shape(sid, page, 'RECTANGLE', x, y, w, h))
    r.append(text(sid, label))
    r.append(fill_border(sid, bg or C_BG_GRAY, rgb(180, 180, 180)))
    r.extend(text_style(sid, 0, len(label), size=9, color=C_GRAY, align='CENTER'))
    return r

def cta(page, sid, label, x, y, w=emu(2.8), h=pt(40)):
    r = []
    r.append(shape(sid, page, 'ROUND_RECTANGLE', x, y, w, h))
    r.append(text(sid, label))
    r.append(fill_solid(sid, C_ORANGE))
    r.extend(text_style(sid, 0, len(label), size=10, bold=True, color=C_WHITE, align='CENTER'))
    return r

def sec_label(page, sid, label, x=emu(0.1), y=pt(6)):
    r = []
    r.append(shape(sid, page, 'RECTANGLE', x, y, emu(2.2), pt(20)))
    r.append(text(sid, label))
    r.append(fill_solid(sid, rgb(220, 220, 220)))
    r.extend(text_style(sid, 0, len(label), size=7, color=C_GRAY))
    return r

def h2(page, sid, title, bg=None):
    r = []
    r.append(shape(sid, page, 'RECTANGLE', emu(0.5), emu(0.15), emu(9), pt(42)))
    r.append(text(sid, title))
    r.append(fill_solid(sid, bg or C_BG_WHITE))
    r.extend(text_style(sid, 0, len(title), size=18, bold=True, color=C_DARK, align='CENTER'))
    return r


# ===== 認証 =====
def authenticate():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'w') as f:
            f.write(creds.to_json())
    return creds


# ===== スライド構築 =====
def build_slides(svc, pid):
    pres = svc.presentations().get(presentationId=pid).execute()
    first_id = pres['slides'][0]['objectId']

    # セクション順序:
    # S1.ファーストビュー  S2.お悩み  S3.ソリューション
    # S4.効果（4つ）  S5.ナノバブルとは  S6.科学的検証
    # S7.品質アピール  S8.取り付け  S9.FAQ  S10.最終CTA
    slide_ids = [first_id, 'sl002', 'sl003', 'sl004', 'sl005',
                 'sl006', 'sl007', 'sl008', 'sl009', 'sl010']

    creates = []
    for i, sid in enumerate(slide_ids[1:], 1):
        creates.append({'createSlide': {
            'objectId': sid, 'insertionIndex': i,
            'slideLayoutReference': {'predefinedLayout': 'BLANK'}
        }})
    svc.presentations().batchUpdate(presentationId=pid, body={'requests': creates}).execute()

    reqs = []

    # ============================================================
    # S1: ファーストビュー
    # ============================================================
    s = first_id
    reqs.append(slide_bg(s, C_BG_DARK))
    reqs += placeholder(s, f'{s}_img', '📷 キッチン使用シーン写真 / 製品画像',
                        emu(0.2), emu(0.15), emu(6.2), emu(4.5), rgb(40, 55, 65))
    catch = '今あるキッチン蛇口に付けるだけ。\n毎日の料理・洗い物を、ナノバブルの力で変える。'
    reqs.append(shape(f'{s}_catch', s, 'RECTANGLE', emu(6.5), emu(0.4), emu(2.5), pt(70)))
    reqs.append(text(f'{s}_catch', catch))
    reqs.append(fill_solid(f'{s}_catch', C_BG_DARK, 0.0))
    reqs.extend(text_style(f'{s}_catch', 0, len(catch), size=11, bold=True, color=C_WHITE))
    worries = ['排水口のぬめり・臭い', '食材の農薬・汚れ', '手荒れ・洗剤の使いすぎ']
    for i, w in enumerate(worries):
        wy = emu(1.8) + i * pt(38)
        wid = f'{s}_worry{i}'
        reqs.append(shape(wid, s, 'RECTANGLE', emu(6.5), wy, emu(2.5), pt(32)))
        reqs.append(text(wid, f'✗  {w}'))
        reqs.append(fill_border(wid, rgb(255, 240, 240), rgb(200, 100, 100)))
        reqs.extend(text_style(wid, 0, len(w) + 4, size=8, color=rgb(200, 50, 50)))
    reqs += cta(s, f'{s}_cta', 'Amazonで詳細を見る', emu(6.4), emu(4.0))
    reqs += sec_label(s, f'{s}_lbl', 'SECTION 01 | ファーストビュー')

    # ============================================================
    # S2: お悩みセクション
    # ============================================================
    s = 'sl002'
    reqs.append(slide_bg(s, C_BG_WHITE))
    reqs += h2(s, f'{s}_h', 'こんなキッチンのお悩み、ありませんか？')
    problems = [
        'シンクや排水口のぬめり・嫌な臭いが取れない',
        '毎日の食器洗いで手が荒れてしまう',
        '野菜の農薬や細かい汚れが落ちているか不安',
        '洗剤をたくさん使わないと汚れが落ちない気がする',
    ]
    for i, p in enumerate(problems):
        py = emu(0.85) + i * pt(42)
        reqs.append(shape(f'{s}_p{i}', s, 'RECTANGLE', emu(0.5), py, emu(8.5), pt(36)))
        reqs.append(text(f'{s}_p{i}', f'✗  {p}'))
        reqs.append(fill_border(f'{s}_p{i}', rgb(255, 240, 240), rgb(220, 100, 100)))
        reqs.extend(text_style(f'{s}_p{i}', 0, len(p) + 4, size=10, color=rgb(180, 50, 50)))
    reqs += sec_label(s, f'{s}_lbl', 'SECTION 02 | お悩みセクション')

    # ============================================================
    # S3: ソリューション紹介
    # ============================================================
    s = 'sl003'
    reqs.append(slide_bg(s, C_BG_SECTION))
    reqs += h2(s, f'{s}_h', 'ナノバブルが、キッチンのお悩みをまとめて解決します。', C_BG_SECTION)
    reqs += placeholder(s, f'{s}_img', '📷 製品画像（キッチン蛇口取り付けイメージ）',
                        emu(0.3), emu(0.85), emu(3.8), emu(3.8))
    body = ('AUFBナノバブル発生装置をキッチン蛇口に取り付けるだけ。\n'
            '今まで通り水を使うだけで、最大１億３千万個のナノバブルが\n'
            '汚れの奥まで入り込み、洗浄力をぐっと高めます。\n\n'
            '電気不要・洗剤不要。特別なことは何もいりません。')
    reqs.append(shape(f'{s}_body', s, 'RECTANGLE', emu(4.3), emu(1.0), emu(5.0), emu(3.0)))
    reqs.append(text(f'{s}_body', body))
    reqs.append(fill_solid(f'{s}_body', C_BG_SECTION))
    reqs.extend(text_style(f'{s}_body', 0, len(body), size=10, color=C_DARK))
    reqs += sec_label(s, f'{s}_lbl', 'SECTION 03 | ソリューション紹介')

    # ============================================================
    # S4: 効果説明（4つ）
    # ============================================================
    s = 'sl004'
    reqs.append(slide_bg(s, C_BG_WHITE))
    reqs += h2(s, f'{s}_h', 'ナノバブル水で実感できる、キッチンの変化')
    effects = [
        ('🥬', '野菜の農薬・汚れを\nやさしく、しっかり落とす'),
        ('🍽️', '洗剤を減らしても\n食器がスッキリきれいに'),
        ('🚿', 'シンク・排水口の\nぬめり・臭いを抑える'),
        ('↕️', '360°首振り＋\nストレート/シャワー切替'),
    ]
    for i, (icon, label) in enumerate(effects):
        ex = emu(0.3) + i * emu(2.35)
        reqs.append(shape(f'{s}_e{i}', s, 'ROUND_RECTANGLE', ex, emu(0.85), emu(2.1), emu(3.2)))
        reqs.append(text(f'{s}_e{i}', f'{icon}\n\n{label}'))
        reqs.append(fill_border(f'{s}_e{i}', C_BG_WHITE, rgb(100, 180, 200)))
        reqs.extend(text_style(f'{s}_e{i}', 0, len(icon) + 2, size=28, align='CENTER'))
        reqs.extend(text_style(f'{s}_e{i}', len(icon) + 2, len(icon) + 2 + len(label) + 1,
                               size=9.5, color=C_DARK, align='CENTER'))
    reqs += sec_label(s, f'{s}_lbl', 'SECTION 04 | 効果説明')

    # ============================================================
    # S5: ナノバブルとは
    # ============================================================
    s = 'sl005'
    reqs.append(slide_bg(s, C_BG_SECTION))
    reqs += h2(s, f'{s}_h', '毛穴より小さい泡だから、汚れの奥まで届く。', C_BG_SECTION)
    reqs += placeholder(s, f'{s}_img', '📷 ナノバブルサイズ比較図解\n（毛穴・髪の毛・ナノバブルのサイズ比較）',
                        emu(0.3), emu(0.85), emu(4.2), emu(3.8))
    desc = ('ナノバブルの泡の大きさは0.001mm未満。\n'
            '肉眼では見えないほど極小の泡が、汚れの隙間に入り込み、\n'
            '界面活性力で汚れを剥がし取ります。\n\n'
            '【サイズ比較】\n'
            '毛穴：約0.1〜0.3mm\n'
            '髪の毛：約0.08mm\n'
            'ナノバブル：0.001mm未満')
    reqs.append(shape(f'{s}_desc', s, 'RECTANGLE', emu(4.7), emu(1.0), emu(4.6), emu(3.6)))
    reqs.append(text(f'{s}_desc', desc))
    reqs.append(fill_solid(f'{s}_desc', C_BG_SECTION))
    reqs.extend(text_style(f'{s}_desc', 0, len(desc), size=10, color=C_DARK))
    reqs += sec_label(s, f'{s}_lbl', 'SECTION 05 | ナノバブルとは')

    # ============================================================
    # S6: 科学的検証
    # ============================================================
    s = 'sl006'
    reqs.append(slide_bg(s, C_BG_WHITE))
    reqs += h2(s, f'{s}_h', '目で見てわかる、ナノバブルの洗浄力')
    items = [
        ('📷 納豆汚れ比較画像\n（ビフォーアフター）',
         '納豆の汚れで比べてみた\n水道水：除去率 約15%\nナノバブルあり：除去率 約53%\n※当社試験方法による'),
        ('📷 排水口ビフォーアフター\n（設置1日後）',
         '排水口のぬめりに、設置1日で変化が\nこすり洗い等なしで\n目に見えてぬめりが改善'),
    ]
    for i, (img_lbl, desc_txt) in enumerate(items):
        cx = emu(0.4) + i * emu(4.8)
        reqs += placeholder(s, f'{s}_img{i}', img_lbl, cx, emu(0.75), emu(4.2), emu(2.8))
        reqs.append(shape(f'{s}_txt{i}', s, 'RECTANGLE', cx, emu(3.65), emu(4.2), emu(1.2)))
        reqs.append(text(f'{s}_txt{i}', desc_txt))
        reqs.append(fill_border(f'{s}_txt{i}', C_BG_SECTION, rgb(180, 210, 220)))
        reqs.extend(text_style(f'{s}_txt{i}', 0, len(desc_txt), size=9, color=C_DARK, align='CENTER'))
    reqs += sec_label(s, f'{s}_lbl', 'SECTION 06 | 科学的検証')

    # ============================================================
    # S7: 品質アピール
    # ============================================================
    s = 'sl007'
    reqs.append(slide_bg(s, C_BG_SECTION))
    reqs += h2(s, f'{s}_h', '長く使える、高品質設計。', C_BG_SECTION)
    strengths = [
        '🔩  ナノバブル発生部分はオールステンレス製',
        '🏆  特許取得済み（特許番号：7142386）',
        '🏭  自動車部品メーカーの品質規格で製造',
        '🇯🇵  日本製',
    ]
    for i, st_text in enumerate(strengths):
        sy = emu(0.85) + i * pt(50)
        reqs.append(shape(f'{s}_st{i}', s, 'RECTANGLE', emu(0.4), sy, emu(5.8), pt(44)))
        reqs.append(text(f'{s}_st{i}', st_text))
        reqs.append(fill_border(f'{s}_st{i}', rgb(240, 250, 255), C_ACCENT))
        reqs.extend(text_style(f'{s}_st{i}', 0, len(st_text), size=10, color=C_DARK))
    reqs += placeholder(s, f'{s}_img', '📷 製品クローズアップ写真\n（ステンレス素材感がわかるもの）',
                        emu(6.4), emu(0.7), emu(3.1), emu(4.0))
    tagline = '一度手に入れたら、長く使い続けられる設計です。'
    reqs.append(shape(f'{s}_tag', s, 'RECTANGLE', emu(0.4), emu(4.55), emu(5.8), pt(30)))
    reqs.append(text(f'{s}_tag', tagline))
    reqs.append(fill_solid(f'{s}_tag', C_BG_SECTION))
    reqs.extend(text_style(f'{s}_tag', 0, len(tagline), size=9, bold=True, color=C_ACCENT))
    reqs += sec_label(s, f'{s}_lbl', 'SECTION 07 | 品質アピール')

    # ============================================================
    # S8: 簡単取り付け
    # ============================================================
    s = 'sl008'
    reqs.append(slide_bg(s, C_BG_WHITE))
    reqs += h2(s, f'{s}_h', '工具不要！今お使いの蛇口に取り付けるだけ')
    steps = [
        '泡沫水栓（先端のパーツ）を取り外す',
        'AUFBナノバブル発生装置を取り付ける',
        '完成！',
    ]
    for i, step in enumerate(steps):
        sy = emu(0.85) + i * pt(52)
        reqs.append(shape(f'{s}_step{i}', s, 'ROUND_RECTANGLE', emu(0.4), sy, emu(4.5), pt(46)))
        reqs.append(text(f'{s}_step{i}', f'STEP {i + 1}  {step}'))
        reqs.append(fill_border(f'{s}_step{i}', C_BG_SECTION, C_ACCENT))
        reqs.extend(text_style(f'{s}_step{i}', 0, len(step) + 8, size=10, color=C_DARK))
    note = ('付属のアダプターで外ネジ・内ネジどちらにも対応\n'
            '外ネジ M22×1.25 ／ 内ネジ M24×1\n'
            '※機種ごとのネジサイズは各メーカーにご確認ください')
    reqs.append(shape(f'{s}_note', s, 'RECTANGLE', emu(0.4), emu(3.7), emu(4.5), pt(52)))
    reqs.append(text(f'{s}_note', note))
    reqs.append(fill_border(f'{s}_note', rgb(245, 250, 255), rgb(180, 210, 220)))
    reqs.extend(text_style(f'{s}_note', 0, len(note), size=8.5, color=C_GRAY))
    reqs += placeholder(s, f'{s}_img', '📷 取り付け手順の写真\n（STEP1〜3の3点）',
                        emu(5.1), emu(0.7), emu(4.5), emu(4.0))
    reqs += sec_label(s, f'{s}_lbl', 'SECTION 08 | 簡単取り付け')

    # ============================================================
    # S9: FAQ
    # ============================================================
    s = 'sl009'
    reqs.append(slide_bg(s, C_BG_SECTION))
    reqs += h2(s, f'{s}_h', 'よくあるご質問', C_BG_SECTION)
    faqs = [
        ('Q. 本当に汚れが落ちますか？',
         '当社の試験では、水道水のみと比較して納豆汚れの除去率が約53%（水道水は約15%）という結果が出ています。'
         'ナノバブルの微細な泡が汚れの隙間に入り込み、洗浄力を高めます。'),
        ('Q. 詰まりませんか？',
         '一般家庭での通常使用では、ほぼ詰まらない構造です。'),
        ('Q. 電気は必要ですか？',
         '電気や特別な動力は一切不要です。水圧だけでナノバブルを発生させます。'),
        ('Q. どんな蛇口に取り付けられますか？',
         '外ネジタイプ（M22×1.25）に標準対応。付属アダプターで内ネジタイプ（M24×1）にも取り付け可能です。'),
    ]
    for i, (q, a) in enumerate(faqs):
        fy = emu(0.6) + i * emu(1.12)
        reqs.append(shape(f'{s}_q{i}', s, 'RECTANGLE', emu(0.4), fy, emu(9.2), pt(26)))
        reqs.append(text(f'{s}_q{i}', q))
        reqs.append(fill_border(f'{s}_q{i}', C_ACCENT))
        reqs.extend(text_style(f'{s}_q{i}', 0, len(q), size=9, bold=True, color=C_WHITE))
        reqs.append(shape(f'{s}_a{i}', s, 'RECTANGLE', emu(0.4), fy + pt(28), emu(9.2), pt(46)))
        reqs.append(text(f'{s}_a{i}', f'A. {a}'))
        reqs.append(fill_border(f'{s}_a{i}', C_BG_WHITE, rgb(180, 210, 220)))
        reqs.extend(text_style(f'{s}_a{i}', 0, len(a) + 3, size=8.5, color=C_DARK))
    reqs += sec_label(s, f'{s}_lbl', 'SECTION 09 | FAQ')

    # ============================================================
    # S10: 最終CTA
    # ============================================================
    s = 'sl010'
    reqs.append(slide_bg(s, C_BG_DARK))
    heading = 'キッチンのお悩みを、蛇口を替えるだけで解決。'
    reqs.append(shape(f'{s}_h', s, 'RECTANGLE', emu(0.5), emu(1.0), emu(9), pt(55)))
    reqs.append(text(f'{s}_h', heading))
    reqs.append(fill_solid(f'{s}_h', C_BG_DARK, 0.0))
    reqs.extend(text_style(f'{s}_h', 0, len(heading), size=20, bold=True, color=C_WHITE, align='CENTER'))
    sub = '取り付けはたったの3ステップ。\n今日から毎日のキッチン作業が、ナノバブルで変わります。'
    reqs.append(shape(f'{s}_sub', s, 'RECTANGLE', emu(0.5), emu(2.1), emu(9), pt(48)))
    reqs.append(text(f'{s}_sub', sub))
    reqs.append(fill_solid(f'{s}_sub', C_BG_DARK, 0.0))
    reqs.extend(text_style(f'{s}_sub', 0, len(sub), size=11, color=rgb(180, 210, 230), align='CENTER'))
    reqs += placeholder(s, f'{s}_img', '📷 製品画像（メイン）',
                        emu(3.8), emu(0.3), emu(2.5), emu(2.5), rgb(40, 55, 65))
    reqs += cta(s, f'{s}_cta', 'Amazonで詳細を見る', emu(3.2), emu(3.5), emu(3.6), pt(52))
    reqs += sec_label(s, f'{s}_lbl', 'SECTION 10 | 最終CTA')

    # ===== バッチ送信（50件ずつ）=====
    BATCH = 50
    total = len(reqs)
    for i in range(0, total, BATCH):
        batch = reqs[i:i + BATCH]
        svc.presentations().batchUpdate(
            presentationId=pid, body={'requests': batch}
        ).execute()
        print(f'  送信: {i}〜{min(i + BATCH, total)} / {total}')


# ===== メイン =====
def main():
    print('=== AUFB キッチン蛇口用 LP ワイヤーフレーム生成 ===')
    creds = authenticate()
    svc = build('slides', 'v1', credentials=creds)

    print('プレゼンテーション作成中...')
    pres = svc.presentations().create(body={
        'title': PRESENTATION_TITLE,
        'pageSize': {
            'width': {'magnitude': SLIDE_W, 'unit': 'EMU'},
            'height': {'magnitude': SLIDE_H, 'unit': 'EMU'}
        }
    }).execute()
    pid = pres['presentationId']
    print(f'ID: {pid}')

    print('スライド生成中...')
    build_slides(svc, pid)

    url = f'https://docs.google.com/presentation/d/{pid}/edit'
    print(f'\n✅ 完了！')
    print(f'URL: {url}')
    with open('wireframe_id.txt', 'w') as f:
        f.write(url)
    return url

if __name__ == '__main__':
    main()
