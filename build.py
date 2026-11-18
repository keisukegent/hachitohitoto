# -*- coding: utf-8 -*-
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_LINE
from pptx.oxml.ns import qn
from lxml import etree
import json, sys

def C(h): return RGBColor.from_string(h.lstrip('#'))
BG='FFFFFF'; GREEN='0C68B7'; NAVY='17294D'; GOLD='FFD24D'; TXT='222B3D'; LGREEN='BFD9EE'; BEIGE='E8F0F8'; CARD='F5FAFF'; WHITE='FFFFFF'; GREY='5D6B82'
FONT='游ゴシック'
FOOT='はちとひとと AI活用研修  |  第8回'

prs=Presentation(); prs.slide_width=Inches(13.333); prs.slide_height=Inches(7.5)
blank=prs.slide_layouts[6]

def shape(s,kind,x,y,w,h,fill=None,line=None,lw=1.5,dash=False):
    sh=s.shapes.add_shape(kind,Inches(x),Inches(y),Inches(w),Inches(h))
    if fill: sh.fill.solid(); sh.fill.fore_color.rgb=C(fill)
    else: sh.fill.background()
    if line:
        sh.line.color.rgb=C(line); sh.line.width=Pt(lw)
        if dash: sh.line.dash_style=MSO_LINE.DASH
    else: sh.line.fill.background()
    sh.shadow.inherit=False
    return sh

def settext(sh,text,size=22,bold=False,color=TXT,align='c',anchor='m',margin=0.08):
    tf=sh.text_frame; tf.word_wrap=True
    tf.margin_left=tf.margin_right=Inches(margin); tf.margin_top=tf.margin_bottom=Inches(0.04)
    tf.vertical_anchor={'m':MSO_ANCHOR.MIDDLE,'t':MSO_ANCHOR.TOP,'b':MSO_ANCHOR.BOTTOM}[anchor]
    lines=text if isinstance(text,list) else [text]
    for i,l in enumerate(lines):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph()
        if isinstance(l,tuple): t,kw=l
        else: t,kw=l,{}
        p.alignment={'c':PP_ALIGN.CENTER,'l':PP_ALIGN.LEFT,'r':PP_ALIGN.RIGHT}[align]
        r=p.add_run(); r.text=t
        r.font.size=Pt(kw.get('size',size)); r.font.bold=kw.get('bold',bold)
        r.font.color.rgb=C(kw.get('color',color)); r.font.name=FONT
        rPr=r._r.get_or_add_rPr()
        for tag in ('a:ea','a:cs'):
            e=etree.SubElement(rPr,qn(tag)); e.set('typeface',FONT)
        if i>0: p.space_before=Pt(kw.get('gap',6))

def box(s,x,y,w,h,text='',kind=MSO_SHAPE.ROUNDED_RECTANGLE,fill=CARD,line=None,**kw):
    sh=shape(s,kind,x,y,w,h,fill,line,kw.pop('lw',1.5),kw.pop('dash',False))
    if kind==MSO_SHAPE.ROUNDED_RECTANGLE: sh.adjustments[0]=kw.pop('r',0.08)
    if text!='': settext(sh,text,**kw)
    return sh

def tb(s,x,y,w,h,text,**kw):
    t=s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h)); settext(t,text,**kw); return t

def circle(s,x,y,d,text,fill=GREEN,color=WHITE,size=26):
    return box(s,x,y,d,d,text,kind=MSO_SHAPE.OVAL,fill=fill,color=color,size=size,bold=True,margin=0)

def arrow(s,x,y,w,h,fill=GOLD):
    return shape(s,MSO_SHAPE.RIGHT_ARROW,x,y,w,h,fill)

def base(title,msg=None,n=0,dark=False):
    s=prs.slides.add_slide(blank)
    bg=s.background.fill; bg.solid(); bg.fore_color.rgb=C(BEIGE if dark else BG)
    shape(s,MSO_SHAPE.RECTANGLE,0.55,0.05,12.2,0.06,GREEN)
    if not dark:
        shape(s,MSO_SHAPE.OVAL,11.95,0.2,1.15,1.15,CARD,'DCE8F3',1)
        tb(s,0.78,0.4,10.8,0.9,title,size=32,bold=True,color=NAVY,align='l') if title else None
        if title: shape(s,MSO_SHAPE.RECTANGLE,0.88,1.28,1.0,0.06,GREEN)
        if msg: tb(s,0.78,1.4,10.8,0.5,msg,size=22,bold=True,color=GREEN,align='l')
    shape(s,MSO_SHAPE.OVAL,0.15,6.9,0.55,0.55,CARD,'DCE8F3',1)
    badge=box(s,0.0,0.15,0.75,0.4,str(n),kind=MSO_SHAPE.PENTAGON,fill=NAVY,color=WHITE,size=16,bold=True,margin=0.02)
    tb(s,0.78,6.95,7,0.4,FOOT,size=14,color=GREY,align='l')
    t=tb(s,10.6,6.95,2.1,0.4,'',size=14,color=GREY,align='r')
    p=t.text_frame.paragraphs[0]
    fld=etree.SubElement(p._p,qn('a:fld')); fld.set('id','{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}'); fld.set('type','slidenum')
    rPr=etree.SubElement(fld,qn('a:rPr')); rPr.set('lang','ja-JP'); rPr.set('sz','1400')
    sf=etree.SubElement(rPr,qn('a:solidFill')); c=etree.SubElement(sf,qn('a:srgbClr')); c.set('val',GREY)
    lt=etree.SubElement(rPr,qn('a:latin')); lt.set('typeface',FONT)
    tt=etree.SubElement(fld,qn('a:t')); tt.text=str(n)
    r2=p.add_run(); r2.text=' / 12'; r2.font.size=Pt(14); r2.font.color.rgb=C(GREY); r2.font.name=FONT
    return s

NOTES=[]; META=[]
def notes(s,mins,title,key,script,ask,work):
    txt=f"【所要時間】{mins}分\n【このスライドで伝えること】{key}\n\n【講師が話す内容】\n{script}\n\n【参加者への問いかけ】\n{ask}\n\n【ワークの進め方】\n{work}"
    s.notes_slide.notes_text_frame.text=txt
    NOTES.append(dict(n=len(NOTES)+1,mins=mins,title=title,key=key,script=script,ask=ask,work=work))

# ---------- 1
s=base('',n=1)
box(s,0.95,0.75,6.6,0.45,'AIと数字で考える！売上・利益アップ実践ワークショップ',fill=GREEN,color=WHITE,size=16,bold=True,r=0.2)
tb(s,0.9,1.4,11.6,1.5,'今日のゴールは、\n売上と利益を増やすこと',size=44,bold=True,color=NAVY,align='l',anchor='t')
shape(s,MSO_SHAPE.RECTANGLE,0.98,2.95,1.6,0.07,GREEN)
steps=['数字を見る','考える','決める','動く']
for i,t in enumerate(steps):
    k=MSO_SHAPE.PENTAGON if i==0 else MSO_SHAPE.CHEVRON
    sh=box(s,0.9+i*2.8,3.45,3.1,1.3,t,kind=k,fill=GREEN if i%2==0 else GOLD,color=WHITE if i%2==0 else NAVY,size=30,bold=True)
box(s,0.9,5.3,11.4,0.95,'2027年3月までの売上・粗利益改善計画を作る',fill=BEIGE,color=NAVY,size=28,bold=True)
notes(s,4,'今日のゴールは、売上と利益を増やすこと','AIの勉強ではなく、実際の経営改善につなげる。',
 'これまでの研修で、AIの基本、プロンプト、販促文づくり、情報整理、マニュアル作成を学んできました。今日は新しい機能の勉強ではありません。実際の販売データをAIで分析し、皆さんの現場の経験と組み合わせて、2027年3月までの売上・粗利益を良くする計画を作ります。物価上昇で価格改定も考える時期なので、売上だけでなく粗利益も大切にします。流れは「数字を見る→考える→決める→動く」。今日は90分で、最後の「動く」の最初の一歩まで決めます。「AIを使うことが目的ではない。AIと数字を活用して仕事と経営を改善することが目的」です。',
 '「最近の業務で、AIを実際に使えましたか？」（挙手や一言で。使えていなくても責めない）',
 '最終ゴール5点を口頭で確認：①販売状況の把握 ②改善機会の発見 ③商品企画・施策を1〜2件決める ④月別の売上・粗利益目標案 ⑤担当者・期限・最初の行動。')

# ---------- 2
s=base('AIと人間、それぞれの得意な仕事','何でもAIに任せる必要はない',2)
cols=[('表','表計算','正確な集計・計算',['合計・件数・粗利益','間違えてはいけない数字'],GREEN),
      ('AI','AI','分析・仮説・アイデア',['傾向の発見','企画のたたき台'],GOLD),
      ('人','人間','判断・接客・実行',['現場の事情の確認','お客様への価値提供'],GREEN)]
for i,(ic,nm,role,subs,col) in enumerate(cols):
    x=0.7+i*4.1; w=3.75
    box(s,x,1.95,w,4.35,fill=CARD,line=LGREEN,lw=2)
    circle(s,x+w/2-0.6,2.15,1.2,ic,fill=col,color=WHITE if col==GREEN else NAVY,size=34)
    tb(s,x,3.45,w,0.6,nm,size=28,bold=True,color=NAVY)
    tb(s,x+0.1,4.05,w-0.2,0.9,role,size=24,bold=True,color=TXT)
    shape(s,MSO_SHAPE.RECTANGLE,x+0.4,4.95,w-0.8,0.04,GOLD)
    tb(s,x+0.1,5.1,w-0.2,1.1,[(subs[0],{}),(subs[1],{})],size=20,color=GREY)
tb(s,0.7,6.4,11.9,0.5,'AIを無理に使わない。任せる仕事と、人が決める仕事を分ける。',size=20,bold=True,color=NAVY)
notes(s,6,'AIと人間、それぞれの得意な仕事','何でもAIに任せる必要はない。役割分担を理解する。',
 '正確な計算や集計は表計算ソフトの仕事です。AIは大量のデータから傾向を見つけたり、仮説やアイデアを出したりするのが得意です。ただしAIは現場の事情や、お客様の顔を知りません。最終的に判断し、接客し、実行するのは人間です。また、AIを無理に使わないという考え方も大切です。簡単な足し算や、AIに任せると確認の手間が増える作業は、表計算で済ませてよいのです。',
 '「AIの分析結果を、そのまま信じてもよいと思いますか？」（数名に聞く。答えは次のスライド以降につなげる）',
 '3つの役割を順に説明し、最後に問いかけて参加者の考えを聞く。正解を急がず、「確認が必要」という感覚を共有する。')

# ---------- 3
s=base('4つの販売チャネルを見てみよう','販売チャネルによって、売れ方も利益の構造も違う',3)
ch=[('店舗','Airレジ','店舗での販売実績',GOLD),('通販','BASE','通販の販売実績',GREEN),
    ('通販','ユーワード','通販の販売実績',GREEN),('通販','リベシティ','通販の販売実績',GREEN)]
for i,(tag,nm,sub,col) in enumerate(ch):
    x=0.7+(i%2)*6.1; y=1.95+(i//2)*2.2
    box(s,x,y,5.85,2.0,fill=CARD,line=col,lw=2.5)
    box(s,x+0.3,y+0.3,1.1,0.5,tag,fill=col,color=NAVY if col==GOLD else WHITE,size=20,bold=True,r=0.5)
    tb(s,x+0.3,y+0.8,5.3,0.75,nm,size=36,bold=True,color=NAVY,align='l')
    tb(s,x+0.3,y+1.45,5.3,0.45,sub,size=20,color=GREY,align='l')
box(s,0.7,6.4,11.95,0.5,'見る数字：商品別の売上・数量・粗利益',fill=BEIGE,color=NAVY,size=20,bold=True,r=0.3)
notes(s,5,'4つの販売チャネルを見てみよう','販売チャネルによって、売れ方や利益の構造は異なる。',
 '今日見るのは4つの販売チャネルです。店舗のAirレジ、通販のBASE、ユーワード、リベシティ。店舗と通販では、お客様も売れる商品も、かかる費用（販売手数料・送料など）も違います。だから合計だけでなく、チャネル別に見る必要があります。今日は商品別の売上、数量、粗利益を調べ、どの商品をどのチャネルで伸ばせるかを考えます。',
 '「店舗と通販で、売れ方が違う商品は何だと思いますか？」',
 '説明中心。後続の予想クイズにつなげる。※参考デザイン「スライド03_4つの販売チャネル.pptx」は未確認のため、本スライドは共通ルールで作成。見本入手後に差し替え可。')

# ---------- 4
s=base('AIを使う前に、数字を予想しよう','AIに聞く前に、まず自分の考えを持つ',4)
qs=['一番売上が大きい商品は？','最も粗利益を生んでいる商品は？','これから伸ばせそうな販売チャネルは？']
for i,q in enumerate(qs):
    y=1.95+i*1.45
    box(s,0.7,y,11.95,1.25,fill=CARD,line=LGREEN,lw=2)
    circle(s,0.95,y+0.15,0.95,f'Q{i+1}',fill=GOLD,color=NAVY,size=28)
    tb(s,2.2,y+0.1,10.3,1.05,q,size=32,bold=True,color=NAVY,align='l')
tb(s,0.7,6.35,11.9,0.5,'各自で予想 → 隣の人と話す（答え合わせはまだ）',size=20,bold=True,color=GREY)
notes(s,5,'AIを使う前に数字を予想しよう','AIの答えを見る前に、自分の考えを持つ。',
 '分析の前に、皆さん自身の予想を持ちましょう。予想があると、AIの結果を見たときに「合っていた」「違った」と判断できます。違ったときこそ発見です。ここでは答えは発表しません。',
 '「一番売上が大きい商品は？」「最も粗利益を生んでいる商品は？」「これから伸ばせそうな販売チャネルは？」',
 '①各自1分で3つの予想をメモ ②隣の人と2分で共有 ③講師は答えを発表せず、次へ進む。予想メモは後のスライド6で見直す。')

# ---------- 5
s=base('AIに販売データを分析してもらおう','AIは、傾向や改善のチャンスを見つける補助役',5)
fl=[('Excel','販売データ'),('AI','ChatGPTで分析'),('結果','傾向・仮説')]
for i,(a,b) in enumerate(fl):
    x=0.7+i*4.3
    box(s,x,1.9,3.5,1.15,[(a,{'size':26,'bold':True,'color':WHITE}),(b,{'size':18,'color':WHITE,'gap':0})],fill=GREEN if i!=1 else GOLD)
    if i==1: s.shapes[-1].text_frame.paragraphs[0].runs[0].font.color.rgb=C(NAVY); s.shapes[-1].text_frame.paragraphs[1].runs[0].font.color.rgb=C(NAVY)
    if i<2: arrow(s,x+3.6,2.2,0.6,0.55)
st=['売上・粗利益が大きい商品を探す','販路ごとの違いを見つける','伸びている・落ちている商品を探す','改善のチャンスを提案してもらう']
for i,t in enumerate(st):
    x=0.7+i*3.05
    box(s,x,3.4,2.8,2.7,fill=CARD,line=LGREEN,lw=2)
    circle(s,x+0.95,3.55,0.9,str(i+1),size=28)
    tb(s,x+0.1,4.55,2.6,1.5,t,size=20,bold=True,color=TXT)
tb(s,0.7,6.3,11.9,0.5,'まずデータの欠損・重複・集計条件を確認する',size=20,bold=True,color=NAVY)
notes(s,8,'AIに販売データを分析してもらおう','AIはデータから傾向や改善機会を見つける補助役になる。',
 'ここは講師が実演します。実際の販売データをChatGPTに渡し、4つのステップで分析します。①売上・粗利益が大きい商品、②販路ごとの違い、③伸びている商品・落ちている商品、④改善のチャンス。AIに渡す前に、データに欠損や重複、集計条件の違い（税込・税抜、返品・キャンセル、送料）がないかを確認します。顧客の氏名・住所などの個人情報は、AIに渡しません。',
 '「このデータを渡すとき、先に確認しておくことは何でしょう？」',
 '講師が画面を共有して実演。参加者は見ながら、気づいたことをメモする。計算が必要な合計は表計算で検算する。事実と推測を分けて話す。')

# ---------- 6
s=base('AIはどんな発見をした？','グラフで事実を確認し、予想との違いを探す',6)
gr=[('チャネル別 売上・粗利益','チャネルの違い'),('重点商品ランキング','上位商品'),('月別売上推移','伸び・落ち')]
for i,(t,c) in enumerate(gr):
    x=0.7+i*4.1; w=3.75
    tb(s,x,1.95,w,0.5,t,size=22,bold=True,color=NAVY)
    box(s,x,2.45,w,3.0,[('実データ差し替え予定',{'size':22,'bold':True,'color':GREY}),('グラフ配置枠',{'size':18,'color':GREY,'gap':2})],fill='FFFFFF',line=GOLD,lw=2.5,dash=True)
    box(s,x,5.55,w,0.6,'見る点：'+c,fill=BEIGE,color=TXT,size=20,r=0.3)
box(s,0.7,6.3,11.95,0.55,'予想と違った結果は、ありましたか？',fill=GREEN,color=WHITE,size=24,bold=True,r=0.3)
notes(s,7,'AIはどんな発見をした？','AIの分析結果から、経営に役立つ気づきを見つける。',
 '実演の結果をグラフで見ます。①チャネル別の売上・粗利益、②重点商品のランキング、③月別の売上推移。1枚で細かい数字は追いません。大きな特徴だけを見て、先ほどの予想と比べます。※現在は実データ未反映のため、各枠は「実データ差し替え予定」です。実際の販売データで集計（元データの合計と照合済みのもの）をグラフに差し替えてください。原価が不明な商品は粗利益を算出しません。',
 '「予想と違った結果はありましたか？」「意外だった商品・チャネルは？」',
 '各自、スライド4の予想メモと見比べる→隣の人と共有→数名に発表。差し替え手順：枠を削除し、挿入＞グラフでExcelの集計表から作成。')

# ---------- 7
s=base('AIの分析は、本当に正しい？','人間が検証して、はじめて使える',7)
box(s,1.2,1.95,10.9,1.5,'AIの答えを、そのまま信じない。',fill=GREEN,color=WHITE,size=40,bold=True,r=0.12)
pil=[('数字の確認','合計は元データと合っている？'),('現場の経験','在庫切れや販促の影響は？'),('お客様の声','実際の反応は？')]
for i,(t,d) in enumerate(pil):
    x=0.7+i*4.1; w=3.75
    shape(s,MSO_SHAPE.RECTANGLE,x+w/2-0.04,3.45,0.08,0.5,GOLD)
    box(s,x,3.95,w,2.35,fill=CARD,line=GOLD,lw=2.5)
    tb(s,x,4.15,w,0.7,t,size=28,bold=True,color=NAVY)
    shape(s,MSO_SHAPE.RECTANGLE,x+w/2-0.5,4.9,1.0,0.05,GOLD)
    tb(s,x+0.15,5.05,w-0.3,1.1,d,size=22,color=TXT)
notes(s,5,'AIの分析は本当に正しい？','AIの答えを鵜呑みにせず、人間が検証する。',
 'AIは便利ですが、間違えることがあります。そしてAIは現場の事情を知りません。例えば、売上が少ない商品は、人気がないのではなく在庫切れだったのかもしれません。売上が増えた月は、一時的な販促やイベントの効果かもしれません。だから「数字の確認」「現場の経験」「お客様の声」の3つで支えて、はじめて使える情報になります。',
 '「AIの結果と、現場の感覚が違ったとき、どちらを信じますか？なぜ？」',
 '講師が例を紹介→参加者から自分の現場の例を1〜2件聞く。時間があれば隣同士で1分共有。')

# ---------- 8
s=base('データから売上・利益のチャンスを探そう','事実と仮説を分けて、改善の種を見つける',8)
vw=[('売上','もっと売れる\n可能性は？'),('粗利益','もっと利益を残せる\n商品は？'),('販路','伸ばせそうな\n販売チャネルは？')]
for i,(t,q) in enumerate(vw):
    x=0.7+i*4.1; w=3.75
    box(s,x,1.95,w,3.0,fill=CARD,line=LGREEN,lw=2)
    box(s,x,1.95,w,0.85,t,fill=GREEN if i!=1 else GOLD,color=WHITE if i!=1 else NAVY,size=28,bold=True,r=0.15)
    tb(s,x+0.1,2.95,w-0.2,1.9,q,size=26,bold=True,color=TXT)
wf=[('個人で考える','3分'),('意見交換','4分'),('3つに絞る','3分')]
for i,(t,m) in enumerate(wf):
    k=MSO_SHAPE.PENTAGON if i==0 else MSO_SHAPE.CHEVRON
    box(s,0.7+i*3.95,5.35,4.0,1.2,[(t,{'size':26,'bold':True,'color':WHITE if i!=1 else NAVY}),(m,{'size':20,'color':WHITE if i!=1 else NAVY,'gap':0})],kind=k,fill=GREEN if i!=1 else GOLD)
notes(s,10,'データから売上・利益のチャンスを探そう','データ分析を、具体的な改善機会につなげる。',
 '分析で終わらず、チャンスを探します。視点は3つ。売上：もっと売れる可能性は？粗利益：もっと利益を残せる商品は？販路：伸ばせそうな販売チャネルは？データで確認できた「事実」と、そこから考えた「仮説」を区別して書いてください。',
 '「それは事実ですか？仮説ですか？」',
 '個人3分（付箋かメモに3視点で書く）→意見交換4分（グループ）→共有・絞り込み3分（チームで3つに）。AIに「他の視点は？」と聞いてもよいが、採用は人が決める。')

# ---------- 9
s=base('3月までの商品・販売企画を考えよう','忙しい現場でも、実行できる企画を選ぶ',9)
ps=['セット・ギフト企画','リピート促進','重点商品の強化',['通販の商品紹介','販売方法の改善']]
for i,t in enumerate(ps):
    x=0.7+i*3.05
    box(s,x,1.95,2.8,2.9,fill=CARD,line=LGREEN,lw=2)
    circle(s,x+0.95,2.15,0.9,str(i+1),fill=GOLD,color=NAVY,size=28)
    tb(s,x+0.1,3.2,2.6,1.55,t if isinstance(t,list) else [t],size=22,bold=True,color=NAVY)
tb(s,0.7,5.0,3,0.5,'選ぶ3つの基準',size=22,bold=True,color=GREY,align='l')
cr=['粗利益への貢献','実行しやすさ','お客様への価値']
for i,t in enumerate(cr):
    box(s,0.7+i*4.1,5.55,3.75,0.95,t,fill=GREEN,color=WHITE,size=26,bold=True,r=0.4)
notes(s,15,'3月までの商品・販売企画を考えよう','忙しい現場でも実行できる商品企画を選ぶ。',
 '4つの方向性から考えます。①セット・ギフト企画 ②リピート促進 ③重点商品の強化 ④通販の商品紹介・販売方法の改善。選ぶ基準は「粗利益への貢献」「実行しやすさ」「お客様への価値」の3つ。新しい業務の負担が大きすぎる企画は、続きません。1〜2件に絞ります。',
 '「これは、今の人数と忙しさでも続けられますか？」',
 '企画案を考える5分→3つの基準で比較する5分→1〜2件に絞る5分。絞る際、新しい業務負担が大きすぎないか必ず確認する。')

# ---------- 10
s=base('売上と粗利益の目標を決めよう','目標は、根拠のある数字から作る',10)
box(s,0.7,1.95,5.85,1.15,[('売上目標 ＝ 販売単価 × 販売数量',{'size':22,'bold':True,'color':NAVY}),('（商品別に合計）',{'size':18,'color':GREY,'gap':0})],fill=CARD,line=GOLD,lw=2.5)
box(s,6.8,1.95,5.85,1.15,[('粗利益目標 ＝ 売上目標 − 売上原価',{'size':22,'bold':True,'color':NAVY})],fill=CARD,line=GOLD,lw=2.5)
rows=5; cols_w=[2.4,2.7,2.7,4.15]
tbl=s.shapes.add_table(rows,4,Inches(0.7),Inches(3.3),Inches(sum(cols_w)),Inches(3.0)).table
hdr=['月','売上目標','粗利益目標','重点施策']; mon=['2026年12月','2027年1月','2027年2月','2027年3月']
tblPr=tbl._tbl.tblPr
for a in ('bandRow','firstRow'): tblPr.set(a,'0')
for j,w in enumerate(cols_w): tbl.columns[j].width=Inches(w)
for i in range(rows):
    tbl.rows[i].height=Inches(0.6)
    for j in range(4):
        c=tbl.cell(i,j); c.vertical_anchor=MSO_ANCHOR.MIDDLE
        c.fill.solid(); c.fill.fore_color.rgb=C(GREEN if i==0 else (CARD if i%2 else BEIGE))
        txt=hdr[j] if i==0 else (mon[i-1] if j==0 else '')
        tf=c.text_frame; p=tf.paragraphs[0]; p.alignment=PP_ALIGN.CENTER
        r=p.add_run(); r.text=txt; r.font.size=Pt(22 if i==0 else 20); r.font.bold=(i==0 or j==0)
        r.font.color.rgb=C(WHITE if i==0 else NAVY); r.font.name=FONT
tb(s,0.7,6.5,11.9,0.4,'数字は空欄（過去実績・重点商品の販売数量から記入）。手数料・送料は別に確認。',size=18,color=GREY,align='l')
notes(s,15,'売上と粗利益の目標を決めよう','目標は根拠のある数字から作る。',
 '基本の式は2つ。売上目標＝販売単価×販売数量（商品別に合計）。粗利益目標＝売上目標−売上原価。まず過去の実績と、重点商品の販売数量を見ます。「だいたいこのくらい」ではなく、数量×単価で説明できる数字にします。2026年12月〜2027年3月の月別に、売上目標・粗利益目標・重点施策を記入します。販売手数料・送料などは粗利益とは分けて、採算を別に確認します。',
 '「この目標は、何個売れば達成できますか？」',
 '表の数字は空欄（実績データ未反映のため、数字を捏造しない）。個人またはチームで記入→最終的な数値目標は経営者・責任者が承認する。')

# ---------- 11
s=base('明日からの行動を決めよう','担当者と期限を決めて、実行につなげる',11)
it=['何をする？','誰がする？','いつまでに？','最初の一歩は？','どう成果を確認する？']
for i,t in enumerate(it):
    y=2.1+i*0.95
    box(s,0.7,y,0.6,0.6,'',kind=MSO_SHAPE.RECTANGLE,fill=CARD,line=GREEN,lw=3)
    box(s,1.5,y-0.08,4.5,0.76,t,fill=GREEN,color=WHITE,size=26,bold=True,align='l',margin=0.25,r=0.2)
    box(s,6.2,y-0.08,6.45,0.76,'',fill=CARD,line=LGREEN,lw=2,r=0.2)
tb(s,6.3,1.25,6.3,0.4,'',size=14)
notes(s,7,'明日からの行動を決めよう','担当者と期限を決めて、実行につなげる。',
 '企画は、実行して初めて意味があります。1〜2件の施策について、5つを決めてください。何をする？誰がする？いつまでに？最初の一歩は？どう成果を確認する？成果は「売上」「粗利益」「販売数量」など、確認する数字まで決めておきます。',
 '「明日、最初に何をしますか？」',
 '改善施策は1〜2件に絞り、担当者・期限・確認する数字を右側の枠に記入（紙または画面上）。チームで1人ずつ発表し、ほかの人は「できそうか」を一言確認する。')

# ---------- 12
s=base('',n=12,dark=True)
box(s,5.65,1.45,2.0,0.55,'まとめ',fill=GREEN,color=WHITE,size=22,bold=True,r=0.5)
tb(s,0.9,2.0,11.5,2.6,['AIは考えるパートナー。','決めて、動くのは私たち。'],size=44,bold=True,color=NAVY)
tb(s,0.9,5.0,11.5,0.6,'数字を見る → 考える → 実行する → 結果を確かめる',size=22,bold=True,color=GREEN)
notes(s,3,'今日から実行して、結果を確かめよう','AIも分析も、行動につながって初めて意味がある。',
 '今日はAIと数字を使って、売上・粗利益を良くする計画を作りました。AIは考えるパートナー。決めて、動くのは私たちです。今日決めたことを、まず1つ実行してください。次回、結果を一緒に振り返ります。',
 '「今日決めたことのうち、明日やる1つは何ですか？」',
 '全員で今日の決定事項を確認して終了。次回の振り返りの日程を共有する。')

out=''+__import__('os').path.dirname(__import__('os').path.abspath(__file__))+'/はちとひとと_第8回_AI販売データ分析研修'
prs.save(out+'.pptx')
tot=sum(n['mins'] for n in NOTES); print('slides',len(prs.slides),'minutes',tot)
md=['# はちとひとと 第8回 講師用シナリオ','','AIと数字で考える！売上・利益アップ実践ワークショップ（90分）','',
    '> 注意：見本PPTX・販売データは制作環境で確認できていません。スライド6・10の数字枠は「実データ差し替え予定」です。','','| # | スライド | 時間 |','|---|---|---|']
for n in NOTES: md.append(f"| {n['n']} | {n['title']} | {n['mins']}分 |")
md.append(f'| | 合計 | {tot}分 |'); md.append('')
for n in NOTES:
    md+= [f"## スライド{n['n']}｜{n['title']}（{n['mins']}分）",'',f"**伝えること**：{n['key']}",'','**講師が話す内容**',n['script'],'',f"**問いかけ**：{n['ask']}",'',f"**ワークの進め方**：{n['work']}",'']
open(''+__import__('os').path.dirname(__import__('os').path.abspath(__file__))+'/はちとひとと_第8回_講師用シナリオ.md','w').write('\n'.join(md))
