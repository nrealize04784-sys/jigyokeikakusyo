# -*- coding: utf-8 -*-
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, DataBarRule
from openpyxl.comments import Comment

F="Meiryo"
INK="1B2330"; ACC="1F4E6B"; RULE="DFD8CA"; BAND="FAF8F3"
INPUT="FFF3C4"   # 記入欄（黄）
EDIT="EAF1F6"    # 選んで変えるセル（水色）
GREEN="1C6647"; AMBER="8A5A12"; RED="993124"; GREY="6E747D"
A="https://claude.ai/code/artifact/"
thin=Side(style="thin",color=RULE); box=Border(left=thin,right=thin,top=thin,bottom=thin)

wb=Workbook()

# ---------------- 選択肢（非表示） ----------------
opt=wb.create_sheet("選択肢")
opt["A1"]="進捗"; opt["A1"].font=Font(name=F,size=9,bold=True)
for i in range(21):
    c=opt.cell(row=2+i,column=1,value=i*0.05); c.number_format="0%"; c.font=Font(name=F,size=9)

ANKEN=["REALIZE CLUB（法人向け）","LIFE MAKE PARTNERS（加盟店募集）","まちの相談窓口","不動産FC加盟"]
STATES=["未着手","進行中","確認待ち"]
OWNERS=["未定","社長","白石","池之上","大熊","片山","寺井","全員",
        "大熊・片山","池之上・片山","大熊・白石","全員・池之上","社長・白石・片山",
        "白石・池之上・大熊","白石・片山","社長・寺井・池之上・弁護士"]
for col,(title,vals) in enumerate([("案件",ANKEN),("状態",STATES),("担当",OWNERS)], start=2):
    h=opt.cell(row=1,column=col,value=title); h.font=Font(name=F,size=9,bold=True)
    opt.column_dimensions[get_column_letter(col)].width=30
    for i,v in enumerate(vals):
        opt.cell(row=2+i,column=col,value=v).font=Font(name=F,size=9)

def rng(col,n): return "'選択肢'!${0}$2:${0}${1}".format(col,n+1)
PCT_LIST   = "'選択肢'!$A$2:$A$22"
ANKEN_LIST = rng("B",len(ANKEN))
STATE_LIST = rng("C",len(STATES))
OWNER_LIST = rng("D",len(OWNERS))

def sheet_header(ws,title,lead,note):
    ws["A1"]=title; ws["A1"].font=Font(name=F,size=16,bold=True,color=INK)
    ws["A2"]=lead;  ws["A2"].font=Font(name=F,size=9,color=GREY)
    ws["A3"]=note;  ws["A3"].font=Font(name=F,size=9,bold=True,color=AMBER)

def write_head(ws,row,head,widths):
    for i,h in enumerate(head,1):
        c=ws.cell(row=row,column=i,value=h)
        c.font=Font(name=F,size=9,bold=True,color="FFFFFF")
        c.fill=PatternFill("solid",fgColor=ACC)
        c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
        c.border=box
    ws.row_dimensions[row].height=34
    for i,w in enumerate(widths,1):
        ws.column_dimensions[get_column_letter(i)].width=w

def dv(ws,kind,rng_,f1,title,prompt,blank=False):
    d=DataValidation(type=kind,formula1=f1,allow_blank=blank)
    d.promptTitle=title; d.prompt=prompt
    ws.add_data_validation(d); d.add(rng_)
    return d

# ================================================== LP
ws=wb.active; ws.title="LP"
HEAD=["案件","LP名","デザインの方向","原稿・デザイン","問い合わせ導線","本番公開",
      "本番URL（記入欄）","残っていること","更新日"]
W=[24,32,34,14,15,11,36,58,11]
sheet_header(ws,"制作進捗ボード ─ LP",
  "これまでに作った11本のLP。実物を1ページずつ読んで確認したもの。調査日 2026-09-03",
  "水色のセルはクリックすると▼が出て選べます。黄色のG列にURLを入れると、本番公開が自動で○になります。")
HR=5; write_head(ws,HR,HEAD,W)

lp=[
("REALIZE CLUB（法人向け）","会社が、社員の人生を守る時代へ。","深緑 × 和モダン／昭和と令和の対比","○","△",
 "送信ボタンがメーラーを開くだけ。受付フォームの実装が必要／会社概要ページが未作成／本番ドメインへの公開",
 "2026-09-03"),
("REALIZE CLUB（法人向け）","会社が社員の人生を支える時代へ","星空タウン／クリーム × 濃紺のイラスト","○","△",
 "3ページ揃っている唯一のRC案。採用するかの判断／送信はメーラー起動のみ",
 "2026-08-31"),
("REALIZE CLUB（法人向け）","困る前に、灯す。","漆黒 × 炎のグラデーション","○","×",
 "kaisha.html ほかへの相対リンクが残り、単体では飛べない／会社概要ページの作成",
 "2026-08-31"),
("REALIZE CLUB（法人向け）","その安心、今がお得。","スーパーの特売ポップ／白地に赤と黄","○","×",
 "リンクがすべてページ内アンカー。外部への導線が1本もない／会社概要ページ、送信先の設定",
 "2026-08-31"),
("REALIZE CLUB（法人向け）","社員が、溶けていく前に。","アイスのメタファー／手描き風","○","×",
 "会社概要ページが未作成／送信はメーラー起動のみ",
 "2026-08-28"),
("LIFE MAKE PARTNERS（加盟店募集）","家より先に、人生の話を。","韓国料理店の看板／木の壁 × 生成りの横断幕","○","○",
 "本番ドメインへの公開だけが残り。問い合わせはLINE公式（lin.ee/WHYApuM）に集約済み",
 "2026-09-03"),
("LIFE MAKE PARTNERS（加盟店募集）","いつも同じ集客ばっかやってんじゃねぇ！","喝・挑発ポスター／赤と黄の斜めストライプ","○","×",
 "kaisha-katsu.html ほかへの相対リンクが残り飛べない／行き先が「#」のままのボタンが2つ",
 "2026-08-28"),
("LIFE MAKE PARTNERS（加盟店募集）","加盟資格は、人生への愛。ただそれだけ。","星空タウン／RC星空案と同じ世界観","○","△",
 "行き先が「#」のままのボタンが2つ／統廃合の判断",
 "2026-08-24"),
("まちの相談窓口","不動産屋が、まちの相談窓口になる。","銭湯ポスター（リソグラフ2色刷り）／LINE誘導版","○","△",
 "行き先が「#」のままのリンクが2つ残っている／本番ドメインへの公開",
 "2026-09-03"),
("まちの相談窓口","まちの相談窓口（フォーム版）","銭湯ポスター／申込フォーム内蔵","○","△",
 "LINE誘導版に置き換わっている。残すか消すかの判断",
 "2026-08-20"),
("不動産FC加盟","収益が積み上がる不動産FC加盟へ。","ボクシング／黒地 × 赤とゴールド","○","×",
 "ページのタイトルタグが空。共有すると名前なしで出る／資料請求・会社概要ページ、送信先の設定",
 "2026-08-19"),
]

for n,r in enumerate(lp):
    ex=HR+1+n
    anken,name,d,s1,s3,todo,upd=r
    for col,val in ((1,anken),(2,name),(3,d),(4,s1),(5,s3),(8,todo),(9,upd)):
        ws.cell(row=ex,column=col,value=val)
    # URLを入れると ○ がその公開ページへのリンクになる（COUNTIFは表示文字「○」で数える）
    ws.cell(row=ex,column=6,value='=IF(G{0}="","×",HYPERLINK(G{0},"○"))'.format(ex))
    ws.row_dimensions[ex].height=46
    for col in range(1,len(HEAD)+1):
        c=ws.cell(row=ex,column=col); c.border=box
        c.font=Font(name=F,size=9,color=INK)
        c.alignment=Alignment(vertical="center",wrap_text=(col in (1,2,3,8)),
                              horizontal="center" if col in (4,5,6,9) else "left")
        if n%2==1 and col!=7: c.fill=PatternFill("solid",fgColor=BAND)
    for col in (1,4,5):
        ws.cell(row=ex,column=col).fill=PatternFill("solid",fgColor=EDIT)
    g=ws.cell(row=ex,column=7)
    g.fill=PatternFill("solid",fgColor=INPUT); g.font=Font(name=F,size=9,color="0000FF")
    g.alignment=Alignment(vertical="center")

first,last=HR+1,HR+len(lp)
dv(ws,"list","A{0}:A{1}".format(first,last),ANKEN_LIST,"案件を選ぶ","どの案件のLPかを選びます")
dv(ws,"list","D{0}:E{1}".format(first,last),'"○,△,×"',"選んでください","○＝できている／△＝仮づけ／×＝手つかず")
ws.cell(row=first,column=7).comment=Comment(
 "公開した本番LPのURLをここに貼ってください。\n入力すると本番公開が○になります。","制作進捗ボード",height=80,width=250)
for mark,color in (("×",RED),("△",AMBER),("○",GREEN)):
    ws.conditional_formatting.add("D{0}:F{1}".format(first,last),
        CellIsRule(operator="equal",formula=['"%s"'%mark],font=Font(name=F,size=9,bold=True,color=color)))

tot=last+1
c=ws.cell(row=tot,column=3,value="○の数"); c.font=Font(name=F,size=9,bold=True,color=INK)
c.alignment=Alignment(horizontal="right")
for col in (4,5,6):
    L=get_column_letter(col)
    c=ws.cell(row=tot,column=col,value='=COUNTIF({0}{1}:{0}{2},"○")'.format(L,first,last))
    c.font=Font(name=F,size=9,bold=True,color=ACC); c.alignment=Alignment(horizontal="center")
    c.number_format='0"／11"'
c=ws.cell(row=tot,column=7,value='=COUNTA(G{0}:G{1})&" 本 公開ずみ"'.format(first,last))
c.font=Font(name=F,size=9,bold=True,color=ACC)
for col in range(3,8):
    ws.cell(row=tot,column=col).border=Border(top=Side(style="medium",color=ACC))
ws.freeze_panes="C6"; ws.auto_filter.ref="A{0}:I{1}".format(HR,last); ws.sheet_view.showGridLines=False

# ================================================== タスクシート
THEAD=["何を作るか／するか","どこで使う","担当","進捗","状態","URL記入欄","メモ"]
TW=[34,30,24,10,11,34,44]

def task_sheet(name,title,lead,note,items):
    s=wb.create_sheet(name)
    sheet_header(s,title,lead,note); write_head(s,5,THEAD,TW)
    r=6
    for nm,where,own,pct,state,memo in items:
        for col,val in ((1,nm),(2,where),(3,own),(4,pct/100.0),(5,state),(7,memo)):
            s.cell(row=r,column=col,value=val)
        s.row_dimensions[r].height=32
        for col in range(1,len(THEAD)+1):
            c=s.cell(row=r,column=col); c.border=box; c.font=Font(name=F,size=9,color=INK)
            c.alignment=Alignment(vertical="center",wrap_text=(col in (1,2,7)),
                                  horizontal="center" if col in (4,5) else "left")
        s.cell(row=r,column=4).number_format="0%"
        s.cell(row=r,column=5).font=Font(name=F,size=9,bold=True,
            color={"進行中":AMBER,"確認待ち":GREEN,"未着手":GREY}[state])
        if own=="未定":
            s.cell(row=r,column=3).font=Font(name=F,size=9,bold=True,color=RED)
        for col in (3,4,5):
            s.cell(row=r,column=col).fill=PatternFill("solid",fgColor=EDIT)
        s.cell(row=r,column=6).fill=PatternFill("solid",fgColor=INPUT)
        s.cell(row=r,column=6).font=Font(name=F,size=9,color="0000FF")
        r+=1
    f,l=6,r-1
    s.conditional_formatting.add("D{0}:D{1}".format(f,l),
        DataBarRule(start_type="num",start_value=0,end_type="num",end_value=1,color=ACC,showValue=True))
    dv(s,"list","C{0}:C{1}".format(f,l),OWNER_LIST,"担当を選ぶ","一覧にない場合は直接入力もできます",blank=True)
    dv(s,"list","D{0}:D{1}".format(f,l),PCT_LIST,"進捗を選ぶ","5%きざみで選べます。直接入力もできます。",blank=True)
    dv(s,"list","E{0}:E{1}".format(f,l),STATE_LIST,"状態を選ぶ","未着手／進行中／確認待ち")
    t=r
    c=s.cell(row=t,column=3,value="平均"); c.font=Font(name=F,size=9,bold=True,color=INK)
    c.alignment=Alignment(horizontal="right")
    c=s.cell(row=t,column=4,value="=AVERAGE(D{0}:D{1})".format(f,l))
    c.font=Font(name=F,size=9,bold=True,color=ACC); c.number_format="0%"
    c.alignment=Alignment(horizontal="center")
    c=s.cell(row=t,column=5,value='=COUNTIF(E{0}:E{1},"未着手")&" 件未着手"'.format(f,l))
    c.font=Font(name=F,size=9,bold=True,color=ACC)
    c=s.cell(row=t,column=6,value='=COUNTA(F{0}:F{1})&" 件 登録ずみ"'.format(f,l))
    c.font=Font(name=F,size=9,bold=True,color=ACC)
    for col in range(3,7):
        s.cell(row=t,column=col).border=Border(top=Side(style="medium",color=ACC))
    s.cell(row=f,column=3).comment=Comment(
      "クリックすると▼が出ます。一覧から選ぶか、直接書いてください。","制作進捗ボード",height=70,width=240)
    s.freeze_panes="B6"; s.auto_filter.ref="A5:G{0}".format(l); s.sheet_view.showGridLines=False
    return s

edu=[
 ("eラーニング（ツール本体）","LMP事業 · Web・ツール・システム","社長・白石・片山",50,"進行中","加盟店が学ぶ画面そのもの"),
 ("eラーニング動画","LMP事業 · Web・ツール・システム","未定",20,"進行中","教材の中で流す動画。担当が決まっていない"),
 ("教育診断","RC事業 · Web・ツール・システム","白石",80,"進行中","学ぶ前に現在地をはかる"),
 ("教科書の前段階の資料","RC事業 · 動画・コンテンツ","白石",60,"進行中","教材のもとになる原稿"),
 ("各ツールマニュアル作成","LMP事業 · 営業・資料","未定",0,"未着手","担当が決まっていない"),
 ("研修会準備","LMP事業 · 加盟開発","白石・池之上・大熊",0,"未着手",""),
 ("eラーニング接続・設計","RCP事業 · Web・ツール・システム","白石",0,"未着手",""),
 ("AIロープレ（アオ先生）","AIアオ先生 · 営業コーチ","池之上",90,"確認待ち","制作MASTERのみに記載"),
 ("結婚後の暮らしと人生設計動画","RC事業 · 動画・コンテンツ","白石",70,"進行中","動画で唯一動いている1本"),
 ("サービス説明動画","RC事業 · 動画・コンテンツ","白石",0,"未着手",""),
 ("toC動画","RC事業 · 動画・コンテンツ","白石",0,"未着手",""),
 ("サービスデモ動画","RCP事業 · 動画・コンテンツ","白石",0,"未着手",""),
 ("水谷先生の本の動画","RCP事業 · 動画・コンテンツ","白石",0,"未着手",""),
 ("OEM各先生の宣伝動画（5本）","RCP PPモデル · OEM","白石",0,"未着手","AKさん・大橋会長・木原さん・水谷先生・藤田先生"),
 ("宣伝動画・PV ロング／ショート（7本）","RC／RCP／HI／PP · REALIZE QUEST","白石・片山",0,"未着手",""),
 ("採用まわりの動画（2本）","採用 · 学生募集","白石・片山",0,"未着手","PR動画・SNS採用募集動画"),
]
s2=task_sheet("eラーニング・動画","制作進捗ボード ─ eラーニング・動画",
  "工程表（制作MASTER 2026-09-03 ／ 仕事の棚卸し 2026-09-02）の進捗をそのまま引いています。",
  "ツールの箱は50%まで来ていますが、中に入れる動画が20%で止まっています。赤い担当名は「未定」です。",edu)
r=s2.max_row+2
s2.cell(row=r,column=1,value="できている設計書・ガイド").font=Font(name=F,size=11,bold=True,color=ACC)
r+=1
for nm,kind,upd,url in [
 ("LIFE CHECK42 Eラーニング画面設計書","画面設計","2026-08-31",A+"1ce7243d-ad99-47d0-b0d0-ad85c8cc7f16"),
 ("RCP ノーマル版 eラーニング｜UI提案 v1","UI提案","2026-08-05",A+"0ebdd065-594c-447b-a2d6-c1eb4c80d933"),
 ("AIロープレ 操作ガイド","操作ガイド","2026-08-31",A+"fc25e6a1-4589-43ce-8f16-5b416e3fc619"),
 ("生成AI動画講座 比較検討","比較検討","2026-08-21",A+"98260a43-003c-4653-af3c-d667139ea6f8"),
]:
    s2.cell(row=r,column=1,value=nm).font=Font(name=F,size=9,color=INK)
    s2.cell(row=r,column=2,value=kind).font=Font(name=F,size=9,color=GREY)
    s2.cell(row=r,column=3,value=upd).font=Font(name=F,size=9,color=GREY)
    c=s2.cell(row=r,column=7,value="開く"); c.hyperlink=url
    c.font=Font(name=F,size=9,color=ACC,underline="single")
    r+=1

ad=[
 ("広告制作","RC法人開拓 · 地域企業への入口","大熊・片山",0,"未着手","バナー・原稿などのクリエイティブ"),
 ("広告運用","RC法人開拓／LMP加盟開発","池之上・片山",0,"未着手","出稿・予算・改善"),
 ("広告クリエイティブチェック","LMP事業 · 全体品質確認","白石",0,"未着手","出す前の品質確認"),
 ("広告表示のチェック（景表法）","社内業務 · 法務・業法","社長・寺井・池之上・弁護士",0,"未着手","言い切り表現・実績表示の法務確認"),
 ("広報・告知（LP・SNS）4件","RC／RCP／HI／PP · REALIZE QUEST","白石",0,"未着手",""),
]
task_sheet("広告","制作進捗ボード ─ 広告",
  "工程表上のどの項目も0%。制作も運用もチェック体制もまだ始まっていません。",
  "広告を出すには、飛ばす先のLPが公開され、問い合わせが届く状態である必要があります。着手日はLP公開日で決まります。",ad)

# ================================================== 凡例
lg=wb.create_sheet("凡例と出典"); lg.sheet_view.showGridLines=False
lg.column_dimensions["A"].width=26; lg.column_dimensions["B"].width=98
lg["A1"]="凡例と出典"; lg["A1"].font=Font(name=F,size=15,bold=True,color=INK)
def head(r,t): lg.cell(row=r,column=1,value=t).font=Font(name=F,size=12,bold=True,color=ACC)
def kv(r,k,v):
    a=lg.cell(row=r,column=1,value=k); a.font=Font(name=F,size=9,bold=True,color=INK)
    a.alignment=Alignment(vertical="top")
    b=lg.cell(row=r,column=2,value=v); b.font=Font(name=F,size=9,color=INK)
    b.alignment=Alignment(vertical="top",wrap_text=True); lg.row_dimensions[r].height=34
head(3,"LPの3段")
kv(4,"原稿・デザイン","本編の文章とデザインが最後まで通っている")
kv(5,"問い合わせ導線","ボタンの行き先が実在し、問い合わせが届く先につながっている")
kv(6,"本番公開","独自ドメインで公開され、G列に本番URLが入っている（自動判定）。○はそのページへのリンクになり、クリックすると開きます。")
head(8,"記号")
kv(9,"○","できている")
kv(10,"△","仮づけ。メーラーを開くだけ、または行き先が「#」のまま")
kv(11,"×","手つかず／ページが存在しない")
head(13,"状態")
kv(14,"確認待ち","作業は終わり、誰かの確認を待っている")
kv(15,"進行中","いま手が動いている")
kv(16,"未着手","まだ始まっていない")
kv(17,"赤い担当名","担当が「未定」の仕事。誰かを決めないと進みません")
head(19,"クリックして選べるところ（水色のセル）")
kv(20,"案件","LPシートのA列。どの案件のLPかを選びます。")
kv(21,"担当","タスクシートのC列。一覧から選ぶか、直接書き換えます。")
kv(22,"進捗","タスクシートのD列。0%〜100%を5%きざみで選べます。選ぶとセルの中のバーが伸び縮みします。")
kv(23,"状態","タスクシートのE列。未着手／進行中／確認待ち。")
kv(24,"LPの○△×","LPシートのD・E列。本番公開（F列）はG列のURL有無から自動で決まります。")
kv(25,"黄色の記入欄","公開したURL・動画URL・広告の入稿先を貼ります。LPシートは、貼るとF列の○がそのページへのリンクになります。")
kv(26,"列見出しの▼","オートフィルタです。値を選ぶとその行だけ表示されます。複数の列を組み合わせられます。")
kv(27,"選択肢を増やしたいとき","非表示シート「選択肢」に一覧があります。値を足し、入力規則の参照範囲を広げれば選べるようになります。")
head(29,"出典")
kv(30,"LPシート","公開済みアーティファクト43件のうちLPに該当する22ページ（本編11＋付随11）を、1ページずつHTMLを読んで確認。ボタンの行き先・フォームの送信方法・付随ページの有無は実物ベースです。")
kv(31,"進捗％と担当","社内の「REALIZE OS 制作MASTER」(2026-09-03) と「仕事の棚卸し」(2026-09-02) の数値をそのまま引いています。")
kv(32,"担当の食い違い","広告運用の担当は2資料で記載が違います（制作MASTER＝大熊・片山／棚卸し＝池之上・片山）。広告制作は大熊・片山、広告運用は池之上・片山としています。")
kv(33,"未確認","LP付随ページ4本（RC星空案の会社概要、RCアイス案の資料請求、LMP星空案の資料請求と会社概要）は、親LPからのリンクで存在を確認しただけで中身は開いていません。")
kv(34,"調査日","2026-09-03")
for r in list(range(4,7))+[9,10,11]+list(range(14,18))+list(range(20,28))+list(range(30,35)):
    lg.row_dimensions[r].height=36

opt.sheet_state="hidden"
wb.calculation.fullCalcOnLoad=True
wb.save("/home/user/jigyokeikakusyo/制作進捗ボード.xlsx")
print("saved")
