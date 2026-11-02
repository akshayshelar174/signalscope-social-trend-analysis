from pathlib import Path

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from pptx import Presentation
from pptx.util import Inches as PI, Pt as PPT
from pptx.dml.color import RGBColor as PRGB
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_THEME_COLOR

NAVY='142331'; BLUE='536EF2'; MINT='65CDA7'; INK='203040'; MUTED='657483'; LIGHT='F4F6F8'; PALE='EBEFFD'; GOLD='DCA653'

def shade(cell, fill):
    tcPr=cell._tc.get_or_add_tcPr(); shd=OxmlElement('w:shd'); shd.set(qn('w:fill'),fill); tcPr.append(shd)

def set_cell_text(cell,text,bold=False,color=INK,size=9):
    cell.text=''; p=cell.paragraphs[0]; p.paragraph_format.space_after=Pt(0)
    r=p.add_run(text); r.bold=bold; r.font.name='Aptos'; r.font.size=Pt(size); r.font.color.rgb=RGBColor.from_string(color)
    cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER

def make_report():
    d=Document(); sec=d.sections[0]; sec.top_margin=Inches(.72); sec.bottom_margin=Inches(.7); sec.left_margin=Inches(.82); sec.right_margin=Inches(.82)
    styles=d.styles
    styles['Normal'].font.name='Aptos'; styles['Normal'].font.size=Pt(10); styles['Normal'].font.color.rgb=RGBColor.from_string(INK)
    styles['Normal'].paragraph_format.space_after=Pt(6); styles['Normal'].paragraph_format.line_spacing=1.12
    for name,size,col in [('Title',30,NAVY),('Heading 1',19,NAVY),('Heading 2',13,BLUE)]:
        s=styles[name]; s.font.name='Aptos Display'; s.font.size=Pt(size); s.font.bold=True; s.font.color.rgb=RGBColor.from_string(col); s.paragraph_format.space_before=Pt(13); s.paragraph_format.space_after=Pt(6)
    hp=sec.header.paragraphs[0]; hp.text='SIGNALSCOPE     /     SOCIAL INTELLIGENCE'; hp.style='Caption'; hp.runs[0].font.color.rgb=RGBColor.from_string(MUTED); hp.runs[0].font.size=Pt(8)
    fp=sec.footer.paragraphs[0]; fp.alignment=WD_ALIGN_PARAGRAPH.RIGHT; rr=fp.add_run('SIGNALSCOPE  •  DEMO CONCEPT  •  OCTOBER 2026'); rr.font.size=Pt(8); rr.font.color.rgb=RGBColor.from_string(MUTED)
    p=d.add_paragraph(); p.paragraph_format.space_before=Pt(43); r=p.add_run('PRODUCT CONCEPT  /  01'); r.bold=True; r.font.size=Pt(9); r.font.color.rgb=RGBColor.from_string(BLUE)
    d.add_paragraph('Social Media\nTrend Analysis',style='Title')
    p=d.add_paragraph('SignalScope turns fast-moving social conversations into a clear view of what is rising, why it matters, and what to do next.'); p.paragraph_format.space_after=Pt(20); p.runs[0].font.size=Pt(15); p.runs[0].font.color.rgb=RGBColor.from_string(MUTED)
    t=d.add_table(rows=1,cols=1); t.alignment=WD_TABLE_ALIGNMENT.LEFT; cell=t.cell(0,0); shade(cell,NAVY); set_cell_text(cell,'A product brief, measurement framework, and launch roadmap\nIllustrative demo metrics throughout; no live platform data is connected.',False,'FFFFFF',10)
    d.add_paragraph('Prepared October 2026',style='Subtitle')
    d.add_page_break()
    d.add_heading('Executive summary',level=1)
    d.add_paragraph('Social teams face a discovery problem: useful signals are spread across formats, platforms, and conversations, while raw volume is easy to mistake for a durable trend. SignalScope is a proposed web dashboard that organizes listening data into a repeatable workflow: monitor topic volume, detect acceleration, understand audience response, compare channels, and convert evidence into a testable action.')
    d.add_paragraph('The accompanying website is a front-end prototype. Its sample metrics and trend names are illustrative, designed to show how the experience could work. They are not observations of current platform activity and should not be used for business decisions.')
    d.add_heading('The opportunity',level=2)
    d.add_paragraph('Teams need a shared view that connects a signal to evidence and a next step. A useful trend brief should answer four questions: what is changing, where is it happening, who is engaging, and what small action can validate the opportunity?')
    # Key principles table
    t=d.add_table(rows=1,cols=3); t.alignment=WD_TABLE_ALIGNMENT.CENTER; t.style='Table Grid'
    for i,h in enumerate(['DISCOVER','INTERPRET','ACT']): shade(t.rows[0].cells[i],NAVY); set_cell_text(t.rows[0].cells[i],h,True,'FFFFFF',9)
    vals=['Spot topics accelerating beyond their own baseline.','Check sentiment, audience, format, and channel context.','Turn a signal into a low-cost content or research test.']
    for i,x in enumerate(vals): set_cell_text(t.add_row().cells[i],x,False,INK,9)
    d.add_heading('Product concept',level=1)
    d.add_paragraph('SignalScope combines a high-level pulse with a trend explorer and practical recommendations. The design is intended for marketers, social strategists, and small brand teams that need a concise daily or weekly readout.')
    d.add_heading('Core dashboard areas',level=2)
    areas=[('Overview','Mention volume, engagement, emerging topics, and sentiment in one scan.'),('Trend explorer','Ranked topics with momentum, mentions, engagement, and status.'),('Audience pulse','Estimated positive, neutral, negative, and unclassified conversation shares.'),('Channel mix','Relative share by platform, with filters to compare patterns.'),('Insights','Short evidence-linked observations paired with suggested experiments.')]
    for title,desc in areas:
        p=d.add_paragraph(style='List Bullet'); r=p.add_run(title+' — '); r.bold=True; p.add_run(desc)
    d.add_heading('Illustrative dashboard readout',level=2)
    d.add_paragraph('The prototype shows 128.4K tracked mentions, a 6.4% engagement rate, 12 emerging topics, and 62% positive sentiment. The sample channel mix is Instagram 35%, TikTok 29%, YouTube 22%, and X 14%. These numbers demonstrate dashboard components only; they do not describe actual traffic.')
    t=d.add_table(rows=1,cols=4); t.style='Table Grid'; t.alignment=WD_TABLE_ALIGNMENT.CENTER
    for i,h in enumerate(['SAMPLE TOPIC','MENTIONS','ENGAGEMENT','STATUS']): shade(t.rows[0].cells[i],PALE); set_cell_text(t.rows[0].cells[i],h,True,BLUE,8)
    for row in [('Creator-led reviews','24.8K','8.2%','Rising'),('Quiet luxury','18.2K','7.1%','Rising'),('AI editing tools','15.6K','5.8%','Emerging'),('Refill routines','9.3K','6.5%','Watch'),('Sunday reset','7.9K','6.9%','Steady')]:
        cells=t.add_row().cells
        for i,v in enumerate(row): set_cell_text(cells[i],v,i==0,INK,8)
    d.add_heading('Measurement framework',level=1)
    d.add_paragraph('A trend score should reward sustained, unusual growth while controlling for topic size, platform coverage, and noisy spikes. The prototype uses a simple ranked view; a production model should keep each input visible so analysts can explain why a topic moved.')
    t=d.add_table(rows=1,cols=3); t.style='Table Grid'
    for i,h in enumerate(['LAYER','MEASURES','WHY IT MATTERS']): shade(t.rows[0].cells[i],NAVY); set_cell_text(t.rows[0].cells[i],h,True,'FFFFFF',8)
    for row in [('Volume','Unique mentions; share of voice','Shows scale and brand/category context.'),('Momentum','7-day vs. 30-day growth; acceleration','Distinguishes rising topics from large but flat topics.'),('Response','Engagement per post; sentiment mix','Adds audience reaction, not just raw activity.'),('Coverage','Platforms, creators, geographies','Reveals where the signal is concentrated.'),('Confidence','Source count, sampling, duplicate rate','Signals when the evidence is too thin to act.')]:
        cells=t.add_row().cells
        for i,v in enumerate(row): set_cell_text(cells[i],v,i==0,INK,8)
    d.add_heading('Recommended operating rhythm',level=2)
    for x in ['Daily scan: review unusual acceleration and possible safety or reputation issues.','Weekly brief: select two or three signals, document evidence, and assign experiments.','Monthly review: compare content tests with baseline outcomes and retire weak signals.']:
        d.add_paragraph(x,style='List Bullet')
    d.add_heading('Data, quality, and responsible use',level=1)
    d.add_paragraph('A live product depends on authorized data access and platform terms. Source coverage will vary, so the interface should name connected sources and show missing or delayed data. Avoid collecting private content or inferring sensitive traits. Aggregate where possible, set retention limits, and provide traceable examples for analysts to inspect.')
    d.add_paragraph('Sentiment classification is uncertain, especially for sarcasm, slang, code-switching, and context-specific language. Treat it as a directional estimate, display confidence and unclassified share, and allow human review. Deduplicate reposts and separate paid, owned, and earned activity where the source permits.')
    d.add_heading('Launch roadmap',level=1)
    t=d.add_table(rows=1,cols=3); t.style='Table Grid'
    for i,h in enumerate(['PHASE','DELIVERABLE','EXIT CHECK']): shade(t.rows[0].cells[i],NAVY); set_cell_text(t.rows[0].cells[i],h,True,'FFFFFF',8)
    for row in [('1 · Validate','Interview target teams; define top decisions and sources.','Two priority workflows and source access confirmed.'),('2 · Instrument','Connect permitted sources; define taxonomy and baselines.','Coverage, deduplication, and refresh status visible.'),('3 · Pilot','Run a small cohort through weekly trend briefs.','Analysts can reproduce and explain ranked signals.'),('4 · Evaluate','Compare experiments with baseline results.','Adoption, action rate, and outcome measures reviewed.')]:
        cells=t.add_row().cells
        for i,v in enumerate(row): set_cell_text(cells[i],v,i==0,INK,8)
    d.add_heading('Success measures',level=2)
    d.add_paragraph('Track time from signal to review, share of reviewed signals that become an experiment, analyst agreement on trend relevance, source coverage, and whether content tests outperform an appropriate baseline. Do not optimize solely for mentions or clicks; include audience quality and brand-fit checks.')
    d.add_heading('Conclusion',level=1)
    d.add_paragraph('SignalScope is a practical starting point for a social trend analysis product: make the signal visible, show the evidence, and connect each recommendation to a measurable test. The next step is to validate the workflows with target users and establish permitted source integrations before replacing the prototype’s illustrative data.')
    root = Path(__file__).resolve().parents[1]
    (root / 'report').mkdir(parents=True, exist_ok=True)
    d.save(root / 'report' / 'SignalScope_Report.docx')

def add_text(slide,x,y,w,h,text,size=16,color=INK,bold=False,font='Aptos',align=PP_ALIGN.LEFT):
    box=slide.shapes.add_textbox(PI(x),PI(y),PI(w),PI(h)); tf=box.text_frame; tf.clear(); tf.word_wrap=True
    p=tf.paragraphs[0]; p.alignment=align; p.space_after=PPT(0); r=p.add_run(); r.text=text; r.font.name=font; r.font.size=PPT(size); r.font.bold=bold; r.font.color.rgb=PRGB.from_string(color); return box

def rect(slide,x,y,w,h,fill, radius=False,line=None):
    shape=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE,PI(x),PI(y),PI(w),PI(h)); shape.fill.solid(); shape.fill.fore_color.rgb=PRGB.from_string(fill); shape.line.fill.background() if not line else None
    if line: shape.line.color.rgb=PRGB.from_string(line)
    if radius:
        try: shape.adjustments[0]=.08
        except: pass
    return shape

def base_slide(prs,title,kicker='SIGNALSCOPE  /  PRODUCT CONCEPT'):
    s=prs.slides.add_slide(prs.slide_layouts[6]); s.background.fill.solid(); s.background.fill.fore_color.rgb=PRGB.from_string('F7F8FA')
    rect(s,0,0,.13,7.5,BLUE)
    add_text(s,.55,.32,11,.25,kicker,9,BLUE,True)
    add_text(s,.55,.72,12, .6,title,27,NAVY,True,'Aptos Display')
    add_text(s,.55,7.12,9,.18,'ILLUSTRATIVE SAMPLE DATA  ·  OCTOBER 2026',8,MUTED,False)
    add_text(s,12.05,7.1,.7,.2,f'{len(prs.slides):02d}',8,MUTED,True,align=PP_ALIGN.RIGHT)
    return s

def bullet_list(slide,x,y,w,items,size=15,color=INK,gap=.65):
    for i,item in enumerate(items):
        yy=y+i*gap; rect(slide,x,yy+.09,.1,.1,BLUE,True); add_text(slide,x+.26,yy,w-.26,.46,item,size,color)

def make_ppt():
    prs=Presentation(); prs.slide_width=PI(13.333); prs.slide_height=PI(7.5)
    # Cover
    s=prs.slides.add_slide(prs.slide_layouts[6]); s.background.fill.solid(); s.background.fill.fore_color.rgb=PRGB.from_string(NAVY)
    rect(s,9.1,0,4.24,7.5,'192F41'); rect(s,9.1,0,.13,7.5,MINT)
    add_text(s,.8,.75,6,.3,'PRODUCT CONCEPT     /     01',10,MINT,True)
    add_text(s,.8,1.7,7.7,1.9,'Social Media\nTrend Analysis',36,'FFFFFF',True,'Aptos Display')
    add_text(s,.85,4.08,6.7,.85,'SignalScope turns social conversations into a clear view of what is rising, why it matters, and what to do next.',18,'C9D2DC')
    rect(s,.85,5.5,2.4,.44,BLUE,True); add_text(s,1.02,5.58,2.1,.2,'PRODUCT BRIEF · OCT 2026',9,'FFFFFF',True)
    # decorative visual
    for i,(h,c) in enumerate([(1.4,'405A70'),(2.2,'58749A'),(3.1,'697EF4'),(2.4,'69CDA8'),(3.8,'697EF4'),(4.5,'75DAB2')]): rect(s,9.75+i*.48,5.75-h,.25,h,c,True)
    add_text(s,9.7,1.4,2.8,.5,'SIGNAL',12,'A8B6C4',True)
    add_text(s,9.7,2.05,2.8,.5,'→  INSIGHT',14,'FFFFFF',True)
    add_text(s,9.7,2.75,2.8,.5,'→  ACTION',14,MINT,True)
    add_text(s,.85,7.05,5,.2,'DEMO METRICS ARE ILLUSTRATIVE; NOT LIVE PLATFORM DATA',8,'A8B6C4')
    # 2 problem
    s=base_slide(prs,'The signal is scattered across the feed')
    add_text(s,.7,1.65,5.3,.8,'Teams can see a lot of activity and still miss what is changing.',21,NAVY,True)
    bullet_list(s,.75,2.75,5.2,['Topics rise across different platforms and formats.','Raw volume rewards large conversations, even when they are flat.','Sentiment without context can mislead a content decision.','Insights often stop at a chart instead of becoming a test.'],14)
    rect(s,6.7,1.8,5.6,4.45,'FFFFFF',True,'E6EAF0')
    add_text(s,7.15,2.15,4.8,.35,'THE DECISION LOOP',10,BLUE,True)
    for i,(num,title,desc) in enumerate([('01','Notice','What is accelerating?'),('02','Contextualize','Where and who is driving it?'),('03','Test','What can we learn this week?')]):
        yy=2.8+i*1.0; rect(s,7.15,yy,.42,.42,PALE,True); add_text(s,7.23,yy+.1,.3,.18,num,9,BLUE,True); add_text(s,7.8,yy,3,.27,title,15,NAVY,True); add_text(s,7.8,yy+.35,3.8,.3,desc,11,MUTED)
    # 3 product
    s=base_slide(prs,'SignalScope makes trend work a repeatable workflow')
    add_text(s,.7,1.55,11,.55,'One concise workspace connects discovery, audience response, and next steps.',17,MUTED)
    cols=[('DISCOVER','Monitor topic volume and momentum.','Ranked trends · watchlists'),('INTERPRET','Compare sentiment, formats, and channels.','Audience pulse · channel mix'),('ACT','Turn evidence into a small experiment.','Insights · weekly brief')]
    for i,(h,b,foot) in enumerate(cols):
        x=.75+i*4.15; rect(s,x,2.55,3.7,2.7,'FFFFFF',True,'E6EAF0'); add_text(s,x+.25,2.85,3,.3,f'0{i+1}   {h}',11,BLUE,True); add_text(s,x+.25,3.45,3.1,.75,b,17,NAVY,True); add_text(s,x+.25,4.55,3.1,.35,foot,10,MUTED)
    add_text(s,.85,5.8,11,.6,'Designed for marketing and social teams that need a useful daily scan and a credible weekly readout.',15,INK)
    # 4 overview mock
    s=base_slide(prs,'The overview answers “what changed?”')
    rect(s,.75,1.65,11.75,4.9,'FFFFFF',True,'E6EAF0')
    add_text(s,1.05,1.9,4,.3,'YOUR SOCIAL PULSE',11,NAVY,True)
    for i,(lab,val,delta) in enumerate([('MENTIONS','128.4K','+18.6%'),('ENGAGEMENT','6.4%','+0.8 pts'),('EMERGING TOPICS','12','+3 new'),('POSITIVE','62%','+4 pts')]):
        x=1.05+i*2.75; rect(s,x,2.45,2.4,.95,'F8F9FB',True,'EDF0F3'); add_text(s,x+.15,2.59,2.1,.2,lab,8,MUTED,True); add_text(s,x+.15,2.82,1.2,.3,val,20,NAVY,True,'Aptos Display'); add_text(s,x+1.25,2.88,1,.2,delta,9,'17845E',True)
    add_text(s,1.1,3.8,4,.3,'MENTIONS OVER TIME',9,MUTED,True)
    # line illustration using connectors
    vals=[4.5,4.25,4.35,3.75,3.95,3.5,3.6,3.05,3.2,2.75,2.95,2.35,2.5,2.0]
    for i in range(len(vals)-1):
        x1=1.2+i*.78; x2=1.2+(i+1)*.78
        sh=s.shapes.add_connector(1,PI(x1),PI(vals[i]),PI(x2),PI(vals[i+1])); sh.line.color.rgb=PRGB.from_string(BLUE); sh.line.width=PPT(2.4)
    add_text(s,1.1,4.95,9,.45,'A weekly comparison makes acceleration visible without hiding the reporting window.',11,MUTED)
    rect(s,10.05,3.7,2.05,1.35,'EEF8F3',True); add_text(s,10.25,3.9,1.6,.22,'TOP SIGNAL',8,'17845E',True); add_text(s,10.25,4.2,1.65,.55,'Creator-led\nreviews',14,NAVY,True)
    add_text(s,.85,6.75,11,.25,'Illustrative values shown to explain the interface. Production data requires authorized source connections.',9,MUTED)
    # 5 rank
    s=base_slide(prs,'Rank by momentum, then inspect the evidence')
    headers=['TOPIC','MENTIONS','ENGAGEMENT','STATUS']; widths=[5.25,2.2,2.2,1.6]; x0=.9; yy=1.8
    xx=x0
    for h,w in zip(headers,widths): add_text(s,xx,yy,w,.3,h,9,MUTED,True); xx+=w
    dataset=[('Creator-led reviews','24.8K','8.2%','Rising'),('Quiet luxury','18.2K','7.1%','Rising'),('AI editing tools','15.6K','5.8%','Emerging'),('Refill routines','9.3K','6.5%','Watch'),('Sunday reset','7.9K','6.9%','Steady')]
    for idx,row in enumerate(dataset):
        y=2.22+idx*.72; rect(s,.82,y,11.8,.57,'FFFFFF',True,'E9EDF1'); xx=1.0
        for j,(v,w) in enumerate(zip(row,widths)):
            col=BLUE if j==3 and v in ('Rising','Emerging') else ('B87B29' if j==3 and v=='Watch' else INK)
            add_text(s,xx,y+.16,w-.1,.22,v,11,col,j==0 or j==3); xx+=w
    add_text(s,.95,6.15,10.8,.55,'Momentum is a signal, not a verdict. Require source coverage, baseline comparison, and a human relevance check before acting.',12,MUTED)
    # 6 interpret
    s=base_slide(prs,'Separate audience response from topic volume')
    # donut approximation with wedges shapes not easy; use circles and label distribution bars
    rect(s,.75,1.75,5.5,4.6,'FFFFFF',True,'E6EAF0'); add_text(s,1.05,2.05,4,.3,'SENTIMENT MIX · SAMPLE',10,NAVY,True)
    parts=[('Positive',62,BLUE),('Neutral',24,MINT),('Negative',10,GOLD),('Unclassified',4,'CBD2DB')]
    for i,(name,val,col) in enumerate(parts):
        y=2.75+i*.75; add_text(s,1.1,y,1.6,.22,name,11,INK); rect(s,2.75,y+.02,2.45,.16,'EDF0F3',True); rect(s,2.75,y+.02,2.45*val/100,.16,col,True); add_text(s,5.35,y, .6,.22,f'{val}%',10,NAVY,True,align=PP_ALIGN.RIGHT)
    add_text(s,1.1,5.85,4.6,.28,'Estimates need confidence and human review.',9,MUTED)
    rect(s,6.65,1.75,5.65,4.6,'FFFFFF',True,'E6EAF0'); add_text(s,7.0,2.05,4.7,.3,'CHANNEL SHARE · SAMPLE',10,NAVY,True)
    for i,(name,val,sym) in enumerate([('Instagram',35,'◎'),('TikTok',29,'♪'),('YouTube',22,'▶'),('X',14,'𝕏')]):
        y=2.72+i*.77; add_text(s,7.05,y,1.5,.24,f'{sym}  {name}',11,INK); rect(s,8.9,y+.03,2.3,.16,'EDF0F3',True); rect(s,8.9,y+.03,2.3*val/40,.16,BLUE if i<2 else MINT,True); add_text(s,11.45,y,.5,.22,f'{val}%',10,NAVY,True,align=PP_ALIGN.RIGHT)
    add_text(s,7.05,5.85,4.6,.28,'Coverage varies by source and access.',9,MUTED)
    # 7 action
    s=base_slide(prs,'A useful insight ends with a test')
    add_text(s,.8,1.65,6.2,.4,'Sample signal',10,BLUE,True)
    add_text(s,.8,2.15,6.0,1.1,'Creator-led reviews pair rising conversation with above-average engagement.',22,NAVY,True,'Aptos Display')
    add_text(s,.8,3.55,5.9,.85,'A directional observation from the synthetic sample. Validate with audience, source, and content-level context.',13,MUTED)
    rect(s,7.1,1.8,5.1,4.2,'FFFFFF',True,'E6EAF0')
    add_text(s,7.5,2.15,4.2,.3,'LOW-COST EXPERIMENT',10,BLUE,True)
    bullet_list(s,7.55,2.75,4.2,['Publish two short comparison formats.','Track saves, qualified comments, and completion.','Compare with recent posts in the same category.','Keep, revise, or stop after the agreed window.'],13,gap=.67)
    rect(s,.85,5.3,5.6,.55,NAVY,True); add_text(s,1.1,5.47,5.1,.18,'DECISION: TEST THE FORMAT, NOT THE TREND LABEL',9,'FFFFFF',True)
    # 8 measurement
    s=base_slide(prs,'Build trust into the trend score')
    headers=['INPUT','QUESTION IT ANSWERS','GUARDRAIL']; ws=[2.4,4.3,4.35]; xx=.9
    for h,w in zip(headers,ws): add_text(s,xx,1.72,w,.28,h,9,MUTED,True); xx+=w
    items=[('Volume','How large is the conversation?','Deduplicate; separate source types.'),('Momentum','Is growth unusual and sustained?','Compare to topic and seasonal baseline.'),('Response','How are people reacting?','Show sentiment confidence and unknowns.'),('Coverage','Where is the signal concentrated?','Expose connected sources and gaps.'),('Confidence','Is evidence sufficient to act?','Display sample size and freshness.')]
    for i,row in enumerate(items):
        y=2.12+i*.72; rect(s,.82,y,11.8,.56,'FFFFFF',True,'E9EDF1'); xx=1.0
        for j,(v,w) in enumerate(zip(row,ws)): add_text(s,xx,y+.16,w-.15,.22,v,11,BLUE if j==0 else INK,j==0); xx+=w
    add_text(s,.95,6.1,11,.45,'Do not optimize a single score. Keep the ingredients visible so an analyst can reproduce and explain a ranking.',12,MUTED)
    # 9 roadmap
    s=base_slide(prs,'Validate the workflow before scaling the data')
    phases=[('01 · Validate','Interview target teams.\nSelect priority decisions.','Workflow and source access confirmed.'),('02 · Instrument','Connect permitted sources.\nSet taxonomy and baselines.','Coverage and freshness visible.'),('03 · Pilot','Run weekly briefs with a small cohort.','Ranked signals reproducible.'),('04 · Evaluate','Compare tests against baselines.','Adoption and outcomes reviewed.')]
    for i,(h,b,c) in enumerate(phases):
        x=.78+i*3.15; rect(s,x,2.0,2.75,3.4,'FFFFFF',True,'E6EAF0'); add_text(s,x+.2,2.3,2.35,.28,h,11,BLUE,True); add_text(s,x+.2,2.95,2.3,.8,b,14,NAVY,True); rect(s,x+.2,4.35,2.35,.015,'E6EAF0'); add_text(s,x+.2,4.62,2.25,.56,c,10,MUTED)
    add_text(s,.9,5.95,11,.6,'Success measures: time to review · reviewed signals converted to experiments · analyst agreement · source coverage · results vs. baseline',12,INK,True)
    # 10 close
    s=prs.slides.add_slide(prs.slide_layouts[6]); s.background.fill.solid(); s.background.fill.fore_color.rgb=PRGB.from_string(NAVY)
    rect(s,0,0,.13,7.5,MINT); add_text(s,.85,.75,5,.3,'SIGNALSCOPE  /  NEXT STEP',10,MINT,True)
    add_text(s,.85,1.55,10.7,1.0,'Make the signal visible.\nMake the next move measurable.',29,'FFFFFF',True,'Aptos Display')
    add_text(s,.9,4.0,10.4,.85,'Validate the core workflow with target users, then connect permitted sources and replace the illustrative sample with traceable data.',17,'C8D2DC')
    add_text(s,.9,6.9,9,.22,'DEMO CONCEPT · SAMPLE METRICS ARE NOT LIVE PLATFORM DATA',9,'A8B6C4')
    root = Path(__file__).resolve().parents[1]
    (root / 'presentation').mkdir(parents=True, exist_ok=True)
    prs.save(root / 'presentation' / 'SignalScope_Presentation.pptx')

if __name__=='__main__':
    make_report(); make_ppt()
