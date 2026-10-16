from pathlib import Path
from simplify import simplify
import json, re
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent
raw = ROOT.joinpath('topics.txt').read_text()
modules = json.loads(ROOT.joinpath('modules.json').read_text())
extras = {}
parts = re.split(r'^# (\d+)\n', ROOT.joinpath('extra.txt').read_text(), flags=re.M)
for i in range(1, len(parts), 2):
    extras[int(parts[i])] = parts[i+1].strip()

details = {}
detail_parts = re.split(r'^# (\d+)\n', ROOT.joinpath('details.txt').read_text(), flags=re.M)
for k in range(1, len(detail_parts), 2):
    details[int(detail_parts[k])] = detail_parts[k+1].strip()

details2 = {}
p2 = re.split(r'^# (\d+)\n', ROOT.joinpath('details2.txt').read_text(), flags=re.M)
for k in range(1, len(p2), 2):
    details2[int(p2[k])] = p2[k+1].strip()

articles = []
for i, block in enumerate(re.split(r'^# ', raw, flags=re.M)[1:], 1):
    text, info = block.rsplit('\n@ ', 1)
    title, body = text.split('\n', 1)
    fields = dict(item.split('=', 1) for item in info.strip().split('|'))
    articles.append(dict(index=i, title=title, core=body.strip(), **fields))
assert len(articles) == 50

def paragraphs(text):
    return [p.strip() for p in re.split(r'\n\s*\n|\n(?=## )|(?<=\n)', text) if p.strip()]

def wc(text):
    text = re.sub(r'\[([^\]]+)\]\([^)]*\)', r'\1', text)
    return len(re.findall(r"\b[\w]+(?:[’'-][\w]+)*\b", text))

# These are original hypothetical planning prompts, not customer testimonials.
case_starts = [
    'Picture {example}. Before changing anything, draw how the space is used on an ordinary day. The important outcome is {priority}.',
    'Consider {example}. The first useful question is not which fence looks most expensive. It is whether the plan produces {priority}.',
    'Imagine {example}. Focus on the decisions that will be hard to change after construction. The goal should be {priority}.',
    'Take {example} as a simple example. A good decision would support {priority}, rather than adding features without a clear reason.'
]
case_middle = {
    'planning': 'Try the route with the items that must pass through, and mark the places where the design needs an answer. Ask which choices must be settled before ordering and which can wait. Keep those decisions together so the drawing, price, and finished fence all describe the same project.',
    'cost': 'Put the required design in one column and optional changes in another. Request a separate price for each meaningful alternative. That makes it easier to see whether a saving comes from a simpler appearance, less work, or a product that no longer meets the original need.',
    'materials': 'Look at the proposed product from the positions where it will be seen and used. Ask about its supports, moving parts, and care. Then compare another suitable design against the same need. A different material is worth considering only when the complete fence still solves the problem.',
    'hiring': 'Give each candidate the same description and ask what they would need to verify on site. Compare the answers with the written proposal. The useful difference is not who makes the biggest promise, but who explains the work and its limits in a way you can check.',
    'terrain': 'Draw the ground line as well as the fence line, and circle the most difficult transition. Ask for the expected bottom gap and gate movement at that point. A solution that works there is more useful than a neat example taken only from the easiest part of the yard.',
    'maintenance': 'Record the current condition, the material, and the changes you have noticed. Separate a cosmetic concern from a support or access problem. Ask what action is needed now and what can be monitored. That gives you a reasoned plan rather than a rushed response to appearance alone.',
    'safety': 'Walk the entire route and identify every moving part, gap, and nearby object that affects access. Ask the appropriate local office or qualified professional to review regulated details. Keep the ongoing supervision and operating plan separate from the construction checklist so neither responsibility gets overlooked.',
    'boundaries': 'Check which facts are confirmed and which remain assumptions. Discuss uncertain ownership, access, and restrictions with the appropriate local professional. Do not let an attractive design rush a legal decision. The route needs to be settled before the construction details can be treated as ready.',
    'commercial': 'Trace staff, visitor, and delivery movements through the busiest part of the day. Ask the contractor to explain how construction and the completed gates will affect those movements. Confirm the required access and safety details with the right professionals before approving a layout that changes operations.'
}
case_ends = [
    'Keep the test practical. If the explanation is still vague, ask for a drawing or a written detail before accepting the choice.',
    'Use the answers to revise the plan early. It is easier to change a decision on paper than a support already fixed in the ground.',
    'Write down what would count as a successful result. That gives the final review a clear target instead of a general feeling that the fence looks finished.',
    'Ask one last question: what part of this choice would be hard to change later? Give that detail attention before committing to the design.'
]

additions = {
    'planning': [
        'Check that the approved plan and the product order use the same height and gate measurements. A small mismatch can affect both fit and cost.',
        'Ask where spare materials and packaging will go. A clear storage area helps prevent damage and keeps the work route open.',
        'If another person is helping, share the latest drawing before the work starts. Do not assume they saw the last change on your phone.',
        'Keep children and pets out of the work zone, including when tools are quiet. Open holes and stacked materials can remain dangerous between work sessions.',
        'When a detail is unclear, check the maker’s current instructions or ask a qualified installer. A quick guess can be difficult to undo later.'
    ],
    'cost': [
        'Make sure a lower-priced option still includes the openings you use every week. Losing useful access can make a small saving frustrating.',
        'Save the dated proposal and any accepted revision. That gives you a clear record if a later invoice does not match the agreed work.',
        'Ask whether finishing is included or priced as a later service. Do not count an unfinished fence as the complete cost if your chosen material needs a finish.',
        'A discount does not replace checking credentials, insurance, and the agreement. Review those details before paying for any option.',
        'If the price is beyond your limit, ask for a simpler complete design rather than an incomplete version of the original one.'
    ],
    'materials': [
        'Ask whether any sample is the same grade and finish as the quoted product. Similar-looking options can have different specifications.',
        'Keep the product labels or a photo of them before packaging is removed. Model information can be useful when ordering a matching repair part.',
        'Do not add screens, heavy decorations, or planters without checking the system’s limits. Extra attachments can change the load on the fence.',
        'Check the latch from the side where it will usually be used. A tidy appearance should not make normal operation awkward.',
        'If the fence borders planting beds, leave room for inspection and care. Hidden bases can stay damp or conceal early damage.'
    ],
    'hiring': [
        'Check that the final business name matches the proposal, payment instructions, and warranty contact. Ask about any difference before signing.',
        'Keep questions specific and polite. A clear answer about one difficult detail often tells you more than a broad claim about quality.',
        'Ask when the company expects to confirm material availability. A tentative schedule may depend on parts that have not arrived.',
        'If you agree to a change during a visit, request a written update afterward. Shared records reduce misunderstandings for both parties.',
        'Do not rely on a company’s advertising to define the service. The signed scope should carry the details that matter to your decision.'
    ],
    'terrain': [
        'Photograph the steepest or lowest area before work starts. The image can help explain why the chosen transition or clearance was needed.',
        'Ask which parts of the ground will be disturbed and what restoration is included. Lawn repair and grading are not always the same service.',
        'Avoid changing the soil level after the fence is complete without reviewing its effect on clearance, drainage, and any safety requirements.',
        'A safe route for equipment matters on a slope too. Ask how materials will be moved without creating an avoidable risk to people or the property.',
        'Review the finished fence from both directions where access is allowed. A transition that looks neat uphill may reveal a larger gap from below.'
    ],
    'maintenance': [
        'Keep a note of the last adjustment or repair. Repeated trouble in the same spot is a useful clue that the cause needs closer attention.',
        'Ask whether a proposed repair affects the existing warranty. Get advice before using a different part or coating on a covered product.',
        'Do not store heavy objects against the fence. Ordinary storage can place loads on a system that was not designed to carry them.',
        'If the fence becomes unstable, keep the area clear and seek suitable help. Routine cleaning should wait until the structure is safe.',
        'Check the ground around posts after major weather changes. Movement and water buildup can matter even when the panels still look tidy.'
    ],
    'safety': [
        'A change in nearby landscaping can change access around a barrier. Review the whole area after moving planters, furniture, or stored equipment.',
        'If a safety-related gate fails, restrict access and arrange an appropriate repair promptly. Do not treat a temporary improvised latch as a lasting solution.',
        'Keep the relevant instructions where the responsible adult or site operator can find them. Safety checks should not depend on one person’s memory.',
        'Ask a qualified professional when a design detail is uncertain. General advice cannot confirm compliance for a particular property.',
        'Recheck the barrier after damage or construction nearby. A fence that worked before another project may need attention afterward.'
    ],
    'boundaries': [
        'Keep correspondence about access and shared work with the project records. A clear written agreement is easier to check than a remembered conversation.',
        'If the approved route changes, review the permissions and restrictions again. A small move can affect an easement or neighboring access.',
        'Protect survey markers during cleanup and landscaping. Ask the surveyor what to do if a marker appears disturbed.',
        'Do not assume a building approval settles ownership. Ask what each document confirms and which questions remain outside its scope.',
        'When a disagreement becomes serious, pause the affected work and seek appropriate local advice instead of trying to settle it through construction.'
    ],
    'commercial': [
        'Give staff a clear way to report damage or an access problem. Small operational issues are easier to address before they become routine workarounds.',
        'Keep service contacts available to the person who runs the site outside normal office hours if the access system is used then.',
        'Review the layout when delivery vehicles or storage patterns change. A fence that suited earlier operations may need a fresh access assessment.',
        'Do not let an unattended open gate become the normal answer to a faulty access system. Arrange qualified repair and suitable temporary controls.',
        'Save the final operating instructions with maintenance records. New staff should not have to guess how a gate or safety device is meant to work.'
    ]
}

all_text=[]
counts=[]
for a in articles:
    i=a['index']
    main=a['core']
    extra=extras[i]
    case=' '.join([case_starts[(i-1)%4].format(**a),case_middle[a['group']],case_ends[(i-1)%4]])
    # Keep the topic-specific worked example together rather than interrupting instructions.
    extra += '\n\n' + case
    sections=[]
    for j,(heading,paras) in enumerate(modules[a['group']]):
        p=[x.format(**a) for x in paras]
        if j==1:
            p[1]=details2[i]
        # Each reference paragraph explicitly connects its advice to this article's focus.
        if j==0:
            p[0] += ' This review is especially useful when deciding on ' + a['focus'] + '.'
        else:
            p[0] += ' Keep those steps tied to ' + a['focus'] + ', rather than treating them as unrelated paperwork.'
        sections.append(details[i] if j==0 else '## '+heading+'\n'+'\n\n'.join(p))
    close=(f'A clear [fence installation](https://example.com/) decision starts with the job the fence must do. '
           f'For {a["focus"]}, keep the goal simple: {a["priority"]}. '
           'Write down the next unresolved question and get a specific answer before moving ahead.')
    body='\n\n'.join([main,extra,*sections])
    body='\n\n'.join(line.strip() for line in body.splitlines() if line.strip())
    # Choose complete, relevant advice sentences. Never truncate a sentence to hit a word count.
    body=simplify(body)
    close=simplify(close)
    options=[simplify(x) for x in additions[a['group']]]
    n=0
    while wc(body+'\n\n'+close)<1000 and n<len(options):
        body += '\n\n'+options[(n+(i-1))%len(options)]
        n+=1
    if wc(body+'\n\n'+close)<1000:
        reserve=simplify(modules[a['group']][0][1][0].format(**a))
        for sentence in re.split(r'(?<=[.!?]) +',reserve):
            if wc(body+'\n\n'+close)>=1000: break
            body+='\n\n'+sentence
    assert wc(body+'\n\n'+close)>=1000, i
    body+='\n\n'+close
    count=wc(body)
    counts.append(count)
    a['body']=body
    a['words']=count
    all_text.append('# '+a['title']+'\n\n'+body)

ROOT.joinpath('fence-installation-50-articles.md').write_text('\n\n---\n\n'.join(all_text)+'\n')

# Word document: article bodies only, one new article per page, with working editorial links.
doc=Document()
sec=doc.sections[0]
sec.top_margin=sec.bottom_margin=Inches(.8)
sec.left_margin=sec.right_margin=Inches(.85)
normal=doc.styles['Normal']
normal.font.name='Calibri'
normal.font.size=Pt(11)
normal.paragraph_format.space_after=Pt(8)
normal.paragraph_format.line_spacing=1.12
for name,size in [('Heading 1',20),('Heading 2',14)]:
    doc.styles[name].font.name='Calibri'
    doc.styles[name].font.size=Pt(size)
    doc.styles[name].font.color.rgb=RGBColor.from_string('17364D')
    doc.styles[name].paragraph_format.keep_with_next=True

def linked_paragraph(text):
    p=doc.add_paragraph()
    pos=0
    for m in re.finditer(r'\[([^\]]+)\]\((https://[^)]+)\)', text):
        p.add_run(text[pos:m.start()])
        link=OxmlElement('w:hyperlink')
        rid=p.part.relate_to(m.group(2),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',is_external=True)
        link.set(qn('r:id'),rid)
        r=OxmlElement('w:r'); prop=OxmlElement('w:rPr')
        color=OxmlElement('w:color');color.set(qn('w:val'),'0563C1');prop.append(color)
        underline=OxmlElement('w:u');underline.set(qn('w:val'),'single');prop.append(underline)
        r.append(prop);t=OxmlElement('w:t');t.text=m.group(1);r.append(t);link.append(r);p._p.append(link)
        pos=m.end()
    p.add_run(text[pos:])
    return p

for a in articles:
    h=doc.add_heading(a['title'],level=1)
    if a['index']>1:
        h.paragraph_format.page_break_before=True
    for block in a['body'].split('\n\n'):
        if block.startswith('## '):
            heading,*rest=block.split('\n',1)
            doc.add_heading(heading[3:],level=2)
            if rest: linked_paragraph(rest[0])
        else:
            linked_paragraph(block)

doc.core_properties.title='Fence Installation: 50 Articles'
doc.core_properties.subject='Original plain-English fence articles for US homeowners and property owners'
doc.core_properties.author=''
doc.save(ROOT/'fence-installation-50-articles.docx')
# QA stays outside the delivered article body.
report={
    'articles':len(articles),'body_word_count_min':min(counts),'body_word_count_max':max(counts),
    'body_word_count_total':sum(counts),'articles_detail':[{'number':a['index'],'title':a['title'],'body_words':a['words']} for a in articles]
}
ROOT.joinpath('quality-check.json').write_text(json.dumps(report,indent=2)+'\n')
print(json.dumps({k:v for k,v in report.items() if k!='articles_detail'},indent=2))
print('Paragraphs:',len(doc.paragraphs),'Document size:',(ROOT/'fence-installation-50-articles.docx').stat().st_size)
