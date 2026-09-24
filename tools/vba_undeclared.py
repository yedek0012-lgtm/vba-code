"""Option Explicit için statik tarama: atama / For / ReDim hedefi olup hiçbir yerde tanımlanmamış değişkenleri listeler.
Kullanım: python3 tools/vba_undeclared.py BOM_v3.5   (ReDim x(..) As T bildirimi yanlış alarm verebilir)"""
import re,glob,sys,os
mods={os.path.basename(f)[:-4]:open(f,encoding='cp1254').read().replace('\r\n','\n') for f in glob.glob(sys.argv[1]+'/*.bas')}
def join_cont(src):
    return re.sub(r' _\n\s*',' ',src)
def strip_comments_strings(line):
    out='';inq=False
    for i,ch in enumerate(line):
        if ch=='"': inq=not inq; out+='"'; continue
        if inq: continue
        if ch=="'" : break
        out+=ch
    return out
globals_=set()
modlevel={}
procs_all=set()
for m,src in mods.items():
    src=join_cont(src)
    ml=set()
    inproc=False
    for line in src.split('\n'):
        l=strip_comments_strings(line).strip()
        if re.match(r'(Public |Private |Friend |)(Static )?(Sub|Function|Property \w+) ',l): inproc=True; 
        mm=re.match(r'(?:Public |Private |)(?:Sub|Function) (\w+)',l)
        if mm: procs_all.add(mm.group(1).lower())
        if re.match(r'End (Sub|Function|Property)',l): inproc=False; continue
        if not inproc:
            mm=re.match(r'(Public|Private|Dim|Global)\s+(Const\s+)?(.+)',l)
            if mm and not re.match(r'(Public|Private)\s+(Sub|Function|Declare|Type|Enum)',l):
                for part in re.split(r',(?![^(]*\))',mm.group(3)):
                    n=re.match(r'\s*(?:WithEvents\s+)?(\w+)',part)
                    if n:
                        (globals_ if mm.group(1) in ('Public','Global') else ml).add(n.group(1).lower())
    modlevel[m]=ml
issues=[]
for m,src in mods.items():
    src=join_cont(src)
    lines=src.split('\n')
    i=0
    while i<len(lines):
        l=strip_comments_strings(lines[i]).strip()
        mm=re.match(r'(?:Public |Private |Friend |)(?:Static )?(Sub|Function) (\w+)\s*\((.*)\)',l)
        if mm:
            kind,pname,params=mm.groups()
            local={pname.lower()}
            for part in params.split(','):
                n=re.match(r'\s*(?:Optional\s+)?(?:ByVal\s+|ByRef\s+)?(?:ParamArray\s+)?(\w+)',part)
                if n: local.add(n.group(1).lower())
            body=[]
            j=i+1
            while j<len(lines) and not re.match(r'\s*End (Sub|Function)\b',strip_comments_strings(lines[j])):
                body.append((j+1,strip_comments_strings(lines[j]))); j+=1
            for ln,b in body:
                for stmt in re.split(r':(?!=)',b):
                    st=stmt.strip()
                    d=re.match(r'(Dim|Static|ReDim(?:\s+Preserve)?|Const)\s+(.+)',st)
                    if d and d.group(1) in ('Dim','Static','Const'):
                        for part in re.split(r',(?![^(]*\))',d.group(2)):
                            n=re.match(r'\s*(\w+)',part)
                            if n: local.add(n.group(1).lower())
            known=local|modlevel[m]|globals_|procs_all
            for ln,b in body:
                # split single-line If ... Then stmt
                segs=[]
                for stmt in re.split(r':(?!=)',b):
                    st=stmt.strip()
                    t=re.match(r'(?:ElseIf|If)\s+.+?\s+Then\s+(.*)$',st)
                    if t and t.group(1): 
                        rest=t.group(1)
                        segs+= re.split(r'\s+Else\s+',rest)
                    else: segs.append(st)
                for st in segs:
                    st=st.strip()
                    cands=[]
                    a=re.match(r'(?:Set\s+|Let\s+)?([A-Za-z]\w*)\s*(?:\([^=]*\))?\s*=(?!=)',st)
                    if a and not re.match(r'(If|ElseIf|Case|Do|Loop|While|Until|Select|Dim|ReDim|Const|Static|Public|Private|Return|Exit|End|Call)\b',st): cands.append(a.group(1))
                    f=re.match(r'For\s+(?:Each\s+)?([A-Za-z]\w*)',st)
                    if f: cands.append(f.group(1))
                    r=re.match(r'ReDim\s+(?:Preserve\s+)?([A-Za-z]\w*)',st)
                    if r: cands.append(r.group(1))
                    for c in cands:
                        if c.lower() not in known:
                            issues.append((m,ln,pname,c,st[:90]))
            i=j
        i+=1
for x in issues: print(*x,sep=' | ')
print(len(issues),'possible undeclared')
