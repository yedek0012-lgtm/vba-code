"""VBA modüllerini LibreOffice Basic (VBASupport) derleyicisinden geçirir; her modül kendi kütüphanesinde.
Gerekenler: LibreOffice + python3-uno. Kullanım: python3 tools/vba_compile_check.py BOM_v3.5
Not: Option Explicit (tanımsız değişken) kontrolü için tools/vba_undeclared.py kullanın."""
import uno, sys, time, glob, os, re, subprocess
from com.sun.star.beans import PropertyValue
PORT=2002
prof='file://'+os.path.abspath(os.environ.get('LO_PROFILE','/tmp/vba_cc_prof'))
p=subprocess.Popen(['soffice','--headless','--invisible','--norestore',f'-env:UserInstallation={prof}',f'--accept=socket,host=localhost,port={PORT};urp;'],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
local=uno.getComponentContext()
res=local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",local)
for i in range(60):
    try:
        ctx=res.resolve(f"uno:socket,host=localhost,port={PORT};urp;StarOffice.ComponentContext"); break
    except Exception: time.sleep(1)
smgr=ctx.ServiceManager
desktop=smgr.createInstanceWithContext("com.sun.star.frame.Desktop",ctx)
libs=smgr.createInstanceWithContext("com.sun.star.script.ApplicationScriptLibraryContainer",ctx)
try: libs.VBACompatibilityMode=True
except Exception as e: print("vbacompat",e)

mods=sorted(glob.glob(sys.argv[1]+'/*.bas'))

# Excel VBA'nın kabul etmediği ama LibreOffice'in geçirdiği yapı: modül düzeyinde bildirim
# ilk prosedürden SONRA ("Variable not defined" / "Only comments may appear after End Sub").
for f in mods:
    inproc=seen=False
    for i,l in enumerate(open(f,encoding='cp1254').read().split('\n'),1):
        t=l.strip()
        if re.match(r'(Public |Private |Friend )?(Static )?(Sub|Function|Property) ',t,re.I): inproc=seen=True; continue
        if re.match(r'End (Sub|Function|Property)\b',t,re.I): inproc=False; continue
        if seen and not inproc and t and not t.startswith("'"):
            print("FAIL",os.path.basename(f)[:-4],": satır",i,"prosedürler arasında bildirim (Excel derlemez):",t)
# Başka modüldeki Private Sub/Function çağrısı (LibreOffice modülleri ayrı derlediği için yakalamaz;
# Excel'de "Sub or Function not defined")
_src={os.path.basename(f)[:-4]:open(f,encoding='cp1254').read() for f in mods}
_priv={}; _pub=set()
for _m,_s in _src.items():
    for _mm in re.finditer(r'^[ \t]*(Private[ \t]+)?(?:Public[ \t]+|Friend[ \t]+)?(?:Sub|Function)[ \t]+(\w+)',_s,re.M|re.I):
        if _mm.group(1): _priv.setdefault(_mm.group(2).lower(),[]).append(_m)
        else: _pub.add(_mm.group(2).lower())
for _m,_s in _src.items():
    _code=re.sub(r"'.*","",_s)
    for _n,_own in _priv.items():
        if _m in _own or _n in _pub: continue
        _hit=re.search(r'\b'+_n+r'\b',_code,re.I)
        if _hit:
            print("FAIL",_m,": satır",_code[:_hit.start()].count('\n')+1,"başka modüldeki Private çağrılıyor:",_n,"(",",".join(_own),")")
names=[]
for f in mods:
    src=open(f,encoding='cp1254').read().replace('\r\n','\n')
    name=os.path.basename(f)[:-4]
    src=re.sub(r'^Attribute VB_Name.*\n','',src)
    src=src.replace('Option Private Module\n','')
    src='Option VBASupport 1\n'+src+'\nPublic Function ZZ_'+name+'()\n    ZZ_'+name+' = 42\nEnd Function\n'
    ln="CHK_"+name
    if not libs.hasByName(ln): libs.createLibrary(ln)
    libs.loadLibrary(ln)
    lib=libs.getByName(ln)
    if lib.hasByName(name): lib.removeByName(name)
    lib.insertByName(name,src)
    names.append(name)
sp=smgr.createInstanceWithContext("com.sun.star.script.provider.MasterScriptProviderFactory",ctx).createScriptProvider("")
ok=True
for name in names:
    try:
        s=sp.getScript(f"vnd.sun.star.script:CHK_{name}.{name}.ZZ_{name}?language=Basic&location=application")
        r=s.invoke((),(),())
        v=r[0] if isinstance(r,tuple) else r
        print("OK  " if v==42 else "FAIL",name, "" if v==42 else "(derlenemedi / sonuc=%r)"%(v,))
    except Exception as e:
        ok=False
        msg=str(getattr(e,'Message',e))
        print("FAIL",name,":",msg[:600].replace('\n',' | '))
try: desktop.terminate()
except Exception: pass
p.wait(timeout=30)
