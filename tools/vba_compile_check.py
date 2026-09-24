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
