#!/usr/bin/env python3
"""Gerçek bir .tow dosyasından yapay kule varyantı üretir (sistemi başka kulelerle zorlamak için).

  python3 tow_varyant.py giris.tow cikis.tow varyant[,varyant...]

Varyantlar:
  dondur90   kule Z ekseni etrafında 90° döner: (x, y) -> (-y, x). Ön ve yan yüz yer değiştirir.
             Simetri kodları 1 <-> 2 (12 -> 2: tek DLL'de kod 12 = yalnız X aynası); elemanların ayna düğüm
             başvurularındaki X / Y son ekleri de yer değiştirir.
  dondur180  (x, y) -> (-x, -y). Aynalar korunur; +X / -Y tercihleri ve işaretler sınanır.
  dikdortgen x 1.3 ile çarpılır: dikdörtgen gövde, ön ve yan yüz farklı genişlikte.
  eksiz      z - 3 m: zemin altına inen bacak (OutlineDrawing en düşük z'yi 0'a kaydırır).
  grupsutun  grup tablosunda "group type"tan önce fazladan bir sütun (daha yeni PLS sürümü taklidi).
             OutlineDrawing'in grup okuyucusu (tam 9 alan bekler) bu tabloyu okuyamaz; tek DLL yaması okur.
  turkce     her grup adının sonuna Windows-1254 "Çığ" eklenir (PLS dosyayı bu kod sayfasıyla yazar).
  uzunad     her 7. düğümün ve her 11. elemanın adı 15+ karaktere uzatılır (OutlineDrawing'in IsPoseLine sınırı).
  buyuk      kule 3 kez kopyalanır (x + 40 m, x + 80 m; aynalarıyla 5 kule yan yana): hız testi.
Bilerek bozuk (uyarı testi):
  eksikdugum ilk elemanın bitiş düğümü tanımsız bir adla değiştirilir.
  xarm       "X-Arm Connectivity" sayısı 2 yapılır (açı elemanı olmayan eleman).
  grupeksik  grup tablosunun 3. satırından malzeme alanı silinir (8 alan).
"""
import re
import sys


def sayi(v):
    s = "%.10g" % v
    return "0" if s in ("-0", "-0.0") else s


def main(giris, cikis, varyantlar):
    metin = open(giris, encoding="latin-1", newline="").read()
    satirlar = metin.split("\n")
    v = set(varyantlar.split(","))

    def xy(x, y):
        if "dondur90" in v:
            x, y = -y, x
        if "dondur180" in v:
            x, y = -x, -y
        if "dikdortgen" in v:
            x *= 1.3
        return x, y

    def kod(k):
        if "dondur90" in v:
            return {1: 2, 2: 1, 12: 2}.get(k, k)
        return k

    # Düğüm adları (ayna son eki çözümü için)
    etiketler = set()
    for i, s in enumerate(satirlar):
        if "; Joints Geometry" in s or "; Secondary Joints" in s:
            n = int(s.split()[0])
            for j in range(n):
                etiketler.add(satirlar[i + 1 + 5 * j].strip().strip("\r").strip("'"))

    def basvuru(r):
        """Eleman / ikincil düğüm başvurusundaki ayna son ekini 90° dönüşte çevirir (X <-> Y)."""
        if "dondur90" not in v:
            return r
        cr = r.rstrip("\r")
        son = r[len(cr):]
        q = cr.strip().strip("'")
        for ek in ("XY", "X", "Y", "P", "S"):
            if q.endswith(ek) and q[: -len(ek)] in etiketler:
                yeni = q[: -len(ek)] + {"X": "Y", "Y": "X"}.get(ek, ek)
                return cr.replace(q, yeni) + son
        return r

    TR = "\u00c7\u00fd\u00f0"   # latin-1 olarak yazılınca 1254'te "Çığ" baytları (C7 FD F0)

    def grup_adi(q):
        return q + TR if "turkce" in v else q

    i = 0
    while i < len(satirlar):
        s = satirlar[i]
        if "; Joints Geometry" in s or "; Secondary Joints" in s:
            n = int(s.split()[0])
            for j in range(n):
                k = i + 1 + 5 * j
                cr = satirlar[k + 2].endswith("\r")
                t = satirlar[k + 2].strip().split()
                x, y, z = float(t[0]), float(t[1]), float(t[2])
                x, y = xy(x, y)
                if "eksiz" in v:
                    if len(t) > 3 and abs(float(t[3]) - z) < 1e-9 and float(t[3]) != 0:
                        t[3] = sayi(float(t[3]) - 3.0)
                    z -= 3.0
                t[0], t[1], t[2] = sayi(x), sayi(y), sayi(z)
                satirlar[k + 2] = " ".join(t) + ("\r" if cr else "")
                cr = satirlar[k + 3].endswith("\r")
                t = satirlar[k + 3].strip().split()
                t[1] = str(kod(int(t[1])))
                satirlar[k + 3] = " ".join(t) + ("\r" if cr else "")
                # ikincil düğümün "from, to" başvuruları
                m = re.match(r"(\s*)'([^']*)'\s+'([^']*)'(.*)", satirlar[k + 4])
                if m and "dondur90" in v:
                    a, b = basvuru(m.group(2)), basvuru(m.group(3))
                    satirlar[k + 4] = "%s'%s' '%s'%s" % (m.group(1), a, b, m.group(4))
            i += 1 + 5 * n
            continue
        if "; Angle Member Connectivity" in s:
            n = int(s.split()[0])
            for j in range(n):
                k = i + 1 + 8 * j
                satirlar[k + 1] = basvuru(satirlar[k + 1])
                satirlar[k + 2] = basvuru(satirlar[k + 2])
                if "eksikdugum" in v and j == 0:
                    cr = satirlar[k + 2].endswith("\r")
                    satirlar[k + 2] = "YOK_DUGUM_1P" + ("\r" if cr else "")
                if "turkce" in v:
                    satirlar[k + 4] = re.sub(r"^(\s*)'([^']*)'", lambda mm: "%s'%s'" % (mm.group(1), grup_adi(mm.group(2))), satirlar[k + 4])
                cr = satirlar[k + 5].endswith("\r")
                govde = satirlar[k + 5].rstrip("\r")
                m = re.match(r"(\s*)(\d+)(.*)", govde)
                satirlar[k + 5] = m.group(1) + str(kod(int(m.group(2)))) + m.group(3) + ("\r" if cr else "")
            i += 1 + 8 * n
            continue
        if "; group label, description, size" in s and ("turkce" in v or "grupeksik" in v):
            n = int(s.split()[0])
            for k in range(i + 1, i + 1 + n):
                cr = satirlar[k].endswith("\r")
                t = re.findall(r"'[^']*'|\S+", satirlar[k].rstrip("\r"))
                if "turkce" in v:
                    t[0] = "'%s'" % grup_adi(t[0].strip("'"))
                if "grupeksik" in v and k == i + 3:
                    del t[3]
                satirlar[k] = " ".join(t) + ("\r" if cr else "")
        if "xarm" in v and re.match(r"^\s*0\s*; X-Arm Connectivity", s):
            satirlar[i] = s.replace("0", "2", 1)
        if "; group label, description, size" in s and "grupsutun" in v:
            n = int(s.split()[0])
            satirlar[i] = s.replace("element type, group type", "element type, connection class, group type")
            for k in range(i + 1, i + 1 + n):
                cr = satirlar[k].endswith("\r")
                t = re.findall(r"'[^']*'|\S+", satirlar[k].rstrip("\r"))
                t.insert(6, "0")
                satirlar[k] = " ".join(t) + ("\r" if cr else "")
            i += 1 + n
            continue
        i += 1
    if "buyuk" in v:
        satirlar = buyut(satirlar, etiketler)
    if "uzunad" in v:
        satirlar = uzat(satirlar, etiketler)
    open(cikis, "w", encoding="latin-1", newline="").write("\n".join(satirlar))


def uzat(satirlar, etiketler, ek="_UZUN_AD_TESTI"):
    """Her 7. düğümün ve 11. elemanın adına ek (15+ karakter); düğüm başvuruları da güncellenir."""
    uzun = set()
    i = 0
    while i < len(satirlar):
        s = satirlar[i]
        if "; Joints Geometry" in s or "; Secondary Joints" in s:
            n = int(s.split()[0])
            for j in range(0, n, 7):
                k = i + 1 + 5 * j
                cr = satirlar[k].endswith("\r")
                ad = satirlar[k].rstrip("\r").strip().strip("'")
                uzun.add(ad)
                satirlar[k] = "'" + ad + ek + "'" + ("\r" if cr else "")
            i += 1 + 5 * n
            continue
        i += 1

    def ref(r):
        cr = r.rstrip("\r")
        son = r[len(cr):]
        q = cr.strip().strip("'")
        for e in ("XY", "X", "Y", "P", "S"):
            if q.endswith(e) and q[: -len(e)] in etiketler:
                return (cr.replace(q, q[: -len(e)] + ek + e) if q[: -len(e)] in uzun else cr) + son
        return r
    i = 0
    while i < len(satirlar):
        s = satirlar[i]
        if "; Joints Geometry" in s or "; Secondary Joints" in s:
            n = int(s.split()[0])
            for j in range(n):
                k = i + 1 + 5 * j
                m = re.match(r"(\s*)'([^']*)'\s+'([^']*)'(.*)", satirlar[k + 4])
                if m:   # ikincil düğümün "from, to" başvuruları (OutlineDrawing konumu bunlardan çözer)
                    satirlar[k + 4] = "%s'%s' '%s'%s" % (m.group(1), ref(m.group(2)), ref(m.group(3)), m.group(4))
            i += 1 + 5 * n
            continue
        if "; Angle Member Connectivity" in s:
            n = int(s.split()[0])
            for j in range(n):
                k = i + 1 + 8 * j
                satirlar[k + 1] = ref(satirlar[k + 1])
                satirlar[k + 2] = ref(satirlar[k + 2])
                if j % 11 == 0:
                    cr = satirlar[k].endswith("\r")
                    satirlar[k] = satirlar[k].rstrip("\r").strip() + ek + ("\r" if cr else "")
            i += 1 + 8 * n
            continue
        i += 1
    return satirlar


def buyut(satirlar, etiketler, kopya=3, aralik=40.0):
    """Düğüm ve açı elemanı kayıtlarını x yönünde kopyalar (ad sonuna _k; başvurular da)."""
    def ref(r, k):
        cr = r.rstrip("\r")
        son = r[len(cr):]
        q = cr.strip().strip("'")
        for ek in ("XY", "X", "Y", "P", "S"):
            if q.endswith(ek) and q[: -len(ek)] in etiketler:
                return cr.replace(q, q[: -len(ek)] + "_%d" % k + ek) + son
        return r
    yeni = []
    i = 0
    while i < len(satirlar):
        s = satirlar[i]
        bas = ("; Joints Geometry" in s or "; Secondary Joints" in s, "; Angle Member Connectivity" in s)
        if bas[0] or bas[1]:
            n = int(s.split()[0])
            boy = 5 if bas[0] else 8
            kayit = satirlar[i + 1: i + 1 + boy * n]
            yeni.append(s.replace(str(n), str(n * kopya), 1))
            yeni.extend(kayit)
            for k in range(2, kopya + 1):
                for j in range(n):
                    r = list(kayit[boy * j: boy * j + boy])
                    cr = r[0].endswith("\r")
                    ad = r[0].rstrip("\r").strip()
                    if ad.startswith("'"):
                        r[0] = "'" + ad.strip("'") + "_%d'" % k + ("\r" if cr else "")
                    else:
                        r[0] = ad + "_%d" % k + ("\r" if cr else "")
                    if bas[0]:
                        cr2 = r[2].endswith("\r")
                        t = r[2].strip().split()
                        t[0] = sayi(float(t[0]) + aralik * (k - 1))
                        r[2] = " ".join(t) + ("\r" if cr2 else "")
                        m = re.match(r"(\s*)'([^']*)'\s+'([^']*)'(.*)", r[4])
                        if m and (m.group(2) or m.group(3)):   # ikincil düğümün "from, to" başvuruları
                            r[4] = "%s'%s' '%s'%s" % (m.group(1), ref(m.group(2), k) if m.group(2) else "",
                                                      ref(m.group(3), k) if m.group(3) else "", m.group(4))
                    else:
                        r[1] = ref(r[1], k)
                        r[2] = ref(r[2], k)
                    yeni.extend(r)
            i += 1 + boy * n
            continue
        yeni.append(s)
        i += 1
    return yeni


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], sys.argv[3])
