#!/usr/bin/env python3
"""PLS-TOWER .tow -> eleman CSV (AutoCAD'siz test icin).

OutlineDrawing'in dataprocessing akisini taklit eder:
  - Primary joint etiketi + 'P', secondary joint etiketi + 'S';
    simetri kodu 1 -> 'X' (y -> -y), 2 -> 'Y' (x -> -x), 3 -> X, Y, XY
  - Eleman simetri kodu 1/2/3 ayni kurala gore cogaltilir (diger kodlar tek eleman)
  - Koordinatlar m -> mm, RotateAllMinus90AboutZ: (x, y) -> (y, -x), en dusuk Z = 0
Cikti satiri: anahtar,sx,sy,sz,ex,ey,ez,grup,boyut,section_label

Kullanim:  python3 tow_to_csv.py kule.tow > kule.csv
"""
import csv
import re
import sys


def tokens(s):
    return re.findall(r"'[^']*'|\S+", s)


def unq(s):
    return s.strip().strip("'")


def main(path):
    lines = open(path, encoding="latin-1").read().split("\n")
    lines = [l.rstrip("\r") for l in lines]
    joints, members, groups = {}, [], {}

    def add_joint(label, x, y, z, sym, own):
        joints[label + own] = (x, y, z)
        if sym in (1, 3):
            joints[label + "X"] = (x, -y, z)
        if sym in (2, 3):
            joints[label + "Y"] = (-x, y, z)
        if sym == 3:
            joints[label + "XY"] = (-x, -y, z)

    i = 0
    while i < len(lines):
        line = lines[i]
        if "; Joints Geometry" in line or "; Secondary Joints" in line:
            own = "S" if "Secondary" in line else "P"
            k = i + 1
            for _ in range(int(line.split()[0])):
                xyz = lines[k + 2].split()
                add_joint(unq(lines[k]), float(xyz[0]), float(xyz[1]), float(xyz[2]),
                          int(lines[k + 3].split()[1]), own)
                k += 5
            i = k
            continue
        if "; Angle Member Connectivity" in line:
            k = i + 1
            for _ in range(int(line.split()[0])):
                gs = tokens(lines[k + 4])
                members.append((lines[k].strip(), lines[k + 1].strip(), lines[k + 2].strip(),
                                unq(gs[0]), unq(gs[1]) if len(gs) > 1 else "",
                                int(lines[k + 5].split()[0])))
                k += 8
            i = k
            continue
        if "; group label, description, size" in line:
            n = int(line.split()[0])
            for k in range(i + 1, i + 1 + n):
                t = tokens(lines[k])
                groups[unq(t[0])] = unq(t[2]).strip()
            i += n + 1
            continue
        i += 1

    rows, missing = [], 0
    for label, a, b, grp, sec, sym in members:
        if a not in joints or b not in joints:
            missing += 1
            continue
        A, B = joints[a], joints[b]
        variants = [("", 1, 1)]
        if sym in (1, 3):
            variants.append(("X", 1, -1))
        if sym in (2, 3):
            variants.append(("Y", -1, 1))
        if sym == 3:
            variants.append(("XY", -1, -1))
        for suf, fx, fy in variants:
            rows.append((label + suf, A[0] * fx, A[1] * fy, A[2], B[0] * fx, B[1] * fy, B[2], grp, groups.get(grp, ""), sec))

    zmin = min(min(r[3], r[6]) for r in rows)
    w = csv.writer(sys.stdout, lineterminator="\n")
    for r in rows:
        w.writerow([r[0],
                    round(r[2] * 1000, 3), round(-r[1] * 1000, 3), round((r[3] - zmin) * 1000, 3),
                    round(r[5] * 1000, 3), round(-r[4] * 1000, 3), round((r[6] - zmin) * 1000, 3),
                    r[7], r[8], r[9]])
    print("dugum={0} eleman={1} uretilen={2} bulunamayan={3}".format(len(joints), len(members), len(rows), missing),
          file=sys.stderr)


if __name__ == "__main__":
    main(sys.argv[1])
