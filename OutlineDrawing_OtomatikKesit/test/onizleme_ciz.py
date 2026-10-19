import sys, matplotlib; matplotlib.use("Agg"); import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle
src, out, baslik = sys.argv[1], sys.argv[2], sys.argv[3]
fig = plt.figure(figsize=(841/25.4*0.75, 594.65/25.4*0.75)); ax = fig.add_axes([0,0,1,1]); ax.set_aspect('equal'); ax.axis('off')
ax.set_xlim(-5, 846); ax.set_ylim(-5, 600)
ax.add_patch(Rectangle((0,0),841,594.65,fill=False,lw=1.2))
olcek = ""
for L in open(src):
    p = L.split()
    if not p: continue
    t = p[0]
    if t in ("L","H","M","D"):
        a,b,c,d = map(float, p[1:5])
        if t == "L": ax.plot([a,c],[b,d],'-',color='#0a7d8c',lw=0.35)
        elif t == "H": ax.plot([a,c],[b,d],'--',color='#3050d0',lw=0.25)
        elif t == "M": ax.plot([a,c],[b,d],'-',color='r',lw=0.4)
        else: ax.add_patch(Rectangle((a,b),c-a,d-b,fill=False,color='#999999',lw=0.2))
    elif t in ("R","Q"):
        a,b,c,d = map(float, p[1:5]); ax.add_patch(Rectangle((a,b),c-a,d-b,fill=(t=="Q"),alpha=0.08 if t=="Q" else 1,color='k',lw=0.6))
    elif t == "T":
        x,y,h = map(float,p[1:4]); ax.text(x,y,' '.join(p[4:]),ha='center',va='bottom',fontsize=max(3, h*2.0))
    elif t == "S": olcek = p[1]
ax.text(700, 150, baslik + "\nÖlçek " + olcek, fontsize=9)
fig.savefig(out, dpi=110)
