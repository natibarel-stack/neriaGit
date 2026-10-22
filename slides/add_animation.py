"""Inject the right-to-left "path" animation into the slide built by build_slide.js.

Usage: python add_animation.py raw.pptx out.pptx

Sequence (shape names come from build_slide.js objectName):
  auto     title, subtitle, start marker
  click 1  dashed road wipes from the right while the marker rides it to station 1,
           circle 1 zooms in, card 1 floats up
  click 2  same to station 2
  click 3  same to station 3
  click 4  road continues to the arrow, footer fades in
"""
import re
import sys
import zipfile

SLIDE = "ppt/slides/slide1.xml"
SLIDE_W_IN = 10.0
STOPS = [9.5, 8.125, 5.205, 2.285, 0.6]  # must match build_slide.js
SPEED_MS_PER_IN = 330


class Ids:
    def __init__(self):
        self.n = 2  # 1 = root, 2 = main sequence

    def __call__(self):
        self.n += 1
        return self.n


def tgt(spid):
    return f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'


def set_visible(nid, spid):
    return (f'<p:set><p:cBhvr><p:cTn id="{nid()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
            f'{tgt(spid)}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
            f'<p:to><p:strVal val="visible"/></p:to></p:set>')


def anim(nid, spid, attr, frm, to, dur):
    return (f'<p:anim calcmode="lin" valueType="num"><p:cBhvr><p:cTn id="{nid()}" dur="{dur}" fill="hold"/>{tgt(spid)}'
            f'<p:attrNameLst><p:attrName>{attr}</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst>'
            f'<p:tav tm="0"><p:val><p:strVal val="{frm}"/></p:val></p:tav>'
            f'<p:tav tm="100000"><p:val><p:strVal val="{to}"/></p:val></p:tav></p:tavLst></p:anim>')


def effect(nid, spid, kind, node, delay=0, dur=500, dx=0.0):
    """kind: appear | fade | wipe_right | zoom | float | path"""
    preset = {"appear": ("entr", 1, 0), "fade": ("entr", 10, 0), "wipe_right": ("entr", 22, 2),
              "zoom": ("entr", 53, 16), "float": ("entr", 42, 0), "path": ("path", 0, 0)}[kind]
    cls, pid, sub = preset
    grp = ' grpId="0"' if cls == "entr" else ' grpId="1"'
    eid = nid()
    body = ""
    if cls == "entr":
        body += set_visible(nid, spid)
    if kind == "fade":
        body += f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{nid()}" dur="{dur}"/>{tgt(spid)}</p:cBhvr></p:animEffect>'
    elif kind == "wipe_right":
        body += f'<p:animEffect transition="in" filter="wipe(right)"><p:cBhvr><p:cTn id="{nid()}" dur="{dur}"/>{tgt(spid)}</p:cBhvr></p:animEffect>'
    elif kind == "zoom":
        body += anim(nid, spid, "ppt_w", "0", "#ppt_w", dur)
        body += anim(nid, spid, "ppt_h", "0", "#ppt_h", dur)
        body += f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{nid()}" dur="{dur}"/>{tgt(spid)}</p:cBhvr></p:animEffect>'
    elif kind == "float":
        body += f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{nid()}" dur="{dur}"/>{tgt(spid)}</p:cBhvr></p:animEffect>'
        body += anim(nid, spid, "ppt_x", "#ppt_x", "#ppt_x", dur)
        body += anim(nid, spid, "ppt_y", "#ppt_y+.05", "#ppt_y", dur)
    elif kind == "path":
        body += (f'<p:animMotion origin="layout" path="M 0 0 L {dx:.5f} 0 E" pathEditMode="relative" ptsTypes="AA">'
                 f'<p:cBhvr><p:cTn id="{nid()}" dur="{dur}" fill="hold"/>{tgt(spid)}'
                 f'<p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr>'
                 f'<p:rCtr x="{int(dx * 50000)}" y="0"/></p:animMotion>')
    return (f'<p:par><p:cTn id="{eid}" presetID="{pid}" presetClass="{cls}" presetSubtype="{sub}" fill="hold"{grp} nodeType="{node}">'
            f'<p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst>{body}</p:childTnLst></p:cTn></p:par>')


def step(nid, delay, effects_fn):
    sid = nid()
    inner = effects_fn()
    return (f'<p:par><p:cTn id="{sid}" fill="hold"><p:stCondLst><p:cond delay="{delay}"/></p:stCondLst>'
            f'<p:childTnLst>{inner}</p:childTnLst></p:cTn></p:par>')


def group(nid, steps, auto=False):
    gid = nid()
    cond = '<p:cond delay="indefinite"/>'
    if auto:
        cond += '<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>'
    return (f'<p:par><p:cTn id="{gid}" fill="hold"><p:stCondLst>{cond}</p:stCondLst>'
            f'<p:childTnLst>{"".join(steps)}</p:childTnLst></p:cTn></p:par>')


def build_timing(sp):
    nid = Ids()
    groups = []

    # Auto intro
    groups.append(group(nid, [
        step(nid, 0, lambda: effect(nid, sp["Title"], "fade", "afterEffect", dur=600)
                             + effect(nid, sp["Subtitle"], "fade", "withEffect", delay=250, dur=600)),
        step(nid, 850, lambda: effect(nid, sp["Marker1"], "zoom", "afterEffect", dur=400)),
    ], auto=True))

    for i in range(4):
        right, left = STOPS[i], STOPS[i + 1]
        dur = max(500, int((right - left) * SPEED_MS_PER_IN))
        dx = -(right - left) / SLIDE_W_IN
        seg, marker = sp[f"Seg{i + 1}"], sp[f"Marker{i + 1}"]

        def travel(seg=seg, marker=marker, dur=dur, dx=dx, i=i):
            out = effect(nid, seg, "wipe_right", "clickEffect", dur=dur)
            if i > 0:  # marker 1 is already visible from the intro
                out += effect(nid, marker, "appear", "withEffect")
            out += effect(nid, marker, "path", "withEffect", dur=dur, dx=dx)
            return out

        steps = [step(nid, 0, travel)]
        if i < 3:
            circle, card = sp[f"Circle{i + 1}"], sp[f"Card{i + 1}"]
            steps.append(step(nid, dur, lambda c=circle, k=card:
                              effect(nid, c, "zoom", "afterEffect", dur=400)
                              + effect(nid, k, "float", "withEffect", delay=200, dur=600)))
        else:
            steps.append(step(nid, dur, lambda: effect(nid, sp["Footer"], "fade", "afterEffect", dur=600)))
        groups.append(group(nid, steps))

    main = (f'<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq">'
            f'<p:childTnLst>{"".join(groups)}</p:childTnLst></p:cTn>'
            f'<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
            f'<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>')

    text_shapes = ["Title", "Subtitle", "Card1", "Card2", "Card3", "Circle1", "Circle2", "Circle3", "Footer"]
    bld = "".join(f'<p:bldP spid="{sp[n]}" grpId="0" animBg="1"/>' for n in text_shapes)
    return (f'<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">'
            f'<p:childTnLst>{main}</p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst>{bld}</p:bldLst></p:timing>')


def main(src, dst):
    zin = zipfile.ZipFile(src)
    xml = zin.read(SLIDE).decode("utf8")
    sp = {name: sid for sid, name in re.findall(r'<p:cNvPr id="(\d+)" name="([^"]+)"', xml)}

    # pptxgenjs drops rtl on multi-run text; every right-aligned paragraph here is Hebrew
    xml = re.sub(r'<a:pPr (?![^>]*rtl=)([^>]*algn="r")', r'<a:pPr rtl="1" \1', xml)
    xml = xml.replace('rtl="1"  algn', 'rtl="1" algn')

    xml = xml.replace("</p:clrMapOvr>", "</p:clrMapOvr>" + build_timing(sp), 1)

    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = xml.encode("utf8") if item.filename == SLIDE else zin.read(item.filename)
            zout.writestr(item, data)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
