"""在已有宽屏 PPTX 的指定原稿页后插入雨课堂原生题页、信息页及页内互动。"""
import argparse
import json
import posixpath
import uuid
from pathlib import Path
from xml.dom import minidom as dom
from xml.etree import ElementTree as ET
from zipfile import ZipFile, ZIP_DEFLATED

from insertion_plan import BUTTONS, blank_records, validate

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"p": P, "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
CONTENT = "application/vnd.openxmlformats-officedocument.presentationml."
TEMPLATES = Path(__file__).resolve().parents[1] / "assets" / "native-pages.zip"


def xml(element):
    return ET.tostring(element, encoding="utf-8", xml_declaration=True)


def append(doc, parent, tag, **attrs):
    qualified = f"{parent.prefix}:{tag}" if parent.prefix else tag
    node = doc.createElementNS(parent.namespaceURI, qualified)
    for key, value in attrs.items():
        node.setAttribute(key, str(value))
    parent.appendChild(node)
    return node


def allocate(parts, folder, stem, extension):
    i = 1
    while f"{folder}/{stem}{i}.{extension}" in parts:
        i += 1
    return f"{folder}/{stem}{i}.{extension}"


def refresh_ids(template):
    for node in template.iter():
        if node.tag.endswith("}creationId"):
            if "id" in node.attrib:
                node.set("id", "{" + str(uuid.uuid4()).upper() + "}")
            elif "val" in node.attrib:
                node.set("val", str(uuid.uuid4().int & 0xffffffff))


def add_tags(parts, types, fields):
    part = allocate(parts, "ppt/tags", "tag", "xml")
    root = ET.Element(f"{{{P}}}tagLst")
    for name, val in fields.items():
        ET.SubElement(root, f"{{{P}}}tag", name=name, val=val)
    parts[part] = xml(root)
    append(types, types.documentElement, "Override", PartName="/"+part, ContentType=CONTENT+"tags+xml")
    return part


def add_button(parts, types, seeds, part, q):
    # DOM 保留原页的命名空间前缀，包括已有扩展使用的 QName。
    doc = dom.parseString(parts[part])
    folder, name = posixpath.split(part)
    relpath = f"{folder}/_rels/{name}.rels"
    rels = dom.parseString(parts[relpath])
    rid = "rIdRainButton"
    used = {r.getAttribute("Id") for r in rels.getElementsByTagNameNS(PKG, "Relationship")}
    while rid in used:
        rid += "x"
    shape = ET.fromstring(seeds.read("button.xml"))
    refresh_ids(shape)
    index, label = BUTTONS[q["kind"]]
    button_id = max(int(n.getAttribute("id")) for n in doc.getElementsByTagNameNS(P, "cNvPr")) + 1
    shape.find("p:nvSpPr/p:cNvPr", NS).attrib.update(id=str(button_id), name=f"{label} {button_id}")
    shape.find(".//p:tags", NS).set(f"{{{R}}}id", rid)
    shape.find(".//a:t", NS).text = q.get("label", label)
    shape.find(".//a:rPr/a:solidFill/a:srgbClr", NS).set("val", "222222")
    x, y, w, h = q.get("box", [24 + (index-1)*234, 470, 208, 46])
    trans = shape.find("p:spPr/a:xfrm", NS)
    trans.find("a:off", NS).attrib.update(x=str(x*12700), y=str(y*12700))
    trans.find("a:ext", NS).attrib.update(cx=str(w*12700), cy=str(h*12700))
    tree = doc.getElementsByTagNameNS(P, "spTree")[0]
    tail = next((n for n in tree.childNodes if n.localName == "extLst"), None)
    tree.insertBefore(doc.importNode(dom.parseString(xml(shape)).documentElement, True), tail)
    tag = add_tags(parts, types, {"INTERACTIONTYPE": str(index), "INTERACTIONBUTTONSTYLE": "0",
                                 "LOCALSHAPEID": str(10**16 + uuid.uuid4().int % (9*10**16))})
    append(rels, rels.documentElement, "Relationship", Id=rid, Type=R+"/tags", Target=posixpath.relpath(tag, folder))
    parts[part] = doc.toxml(encoding="utf-8")
    parts[relpath] = rels.toxml(encoding="utf-8")


def customize(template, tags, q):
    kind = q["kind"]
    slide_tags = tags["rId1"]
    labels = q.get("options", ["对", "错"] if kind == "judge" else [])
    correct_letters = q.get("answer", [])
    if kind == "judge":
        correct_letters = "A" if q["answer"] else "B"
    if kind in ("choice", "multi", "short", "judge"):
        remark = q.get("remark", "")
        slide_tags.update(PROBLEMHASREMARK=str(bool(remark)), PROBLEMREMARK=remark)
        if "score" in q:
            slide_tags["PROBLEMSCORE"] = str(q["score"])
    if kind == "multi":
        slide_tags["PROBLEMSCORE_HALF"] = str(q.get("halfScore", 0))
    if kind == "blank":
        records = blank_records(q)
        slide_tags.update(PROBLEMBLANK=json.dumps(records, ensure_ascii=True, separators=(",", ":")),
                          PROBLEMBLANKORDER=str(q.get("orderInsensitive", False)),
                          PROBLEMSCORE=str(sum(r["Score"] for r in records)))
    if kind == "poll":
        slide_tags["ANONYMOUSPOLLING"] = "True"
    refresh_ids(template)
    items = iter(labels)
    for shape in template.findall(".//p:sp", NS):
        ref = shape.find("p:nvSpPr/p:nvPr/p:custDataLst/p:tags", NS)
        if ref is None:
            continue
        meta = tags[ref.get(f"{{{R}}}id")]
        role = meta.get("RAINPROBLEM")
        runs = shape.findall(".//a:t", NS)
        original = "".join(n.text or "" for n in runs)
        replacement = None
        if role == "ProblemBody":
            replacement = q["body"]
        elif role == "ProblemItem":
            replacement = next(items)
        elif role == "ProblemBullet" and kind in ("choice", "multi", "judge"):
            correct = original in correct_letters
            meta["RAINBULLET"] = "Correct" if correct else "Wrong"
            shape.find("p:spPr/a:solidFill/a:srgbClr", NS).set("val", "00FF00" if correct else "808080")
        elif role == "ProblemBullet" and kind == "poll":
            geom = shape.find("p:spPr/a:prstGeom", NS)
            geom.set("prst", "ellipse" if q["limit"] == 1 else "rect")
        elif meta.get("RAINPROBLEMTYPE") == "ProblemTypeMarker" and original.endswith("分"):
            replacement = slide_tags["PROBLEMSCORE"] + "分"
        elif role == "PollingAnswer":
            replacement = f"最多可选{q['limit']}项"
        elif kind == "poll" and original == "投票":
            replacement = "投票(匿名)"
        if replacement is not None:
            runs[0].text = replacement
            for run in runs[1:]:
                run.text = ""


def insert(base, plan, output):
    if base.suffix.lower() != ".pptx" or output.suffix.lower() != ".pptx":
        raise ValueError("输入和输出均须为 .pptx 文件")
    if output.exists():
        raise FileExistsError(f"请指定新的输出文件：{output}")
    with ZipFile(base) as source:
        parts = {name: source.read(name) for name in source.namelist()}
    presentation = dom.parseString(parts["ppt/presentation.xml"])
    pres_rels = dom.parseString(parts["ppt/_rels/presentation.xml.rels"])
    types = dom.parseString(parts["[Content_Types].xml"])
    size = presentation.getElementsByTagNameNS(P, "sldSz")[0]
    if (size.getAttribute("cx"), size.getAttribute("cy")) != ("12192000", "6858000"):
        raise ValueError("题页模板使用 13⅓ × 7½ 英寸的原生题页")
    order = presentation.getElementsByTagNameNS(P, "sldIdLst")[0]
    originals = list(order.getElementsByTagNameNS(P, "sldId"))
    mapping = {r.getAttribute("Id"): r.getAttribute("Target") for r in pres_rels.documentElement.childNodes if r.nodeType == r.ELEMENT_NODE}
    slide_id = max(int(s.getAttribute("id")) for s in originals)
    validate(plan, len(originals))
    def original_path(anchor):
        target = mapping[anchor.getAttributeNS(R, "id")]
        return posixpath.normpath(posixpath.join("ppt", target)).lstrip("/")
    with ZipFile(TEMPLATES) as seeds:
        pages = []
        for q in plan:
            if q["kind"] in BUTTONS:
                add_button(parts, types, seeds, original_path(originals[q["on"]-1]), q)
            else:
                pages.append(q)
        if any(q["kind"] != "info" for q in pages):
            image = allocate(parts, "ppt/media", "image", "png")
            parts[image] = seeds.read("marker.png")
            if not any(n.getAttribute("Extension") == "png" for n in types.getElementsByTagNameNS("*", "Default")):
                append(types, types.documentElement, "Override", PartName="/"+image, ContentType="image/png")
        # 逆序插入确保同一锚点上的题目仍按清单顺序出现。
        for i, q in reversed(list(enumerate(pages, 1))):
            anchor = originals[q["after"] - 1]
            original_part = original_path(anchor)
            folder, basename = posixpath.split(original_part)
            anchor_rels = ET.fromstring(parts[f"{folder}/_rels/{basename}.rels"])
            layout = next(r.get("Target") for r in anchor_rels if r.get("Type") == R+"/slideLayout")
            layout_part = posixpath.normpath(posixpath.join(folder, layout)).lstrip("/")
            kind = q["kind"]
            template = ET.fromstring(seeds.read(kind+".xml"))
            tags = json.loads(seeds.read(kind+".json"))
            customize(template, tags, q)
            part = allocate(parts, "ppt/slides", "slide", "xml")
            parts[part] = xml(template)
            rels = ET.Element("Relationships", xmlns=PKG)
            def relate(rid, role, dest):
                ET.SubElement(rels, "Relationship", Id=rid, Type=R+"/"+role, Target=posixpath.relpath(dest, "ppt/slides"))
            relate("rId18", "slideLayout", layout_part)
            if kind != "info":
                relate("rId19", "image", image)
            for rid, fields in tags.items():
                relate(rid, "tags", add_tags(parts, types, fields))
            parts["ppt/slides/_rels/"+posixpath.basename(part)+".rels"] = xml(rels)
            append(types, types.documentElement, "Override", PartName="/"+part, ContentType=CONTENT+"slide+xml")
            rid = f"rIdRain{i}"
            while rid in mapping:
                rid += "x"
            mapping[rid] = posixpath.relpath(part, "ppt")
            append(pres_rels, pres_rels.documentElement, "Relationship", Id=rid, Type=R+"/slide", Target=mapping[rid])
            new = anchor.cloneNode(False)
            new.setAttribute("id", str(slide_id+i))
            new.setAttributeNS(R, anchor.getAttributeNodeNS(R,"id").name, rid)
            order.insertBefore(new, anchor.nextSibling)
    for path, doc in [("ppt/presentation.xml", presentation), ("ppt/_rels/presentation.xml.rels", pres_rels), ("[Content_Types].xml", types)]:
        parts[path] = doc.toxml(encoding="utf-8")
    with output.open("xb") as stream:
        try:
            with ZipFile(stream, "w", ZIP_DEFLATED) as out:
                for name, data in parts.items():
                    out.writestr(name, data)
        except BaseException:
            stream.close()
            output.unlink()
            raise


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    for arg in ("base", "plan", "output"):
        parser.add_argument(arg, type=Path)
    args = parser.parse_args()
    insert(args.base, json.loads(args.plan.read_text(encoding="utf-8")), args.output)
