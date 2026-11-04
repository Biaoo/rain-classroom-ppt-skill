"""在已有宽屏 PPTX 的指定原稿页后各插入一道四选项单选题和匿名投票。"""
import argparse
import json
import posixpath
from pathlib import Path
from xml.dom import minidom as dom
from xml.etree import ElementTree as ET
from zipfile import ZipFile, ZIP_DEFLATED

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"p": P, "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
CONTENT = "application/vnd.openxmlformats-officedocument.presentationml."
FIXTURES = Path(__file__).parent / "fixtures"


def xml(element):
    return ET.tostring(element, encoding="utf-8", xml_declaration=True)


def append(doc, parent, tag, **attrs):
    node = doc.createElementNS(parent.namespaceURI, tag)
    for key, value in attrs.items():
        node.setAttribute(key, str(value))
    parent.appendChild(node)
    return node


def allocate(parts, folder, stem, extension):
    i = 1
    while f"{folder}/{stem}{i}.{extension}" in parts:
        i += 1
    return f"{folder}/{stem}{i}.{extension}"


def customize(template, tags, q):
    kind = q["kind"]
    labels = q["options"]
    if len(labels) != 4 or not all(isinstance(x, str) and x.strip() for x in [q["body"], *labels]):
        raise ValueError("每道题需要题干和四个非空选项")
    slide_tags = tags["rId1"]
    if kind == "choice":
        if q["answer"] not in list("ABCD"):
            raise ValueError("单选答案须为 A、B、C 或 D")
        slide_tags["PROBLEMREMARK"] = q["remark"]
    else:
        if not 1 <= q["limit"] <= 4:
            raise ValueError("投票上限须为 1 到 4")
        slide_tags["ANONYMOUSPOLLING"] = "True"
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
        elif role == "ProblemBullet" and kind == "choice":
            correct = original == q["answer"]
            meta["RAINBULLET"] = "Correct" if correct else "Wrong"
            shape.find("p:spPr/a:solidFill/a:srgbClr", NS).set("val", "00FF00" if correct else "808080")
        elif role == "PollingAnswer":
            replacement = f"最多可选{q['limit']}项"
        elif kind == "poll" and original == "投票":
            replacement = "投票(匿名)"
        if replacement is not None:
            runs[0].text = replacement
            for run in runs[1:]:
                run.text = ""


def insert(base, plan, output):
    with ZipFile(base) as source:
        parts = {name: source.read(name) for name in source.namelist()}
    presentation = dom.parseString(parts["ppt/presentation.xml"])
    pres_rels = dom.parseString(parts["ppt/_rels/presentation.xml.rels"])
    types = dom.parseString(parts["[Content_Types].xml"])
    size = presentation.getElementsByTagNameNS(P, "sldSz")[0]
    if (size.getAttribute("cx"), size.getAttribute("cy")) != ("12192000", "6858000"):
        raise ValueError("原型使用 13⅓ × 7½ 英寸的原生题页")
    order = presentation.getElementsByTagNameNS(P, "sldIdLst")[0]
    originals = list(order.getElementsByTagNameNS(P, "sldId"))
    mapping = {r.getAttribute("Id"): r.getAttribute("Target") for r in pres_rels.documentElement.childNodes if r.nodeType == r.ELEMENT_NODE}
    slide_id = max(int(s.getAttribute("id")) for s in originals)
    if sorted(q["kind"] for q in plan) != ["choice", "poll"]:
        raise ValueError("原型需要一道单选和一道匿名投票")
    with ZipFile(FIXTURES / "native-pages.zip") as seeds:
        image = allocate(parts, "ppt/media", "image", "png")
        parts[image] = seeds.read("marker.png")
        if not any(n.getAttribute("Extension") == "png" for n in types.getElementsByTagName("Default")):
            append(types, types.documentElement, "Override", PartName="/"+image, ContentType="image/png")
        # 逆序插入确保同一锚点上的题目仍按清单顺序出现。
        for i, q in reversed(list(enumerate(plan, 1))):
            if not 1 <= q["after"] <= len(originals):
                raise ValueError("原稿页码超出范围")
            anchor = originals[q["after"] - 1]
            target = mapping[anchor.getAttributeNS(R, "id")]
            original_part = posixpath.normpath(posixpath.join("ppt", target)).lstrip("/")
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
            relate("rId19", "image", image)
            for rid, fields in tags.items():
                tag_part = allocate(parts, "ppt/tags", "tag", "xml")
                root = ET.Element(f"{{{P}}}tagLst")
                for name, val in fields.items():
                    ET.SubElement(root, f"{{{P}}}tag", name=name, val=val)
                parts[tag_part] = xml(root)
                append(types, types.documentElement, "Override", PartName="/"+tag_part, ContentType=CONTENT+"tags+xml")
                relate(rid, "tags", tag_part)
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
    with output.open("xb") as stream, ZipFile(stream, "w", ZIP_DEFLATED) as out:
        for name, data in parts.items():
            out.writestr(name, data)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    for arg in ("base", "plan", "output"):
        parser.add_argument(arg, type=Path)
    args = parser.parse_args()
    insert(args.base, json.loads(args.plan.read_text(encoding="utf-8")), args.output)
