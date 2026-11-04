"""在已有宽屏 PPTX 的指定原稿页后插入四选项单选题或匿名投票。"""
import argparse
import json
import posixpath
import uuid
from pathlib import Path
from xml.dom import minidom as dom
from xml.etree import ElementTree as ET
from zipfile import ZipFile, ZIP_DEFLATED

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"p": P, "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
CONTENT = "application/vnd.openxmlformats-officedocument.presentationml."
TEMPLATES = Path(__file__).resolve().parents[1] / "assets" / "native-pages.zip"


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
    slide_tags = tags["rId1"]
    if kind == "choice":
        remark = q.get("remark", "")
        slide_tags["PROBLEMHASREMARK"] = str(bool(remark))
        slide_tags["PROBLEMREMARK"] = remark
    else:
        slide_tags["ANONYMOUSPOLLING"] = "True"
    for node in template.iter():
        if node.tag.endswith("}creationId"):
            if "id" in node.attrib:
                node.set("id", "{" + str(uuid.uuid4()).upper() + "}")
            elif "val" in node.attrib:
                node.set("val", str(uuid.uuid4().int & 0xffffffff))
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
        elif role == "ProblemBullet" and kind == "poll":
            geom = shape.find("p:spPr/a:prstGeom", NS)
            geom.set("prst", "ellipse" if q["limit"] == 1 else "rect")
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
    if not isinstance(plan, list) or not plan:
        raise ValueError("插题清单须为非空数组")
    for i, q in enumerate(plan, 1):
        if not isinstance(q, dict) or q.get("kind") not in ("choice", "poll"):
            raise ValueError(f"第 {i} 项须指定 choice 或 poll")
        if type(q.get("after")) is not int or not 1 <= q["after"] <= len(originals):
            raise ValueError(f"第 {i} 项的 after 须为有效的原稿页码")
        options = q.get("options")
        if not isinstance(options, list) or len(options) != 4 or not all(
            isinstance(x, str) and x.strip() for x in [q.get("body"), *options]
        ):
            raise ValueError(f"第 {i} 项需要题干和四个非空选项")
        if q["kind"] == "choice":
            if q.get("answer") not in ("A", "B", "C", "D") or not isinstance(q.get("remark", ""), str):
                raise ValueError(f"第 {i} 项需要 A 到 D 的单个答案，解析须为文本")
        elif type(q.get("limit")) is not int or not 1 <= q["limit"] <= 4:
            raise ValueError(f"第 {i} 项的投票上限须为 1 到 4 的整数")
        allowed = {"kind", "after", "body", "options"} | ({"answer", "remark"} if q["kind"] == "choice" else {"limit"})
        if set(q) - allowed:
            raise ValueError(f"第 {i} 项包含未支持的字段：{set(q) - allowed}")
    with ZipFile(TEMPLATES) as seeds:
        image = allocate(parts, "ppt/media", "image", "png")
        parts[image] = seeds.read("marker.png")
        if not any(n.getAttribute("Extension") == "png" for n in types.getElementsByTagName("Default")):
            append(types, types.documentElement, "Override", PartName="/"+image, ContentType="image/png")
        # 逆序插入确保同一锚点上的题目仍按清单顺序出现。
        for i, q in reversed(list(enumerate(plan, 1))):
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
