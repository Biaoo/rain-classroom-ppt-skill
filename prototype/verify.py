"""核对 Windows 生成的交付稿、原稿页序和引用。"""
import hashlib
import posixpath
import sys
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from inspect_deck import inspect

base, result = map(Path, sys.argv[1:])
originals = inspect(base)["slides"]
report = inspect(result)
slides = report["slides"]
assert len(slides) == 5 and slides[::2] == originals
assert not report["has_embedded_vba"]
choice, poll = slides[1], slides[3]
assert choice["tags"]["RAINPROBLEM"] == "MultipleChoice"
assert choice["tags"]["PROBLEMSCORE"] == "2"
assert poll["tags"]["RAINPROBLEM"] == "Polling"
assert poll["tags"]["ANONYMOUSPOLLING"] == "True"


def values(slide, field, expected):
    return [s["text"] for s in slide["shapes"] if s["tags"].get(field) == expected]


assert values(choice, "RAINBULLET", "Correct") == ["B"]
assert values(choice, "RAINPROBLEM", "ProblemBody") == ["活性污泥法主要依靠什么去除有机物？"]
assert values(choice, "RAINPROBLEM", "ProblemItem") == ["格栅截留","微生物降解","重力沉淀","紫外消毒"]
assert values(poll, "RAINPROBLEM", "ProblemBody") == ["你最想深入了解哪些污水处理环节？"]
assert values(poll, "RAINPROBLEM", "ProblemItem") == ["预处理","生物处理","深度处理","污泥处理"]
assert values(poll, "RAINPROBLEM", "PollingAnswer") == ["最多可选2项"]

with ZipFile(base) as source, ZipFile(result) as output:
    for path in source.namelist():
        if path.startswith("ppt/slides/"):
            assert output.read(path) == source.read(path), path
    for path in output.namelist():
        if path.endswith(".rels"):
            parent = posixpath.dirname(posixpath.dirname(path))
            for rel in ET.fromstring(output.read(path)):
                if rel.get("TargetMode") != "External":
                    target = posixpath.normpath(posixpath.join(parent, rel.get("Target"))).lstrip("/")
                    assert target in output.namelist(), (path, target)
print("PASS: original pages, slide order, questions, options, answer B, anonymous poll, limit 2, package relationships")
print("SHA256", hashlib.sha256(result.read_bytes()).hexdigest())
