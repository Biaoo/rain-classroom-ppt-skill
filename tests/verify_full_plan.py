"""核对完整清单生成的七页及原稿上的四个互动入口。"""
import json
import sys
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile
from posixpath import dirname, join, normpath

root = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(root / "scripts"))
from inspect_deck import inspect

base, output = map(Path, sys.argv[1:])
original = inspect(base)["slides"]
slides = inspect(output)["slides"]
plan = json.loads((root / "references/full-plan.json").read_text(encoding="utf-8"))
assert len(slides) == 10
assert [slides[0], slides[4]] == original[:2]
assert slides[9]["tags"] == original[2]["tags"]
assert slides[9]["shapes"][:-4] == original[2]["shapes"]
buttons = slides[9]["shapes"][-4:]
assert [b["text"] for b in buttons] == [q["label"] for q in plan[7:]]
assert [b["tags"]["INTERACTIONTYPE"] for b in buttons] == list("1234")
assert {b["tags"]["INTERACTIONBUTTONSTYLE"] for b in buttons} == {"0"}
assert len({b["tags"]["LOCALSHAPEID"] for b in buttons}) == 4
types = ["MultipleChoice", "MultipleChoiceMA", "Polling", "FillBlank", "ShortAnswer", "MultipleChoice", None]
for q, index, kind in zip(plan, [1, 2, 3, 5, 6, 7, 8], types):
    s = slides[index]
    assert s["tags"].get("RAINPROBLEM") == kind
    assert [x["text"] for x in s["shapes"] if x["tags"].get("RAINPROBLEM") == "ProblemBody"] == [q["body"]]
    if "options" in q:
        assert [x["text"] for x in s["shapes"] if x["tags"].get("RAINPROBLEM") == "ProblemItem"] == q["options"]
    if "answer" in q:
        expected = {"choice": ["B"], "multi": ["A", "C"], "judge": ["B"]}[q["kind"]]
        assert [x["text"] for x in s["shapes"] if x["tags"].get("RAINBULLET") == "Correct"] == expected
    if "score" in q:
        assert s["tags"]["PROBLEMSCORE"] == str(q["score"])
        assert str(q["score"])+"分" in [x["text"] for x in s["shapes"]]
assert slides[2]["tags"]["PROBLEMSCORE_HALF"] == "2"
assert slides[3]["tags"]["ANONYMOUSPOLLING"] == "True"
assert any(x["text"] == "最多可选2项" for x in slides[3]["shapes"])
blanks = json.loads(slides[5]["tags"]["PROBLEMBLANK"])
assert [(b["Num"], b["Score"], b["Answers"]) for b in blanks] == [(1, 2, ["氧气", "氧"]), (2, 3, ["重力"])]
assert slides[5]["tags"]["PROBLEMSCORE"] == "5"
assert slides[6]["tags"]["PROBLEMREMARK"] == plan[4]["remark"]
assert [x["text"] for x in slides[7]["shapes"] if x["tags"].get("RAINPROBLEM") == "ProblemItem"] == ["对", "错"]
with ZipFile(output) as z:
    for name in z.namelist():
        if name.endswith(".rels"):
            rels = ET.fromstring(z.read(name))
            assert len({r.get("Id") for r in rels}) == len(rels)
            for r in rels:
                if r.get("TargetMode") != "External":
                    assert normpath(join(dirname(dirname(name)), r.get("Target"))).lstrip("/") in z.namelist()
print("PASS: seven native page formats and four original-slide interactions")
