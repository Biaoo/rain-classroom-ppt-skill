"""跨平台插入清单的字段和原稿页码校验。"""
import re

BUTTONS = {"danmu": (1, "弹幕"), "submission": (2, "投稿"), "rollcall": (3, "点名"), "quickanswer": (4, "抢答")}
FIELDS = {
    "choice": {"options", "answer", "score", "remark"},
    "multi": {"options", "answer", "score", "halfScore", "remark"},
    "poll": {"options", "limit"},
    "blank": {"blanks", "orderInsensitive"},
    "short": {"score", "remark"},
    "judge": {"answer", "score", "remark"},
    "info": set(),
}


def strings(values):
    return isinstance(values, list) and bool(values) and all(isinstance(v, str) and v.strip() for v in values)


def integer(value, low, high):
    return type(value) is int and low <= value <= high


def blank_records(q):
    result = []
    for i, blank in enumerate(q["blanks"], 1):
        if not isinstance(blank, dict) or set(blank) - {"answers", "score", "caseSensitive", "fuzzyMatch"}:
            raise ValueError("填空配置字段不正确")
        if not strings(blank.get("answers")) or not integer(blank.get("score"), 1, 100):
            raise ValueError("每空须提供备选答案列表和 1 到 100 的整数分值")
        record = {"Num": i, "Score": blank["score"], "Answers": blank["answers"]}
        for field in ("caseSensitive", "fuzzyMatch"):
            value = blank.get(field, False)
            if type(value) is not bool:
                raise ValueError(f"{field} 须为布尔值")
            record[field[0].upper() + field[1:]] = value
        result.append(record)
    return result


def validate(plan, pages):
    if not isinstance(plan, list) or not plan:
        raise ValueError("清单须为非空数组")
    for q in plan:
        if not isinstance(q, dict) or not isinstance(q.get("kind"), str):
            raise ValueError("每项须指定 kind")
        kind = q["kind"]
        if kind in BUTTONS:
            if set(q) - {"kind", "on", "label", "box"} or not integer(q.get("on"), 1, pages):
                raise ValueError("互动入口须用 on 指定原稿页码，可选 label、box")
            if not strings([q.get("label", BUTTONS[kind][1])]):
                raise ValueError("按钮文字须为非空文本")
            box = q.get("box")
            if "box" in q:
                if not isinstance(box, list) or len(box) != 4 or not all(type(n) is int for n in box):
                    raise ValueError("box 为整数点值 [左, 上, 宽, 高]")
                x, y, w, h = box
                if min(x, y) < 0 or min(w, h) <= 0 or x+w > 960 or y+h > 540:
                    raise ValueError("互动按钮须位于页面范围内")
            continue
        if kind not in FIELDS or set(q) - ({"kind", "after", "body"} | FIELDS[kind]):
            raise ValueError(f"未知类型或字段：{kind}")
        if not integer(q.get("after"), 1, pages) or not strings([q.get("body")]):
            raise ValueError("题页须提供有效的原稿页码 after 和正文 body")
        if kind in ("choice", "multi", "poll"):
            if not strings(q.get("options")) or len(q["options"]) != 4:
                raise ValueError("选择和投票需要四个非空选项")
        if kind == "choice" and q.get("answer") not in ("A", "B", "C", "D"):
            raise ValueError("单选答案须为 A 到 D 中的一个")
        if "score" in q and not integer(q["score"], 1, 100):
            raise ValueError("分值须为 1 到 100 的整数")
        if kind == "multi":
            answers = q.get("answer")
            if not strings(answers) or len(set(answers)) != len(answers) or len(answers) < 2 or not set(answers) <= set("ABCD"):
                raise ValueError("多选答案须为至少两个不重复的选项字母")
            if not integer(q.get("halfScore", 0), 0, q.get("score", 4)):
                raise ValueError("少选得分须介于零与总分之间")
        if kind == "judge" and type(q.get("answer")) is not bool:
            raise ValueError("判断题答案须为 true 或 false")
        if "remark" in q and not isinstance(q["remark"], str):
            raise ValueError("解析须为文本")
        if kind == "poll" and not integer(q.get("limit"), 1, 4):
            raise ValueError("投票上限须为 1 到 4 的整数")
        if kind == "blank":
            if not isinstance(q.get("blanks"), list) or not q["blanks"] or type(q.get("orderInsensitive", False)) is not bool:
                raise ValueError("填空需要非空 blanks 数组，orderInsensitive 须为布尔值")
            blank_records(q)
            numbers = re.findall(r"\[填空(\d+)\]", q["body"])
            if numbers != [str(i) for i in range(1, len(q["blanks"])+1)]:
                raise ValueError("请按顺序为每空放置一个 [填空1]、[填空2] 等占位")
