#!/usr/bin/env python3
"""调用本机雨课堂插件，从 JSON 数组新建包含原生题目的 PPTX。
用法：python3 create_deck.py questions.json new.pptx
"""
import argparse
import json
import subprocess
import sys
import uuid
from pathlib import Path
from inspect_deck import inspect

def apple_string(value):
    return '"' + value.replace('\\', '\\\\').replace('"', '\\"').replace('\n', '\\n') + '"'

def create(path, questions):
    if sys.platform != 'darwin':
        raise RuntimeError('此创建脚本使用 macOS PowerPoint AppleScript 接口')
    if path.exists():
        raise FileExistsError('目标文件已存在，请指定新文件路径：' + str(path))
    if not path.parent.is_dir() or path.suffix.lower() != '.pptx':
        raise ValueError('Use an existing directory and a .pptx filename')
    out = apple_string(str(path))
    name = apple_string(path.name)
    lines = ['with timeout of 90 seconds', 'tell application "Microsoft PowerPoint"',
             'if not (loaded of add in "RainClassroom.ppam") then error "RainClassroom is not loaded"',
             'set p to make new presentation',
             'tell p to make new slide at end with properties {layout:slide layout blank}',
             'if (count of slides of p) is not 1 then error "Seed creation failed"',
             f'save p in POSIX file {out}',
             f'set p to presentation {name}']
    for i, q in enumerate(questions):
        payload = apple_string(json.dumps(q, ensure_ascii=False))
        op = apple_string('rainclassroom-' + uuid.uuid4().hex)
        lines += [f'if full name of active presentation is not {out} then error "Active document changed"',
                  f'if (count of slides of p) is not {i + 1} then error "Unexpected pre-insert count"',
                  f'run VB macro macro name "RainClassroom.ppam!Message.InsertProb" list of parameters {{{op}, {payload}, "0"}}',
                  f'set p to presentation {name}',
                  f'if (count of slides of p) is not {i + 2} then error "Insert failed; inspect before retrying"']
        if q['ProblemType'] == 0:
            lines += [f'set bodyShape to shape 1 of slide {i + 2} of p',
                      'set left position of bodyShape to 40',
                      'set top of bodyShape to 30',
                      'set width of bodyShape to ((slide width of page setup of p) - 80)',
                      'set height of bodyShape to 120',
                      'set font size of font of text range of text frame of bodyShape to 24']
    lines += ['delete slide 1 of p', 'save p', 'close p saving no',
              f'open POSIX file {out}',
              'repeat 50 times',
              f'if full name of active presentation is {out} then exit repeat',
              'delay 0.1', 'end repeat',
              f'if full name of active presentation is not {out} then error "Reopen not activated"',
              f'set p to presentation {name}',
              f'if (count of slides of p) is not {len(questions)} then error "Reopen count mismatch"',
              'save p',
              'return version & linefeed & (full name of add in "RainClassroom.ppam")',
              'end tell', 'end timeout']
    result = subprocess.run(['osascript', '-'], input='\n'.join(lines), text=True,
                            capture_output=True, timeout=110)
    if result.returncode:
        raise RuntimeError(result.stderr + '\n请检查目标文稿中的实际页数与题目，从已完成的位置继续。')
    version, addin_path = result.stdout.strip().split('\n', 1)
    return {'platform': 'macOS', 'powerpoint_version': version, 'addin_path': addin_path}

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('questions', type=Path)
    parser.add_argument('output', type=Path)
    args = parser.parse_args()
    questions = json.loads(args.questions.read_text(encoding='utf-8'))
    if not isinstance(questions, list) or not questions:
        raise ValueError('题目文件必须是非空 JSON 数组')
    types = {0: None, 1: 'MultipleChoice', 2: 'MultipleChoiceMA', 3: 'Polling',
             4: 'FillBlank', 5: 'ShortAnswer', 6: 'MultipleChoice'}
    for q in questions:
        if q.get('ProblemType') not in types or not isinstance(q.get('Body'), str) or not q['Body'].strip():
            raise ValueError('每题需要受支持的 ProblemType 和非空 Body')
        if q['ProblemType'] in (1, 2, 3) and len(q.get('Options', [])) < 2:
            raise ValueError('选择/投票题至少需要两个 Options')
    output = args.output.resolve()
    environment = create(output, questions)
    report = inspect(output)
    if report['slide_count'] != len(questions) or report['has_embedded_vba']:
        raise RuntimeError('生成文件页数或格式不符合预期')
    for i, (slide, q) in enumerate(zip(report['slides'], questions), 1):
        expected = types[q['ProblemType']]
        if expected and slide['tags'].get('RAINPROBLEM') != expected:
            raise RuntimeError(f'第 {i} 页题型不符合预期')
    print(json.dumps({'status': 'created_and_reopened', 'path': str(output),
                     'environment': environment, **report},
                     ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
