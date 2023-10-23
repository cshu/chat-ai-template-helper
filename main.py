#!/usr/bin/env python3

import docx
import openai
import json
import argparse
import tempfile
import os
#import pyperclip
import sys
import subprocess
import re
import shutil
import xml.sax.saxutils
from pathlib import Path
from os import listdir
from os.path import isfile
from datetime import datetime

parser = argparse.ArgumentParser()
parser.add_argument('template')
nmsce: argparse.Namespace = parser.parse_args()
tfile: str = nmsce.template
tobj: dict = json.loads(Path(tfile).read_text())
strip_end: str = tobj['strip_end']
strip_beg: str = tobj['strip_beg']
min_num_of_lines_after_strip: int = tobj['min_num_of_lines_after_strip']
txtbegin: str = tobj['begin']
txtend: str = tobj['end']
txtmiddle: str = tobj['middle']
editor: list[str] = tobj['editor']
substitution: dict = tobj['substitution']
docxfile: str = tobj['docx']
def_system_content: str = tobj['def_system_content']
def_user_content: str = tobj['def_user_content']
openai.api_key = Path(tobj['keyfile']).read_text().rstrip()


#substitution: str = pyperclip.paste()
#substitution:str = subprocess.check_output(['xclip', '-t', 'text/html', '-o', '-selection', 'clipboard']).decode()
#def_user_content = def_user_content.replace('{CLIPBOARD_TEXT_STUB}', substitution)
for k, v in substitution.items():
 def_user_content = def_user_content.replace(k, Path(v).read_text())
print('##############')
print(def_user_content)
print('##############')

completion = openai.ChatCompletion.create(
 model="gpt-3.5-turbo-16k",
 messages=[
  {"role": "system", "content": def_system_content},
  {"role": "user", "content": def_user_content}
 ]
)
result = completion.choices[0].message.content# type: ignore

#print('##############')
#print(result)
#print(completion.choices[0].message)
#print(completion.choices)
#print(completion)

#pyperclip.copy(result)
slines = result.splitlines()
begidx = 0
endidx = len(slines)
for idx, ln in enumerate(slines):
 if (not begidx) and re.match(strip_beg, ln, re.IGNORECASE):
  begidx = idx+1
 elif re.match(strip_end, ln, re.IGNORECASE):
  endidx = idx
if endidx - begidx > min_num_of_lines_after_strip:
 result = '\n'.join(slines[begidx: endidx])
Path(txtmiddle).write_text(result)
editor.append(txtmiddle)
subprocess.run(editor, check=True)
input('ENTER after you have finished editing the txt')

result=txtbegin+Path(txtmiddle).read_text()+txtend

document = docx.Document(docxfile)
document.add_paragraph(result)
document.save(docxfile+'.out.docx')
