"""
PPT 슬라이드를 PNG 이미지로 내보내는 스크립트
PowerPoint COM 자동화를 사용합니다.
"""

import win32com.client
import os
import time

# 파일명에 한글이 포함되어 있어 os.listdir()로 실제 경로를 찾음
target_keyword = '프로젝트'
base_dir = os.path.abspath('.')
pptx_filename = None
for f in os.listdir(base_dir):
    if target_keyword in f and f.endswith('.pptx'):
        pptx_filename = f
        break

if not pptx_filename:
    raise FileNotFoundError(f'*프로젝트*.pptx 파일을 찾을 수 없습니다.')

pptx_path = os.path.join(base_dir, pptx_filename)
output_dir = os.path.join(base_dir, 'assets', 'slides')
os.makedirs(output_dir, exist_ok=True)

print(f'PPT 파일: {pptx_filename}')
powerpoint = win32com.client.Dispatch("PowerPoint.Application")
powerpoint.Visible = True

presentation = powerpoint.Presentations.Open(pptx_path)
time.sleep(1)

total = presentation.Slides.Count
print(f'총 슬라이드 수: {total}')

for i in range(1, total + 1):
    out_path = os.path.join(output_dir, f'slide_{i:02d}.png')
    presentation.Slides(i).Export(out_path, 'PNG', 1280, 720)
    print(f'  저장: slide_{i:02d}.png')

presentation.Close()
powerpoint.Quit()
print('완료!')
