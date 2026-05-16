"""
PPT 슬라이드를 PNG 이미지로 내보내는 스크립트
PowerPoint COM 자동화를 사용합니다.
"""

import win32com.client
import os
import time

pptx_path = os.path.abspath('fraud_detection_ppt.pptx')
output_dir = os.path.abspath('assets/slides')

os.makedirs(output_dir, exist_ok=True)

print(f'PPT 열기: {pptx_path}')
powerpoint = win32com.client.Dispatch("PowerPoint.Application")
powerpoint.Visible = True

presentation = powerpoint.Presentations.Open(pptx_path)
time.sleep(1)

total = presentation.Slides.Count
print(f'총 슬라이드 수: {total}')

for i in range(1, total + 1):
    slide = presentation.Slides(i)
    out_path = os.path.join(output_dir, f'slide_{i:02d}.png')
    # Export(path, filter, width, height)
    slide.Export(out_path, 'PNG', 1280, 720)
    print(f'  저장: slide_{i:02d}.png')

presentation.Close()
powerpoint.Quit()
print('완료!')
