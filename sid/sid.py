#! python3
# ------------------------------------------------------------------------------=-------------------
# the Barbarian Tools™
# Similar Images Detector
# ------------------------------------------------------------------------------=-------------------

from enum import Enum
from pathlib import Path
from itertools import combinations
import numpy as np
import warnings

# [💪野蛮] torchvision.models._utilsからのUserWarningをすべて無効化
warnings.filterwarnings('ignore', category=UserWarning, module='torchvision.models._utils')

import imgsim
import cv2
#import webview

# ------------------------------------------------------------------------------=-------------------

def main():
    message('Drag and drop the directry and Enter.', Status.MESSAGE)
    path_obj = Path(input())
    
    message('Searching... (jpeg, png)', Status.PROSESS)

    # ディレクトリの存在を確認
    if not path_obj.is_dir():
        message('There is no directory.', Status.FAILURE)
        return None

    # パス以下を再帰的に走査
    extensions = {'.jpg', '.jpeg', '.png'}
    files = []
    for f in path_obj.rglob('*'):
        # 拡張子を確認
        if f.is_file() and f.suffix.lower() in extensions:
            files.append(f)

    # ファイルの有無
    if not files:
        message('No images found.', Status.FAILURE)
        return None
    
    #print('\n'.join(map(str, files)))
    total_calc = len(files) * (len(files) - 1) // 2
    message(f'{len(files)} image files found. {total_calc} calculations will be needed.', 
            Status.SUCCESS)

    # 全ての画像を読み込んで特徴ベクトルに変換（AugNet）
    message('Vectorizing images...', Status.PROSESS)
    vtr = imgsim.Vectorizer()
    vectors = {}
    count = 0
    clear_line = ' ' * 120
    for file_path in files:
        count += 1
        print(f'\r{clear_line}\r🧐 {count}/{len(files)} ...  {str(file_path)}', end='', flush=True)
        img = read_image(file_path)
        vectors[file_path] = vtr.vectorize(img)
    print()
    message('All images have been vectorized.', Status.SUCCESS)
    
    # 総当たりでスコアを出す
    message('Calculating distances...', Status.PROSESS)
    results = []
    for path1, path2 in combinations(files, 2):
        vec1 = vectors[path1]
        vec2 = vectors[path2]
        distance = imgsim.distance(vec1, vec2)
        results.append((path1, path2, distance))
    
    # 以下超雑

    results.sort(key=lambda x: x[2])
    results_for_html = []
    for path1, path2, dist in results[:100]:
        print(f'{dist:.4f}: {path1.name} <-> {path2.name}')
        results_for_html.append((path1, path2, dist, f'{path1}<br>{path2}'))

    #print(path_obj.resolve())
    create_html(results_for_html, path_obj.resolve())
    #print(str(path_obj.resolve() / 'output.html'))

    #window = webview.create_window('🍖 the Barbarian Tools™ - Similar Images Detector', str(path_obj.resolve() / 'output.html'), width=960, height=1080)
    #webview.start()

def read_image(file_path):
    # 日本語パスの画像対応。挙動要確認
    # バイナリで読んでnumpy配列に変換
    with open(file_path, 'rb') as f:
        img_array = np.frombuffer(f.read(), dtype=np.uint8)
    # デコード
    img = cv2.imdecode(img_array, cv2.IMREAD_COLOR)
    return img

def create_html(results, target_path):
    html = ['<!DOCTYPE html><html><head><meta charset="UTF-8"><style>img {max-height: 240px;} p {margin: 0 0 1em;}</style></head><body>']
    for path1, path2, distance, name in results:
        html.append(f'<div><img src="{path1}"><img src="{path2}"><p>{distance:.4f}:<br>{name}</p></div>')
    html.append('</body></html>')
    print(target_path / '0_output.html')
    with open(target_path / '0_output.html', 'w', encoding='utf-8') as f:
        f.write('\n'.join(html))

# ------------------------------------------------------------------------------=-------------------

class Decoration(Enum):
    # Decoration
    RESET =     '\033[0m'
    BOLD =      '\033[1m'
    UNDERLINE = '\033[4m'
    REVERSE =   '\033[7m'
    # Text Color
    RED =       '\033[31m'
    GREEN =     '\033[32m'
    BLUE =      '\033[34m'
    CYAN =      '\033[36m'
    MAGENTA =   '\033[35m'
    YELLOW =    '\033[33m'
    BLACK =     '\033[30m'
    WHITE =     '\033[37m'

class Status(Enum):
    MESSAGE = Decoration.CYAN.value +   '🔽 '
    PROSESS = Decoration.CYAN.value +   '⌛ Prosessing | '
    FAILURE = Decoration.RED.value +    '❌ Failure    | '
    SUCCESS = Decoration.GREEN.value +  '✅ Success    | '
    CAUTION = Decoration.YELLOW.value + '⚠️ Caution    | '

def print_deco(text: str, deco: Decoration):
    print(deco.value + text + Decoration.RESET.value)

def message(text: str, status: Status):
    print(status.value + text + Decoration.RESET.value)

def eyecatch():
    str = '-----------------------\n'\
          '🍖 the Barbarian Tools™\n'\
          'Similar Images Detector\n'\
          '-----------------------\n'\
          'Beta             v0.1.0\n'\
          '-----------------------'
    print_deco(str, Decoration.RED)

# ------------------------------------------------------------------------------=-------------------

if __name__ == '__main__':
    eyecatch()
    main()
    input()
