#! python3
# ------------------------------------------------------------------------------
# the Barbarian Tools™
# Kyakuhon Text Formatter
# ------------------------------------------------------------------------------

import enum
import os
import re
from docx import Document
from docx.shared import Inches
from docx.shared import Pt
from docx.shared import RGBColor


# ------------------------------------------------------------------------------

class Decoration(enum.Enum):
    # Decoration
    RESET = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
    REVERSE = '\033[7m'
    # Text Color
    RED = '\033[31m'
    GREEN = '\033[32m'
    BLUE = '\033[34m'
    CYAN = '\033[36m'
    MAGENTA = '\033[35m'
    YELLOW = '\033[33m'
    BLACK = '\033[30m'
    WHITE = '\033[37m'


class Status(enum.Enum):
    MESSAGE = '🤖 '
    FAILURE = Decoration.RED.value + '❌ [failure] ' + Decoration.RESET.value
    SUCCESS = Decoration.GREEN.value + '✅ [success] ' + Decoration.RESET.value
    CAUTION = Decoration.YELLOW.value + '⚠️ [caution] ' + Decoration.RESET.value
    PROSESSING = Decoration.CYAN.value + '⌛ [prosessing] ' + Decoration.RESET.value


class LineAttribute(enum.Enum):
    HASHIRA = Decoration.YELLOW.value    # Slugline
    TOGAKI = Decoration.BLUE.value       # Action
    SERIFU = Decoration.GREEN.value      # Dialogue
    BUNRITAI = Decoration.RED.value      # Transition
    MIDASHI = Decoration.CYAN.value
    KAIGYO = ''


def println_col(text: str, col: Decoration):
    print(col.value + text + Decoration.RESET.value)


def message(text: str, status: Status):
    escs = ''
    match status:
        case Status.FAILURE:
            escs = Decoration.RED.value
        case Status.SUCCESS:
            escs = Decoration.GREEN.value
        case Status.CAUTION:
            escs = Decoration.YELLOW.value
        case Status.MESSAGE:
            escs = Decoration.CYAN.value
        case Status.PROSESSING:
            escs = Decoration.CYAN.value
        case _:
            pass

    print(escs + status.value + escs + text + Decoration.RESET.value)


# ------------------------------------------------------------------------------

def load_text_file(file_path: str) -> list[str] | None:
    """テキストファイルを読み込んで、その文字列を行ごとのリストにして返す。"""
    # パスの存在を確認
    if not os.path.isfile(file_path):
        message('There is no file.', Status.FAILURE)
        return None

    # 拡張子を確認
    if os.path.splitext(file_path)[1] != '.txt':
        message('This is not a .txt file.', Status.FAILURE)
        return None

    # ファイルを開く
    with open(file_path, mode='r', encoding='utf_8') as f:
        lines = f.read().splitlines()
        message('Text has been read.', Status.SUCCESS)
        return lines


def text_preprocessor(lines: list[str]) -> list[str]:
    """各文字列に前処理を加えて返す。"""
    # 文字列前後の空白文字（半角・全角スペース／タブ／改行）を削除
    buffer = [s.strip() for s in lines]
    # メモ：先頭リストが空の場合とか、何かありそう
    return buffer


def add_attribute_to_line(lines: list[str]) -> list[list[str, LineAttribute]]:
    """行単位で判断できる脚本内属性を各文字列に付与した多次元リストを返す。"""
    # 柱        ：行の先頭に○□記号が存在
    hashira = re.compile(r'^[○□].+')
    # セリフ    ：行が「」記号で囲まれた文で終了
    serifu = re.compile(r'^.*「.+」$')
    # 分離帯    ：「×　　　　　×　　　　　×」
    bunritai = re.compile(r'^×\s+×\s+×$')
    # 見出し    ：行が【】記号で囲まれている
    midashi = re.compile(r'^【.+】$')
    # 改行：空文字列（現状）
    # ト書：その他の文字列

    line_with_attributes = []
    for l in lines:
        if hashira.match(l) is not None:
            line_with_attributes.append([l, LineAttribute.HASHIRA])
        elif serifu.match(l) is not None:
            line_with_attributes.append([l, LineAttribute.SERIFU])
        elif bunritai.match(l) is not None:
            line_with_attributes.append([('　' + l), LineAttribute.BUNRITAI])
        elif midashi.match(l) is not None:
            line_with_attributes.append([l, LineAttribute.MIDASHI])
        elif l == '':
            line_with_attributes.append([l, LineAttribute.KAIGYO])
        else:
            line_with_attributes.append([l, LineAttribute.TOGAKI])

    # メモ：本文の前（最初の柱の前）の情報の扱い一考

    return line_with_attributes


def fix_line_breaks(line_with_attributes: list[list[str, LineAttribute]]) -> list[list[str, LineAttribute]]:
    """コンテキストを基に改行が適切かチェックし修正して返す。"""
    # ・柱の前後は常に空白
    # ・違う属性との切り替わりに空白

    buffer = []
    context = line_with_attributes[0][1]  # 先頭の改行防止策としてとりあえず
    for lwa in line_with_attributes:
        lwa[1]
        if context == lwa[1]:
            # 柱が連続していたら改行を挿入する。
            if lwa[1] == LineAttribute.HASHIRA:
                buffer.append(['', LineAttribute.KAIGYO])
        else:
            # 改行以外の違う属性が連続していたら改行を挿入する。
            if (lwa[1] != LineAttribute.KAIGYO) and (context != LineAttribute.KAIGYO):
                buffer.append(['', LineAttribute.KAIGYO])
        # 再構築
        buffer.append(lwa)
        # コンテクスト記録
        context = lwa[1]

    return buffer


def format_to_docx(line_with_attributes: list[list[str, LineAttribute]]) -> Document:
    """docx形式で文字列を整形してDocumentオブジェクトを返す。"""
    # docxオブジェクトを作成
    doc = Document()
    # ページ設定
    section = doc.sections[0]
    # 用紙サイズ：A4 (210 x 297 mm)
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    # 用紙余白：254 mm (1 inch)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)

    for lwa in line_with_attributes:
        # docxパラグラフの追加
        pgh = None
        if lwa[1] == LineAttribute.HASHIRA:
            pgh = doc.add_heading(lwa[0], 3)
            pgh.style.font.color.rgb = RGBColor(102, 102, 102)
        else:
            pgh = doc.add_paragraph(lwa[0])
        # 文字のスタイル
        pgh.style.font.size = Pt(12)
        pgh.style.font.name = 'Arial'
        # パラグラフのフォーマット
        pgh_format = pgh.paragraph_format
        pgh_format.space_before = Pt(0)
        pgh_format.space_after = Pt(0)
        pgh_format.line_spacing = Pt(18)  # 1.2pt * 1.5 = 1.8pt

        match lwa[1]:
            case LineAttribute.SERIFU:
                # インデント：１段
                pgh_format.left_indent = Inches(0.5)
            case LineAttribute.TOGAKI:
                # インデント：２段
                pgh_format.left_indent = Inches(1)
            case LineAttribute.BUNRITAI:
                # インデント：２段
                pgh_format.left_indent = Inches(1)
            case _:
                pass

    return doc


"""
def save_plain_text_file(text: str, path: str):
    f = open(path, encoding='utf_8', mode='w')
    f.write(text)
    f.close()
"""


# ------------------------------------------------------------------------------

def ktf():
    message('Drag and drop the source text file and Enter.', Status.MESSAGE)
    file_path = input()

    lines = load_text_file(file_path)
    # barrier
    if lines == None:
        return None

    lines = text_preprocessor(lines)
    line_with_attributes = add_attribute_to_line(lines)
    line_with_attributes = fix_line_breaks(line_with_attributes)

    """
    for lwa in line_with_attributes:
        println_col(lwa[0], lwa[1])
    """
    message('Generating .docx file. Please wait...', Status.PROSESSING)

    doc = format_to_docx(line_with_attributes)

    dirname = os.path.dirname(file_path)
    basename_without_ext = os.path.splitext(os.path.basename(file_path))[0]
    target_path = os.path.join(dirname, basename_without_ext + '.docx')
    doc.save(target_path)

    message('.docx file Generated at:\n' + target_path, Status.SUCCESS)


def eyecatch():
    """アイキャッチ"""
    str = '-----------------------\n'\
          '🍖 the Barbarian Tools™\n'\
          'Kyakuhon Text Formatter\n'\
          '-----------------------\n'\
          'Beta             v0.1.0\n'\
          '-----------------------'
    println_col(str, Decoration.RED)


# ------------------------------------------------------------------------------

if __name__ == '__main__':
    try:
        eyecatch()
        while True:
            ktf()
            println_col('\n-------Re-Start--------', Decoration.YELLOW)
            message('(Exit: Ctrl + C)', Status.MESSAGE)
    except KeyboardInterrupt:
        println_col('Process KILLED.', Decoration.MAGENTA)
