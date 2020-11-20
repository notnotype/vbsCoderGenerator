import os
import sys

from pypinyin import lazy_pinyin

keymap = {
    "Shift": "+",
    'Ctrl': "^",
    "Alt": "%",
    "BACKSPACE": "{BACKSPACE}",
    "BREAK": "{BREAK}",
    'CAPS LOCK': "{CAPSLOCK}",
    'DEL': "{DELETE}",
}


# DOWN ARROW    {DOWN}
# END  {END}
# ENTER     {ENTER}or ~
# ESC      {ESC}
# HELP      {HELP}
# HOME      {HOME}
# INS or INSERT {INSERT} or {INS}
# LEFT ARROW     {LEFT}
# NUM LOCK     {NUMLOCK}
# PAGE DOWN     {PGDN}
# PAGE UP     {PGUP}
# PRINT SCREEN    {PRTSC}
# RIGHT ARROW    {RIGHT}
# SCROLL LOCK    {SCROLLLOCK}
# TAB     {TAB}
# UP ARROW    {UP}
# F1 {F1}
def switch_typing() -> list:
    snippet = [
        "' 打开输入法",
        f'WshShell.SendKeys "^+"',
        f'WScript.Sleep 100'
    ]
    return snippet


def my_close_typing() -> list:
    snippet = [
        f'WshShell.SendKeys "^+"',
        f'WScript.Sleep 150'
    ]
    return snippet


def general_close_type() -> list:
    snippet = [
        f'WshShell.SendKeys "^ "',
        f'WScript.Sleep 150'
    ]
    return snippet


def zh_translate(han: str) -> list:
    snippet = []
    snippet += switch_typing()
    for char in han:
        snippet.append(f'WshShell.SendKeys {mapping(char)[-1]}')
        snippet.append(f'WScript.Sleep {sep_time}')
    snippet.append('WshShell.SendKeys " "')
    snippet += close_typing()
    return snippet


def mapping(char: str) -> tuple:
    if char == '\n':
        return 'str', '"{ENTER}"'
    elif char == '{':
        return 'str', '"{{}"'
    elif char == '}':
        return 'str', '"{}}"'
    elif char == '\"':
        return 'str', 'Chr(34)'
    elif char == '(':
        return 'str', '\"{(}\"'
    elif char == ')':
        return 'str', '\"{)}\"'
    elif char == '%':
        return 'str', '\"{%}\"'
    elif ord(char) not in range(0, 128):
        try:
            py = lazy_pinyin(char)[0]
        except AssertionError:
            py = 'cuowu'
        return 'snippet', zh_translate(py)
    else:
        return 'str', f"\"{char}\""


sep_time = 5
zh_char_sep = 30
wait_notepad_time = 700


def translate(code: str) -> str:
    vbs_header = [
        "''''''''''''''''''''''''''",
        "' author: Notype",
        "' github: http://www.github.com/notnotype",
        "' date: 2020/11/21 0:31",
        "''''''''''''''''''''''''''",
        "''''''''''头文件''''''''''''''''",
        'Set WshShell=WScript.CreateObject("WScript.Shell")',
        'WshShell.AppActivate ""',
        'WshShell.run "notepad"',
        f'WScript.Sleep {wait_notepad_time}',
        'MsgBox "Code auto complete, Are you ready?"',
        f'WScript.Sleep {wait_notepad_time}',
        "''''''''''''''''''''''''''",
    ]

    """
    `snippet`像这种格式
    WshShell.SendKeys " "
    WScript.Sleep 30
    """
    code_snippet = []
    for char in code:
        snippet = []

        str_or_snippet = mapping(char)
        if str_or_snippet[0] == 'str':
            snippet.append(f'WshShell.SendKeys {str_or_snippet[-1]}')
            snippet.append(f'WScript.Sleep {sep_time}')
        else:
            snippet += str_or_snippet[-1]

        code_snippet += snippet

    translate_result = '\n'.join(code_snippet)
    translate_result = "\n".join(vbs_header) + '\n' + translate_result

    before_zh_optimize = translate_result
    optimize_snippet = ['WshShell.SendKeys " "']
    optimize_snippet += close_typing()
    optimize_snippet += switch_typing()
    after_zh_optimize = before_zh_optimize.replace("\n".join(optimize_snippet), "' 优化中文输入")
    translate_result = after_zh_optimize

    return translate_result


if __name__ == '__main__':
    os.environ.setdefault('whouse', 'me')

    if os.environ.get('whouse', default='you') == 'me':
        close_typing = my_close_typing
    else:
        close_typing = general_close_type
    # student.h
    old_my_code = \
        '''adsfdsaf'''
    old_my_code2 = '你好世界,我是傻逼,啦哈剌是我儿子哈哈哈哈哈哈哈'

    my_code = ''
    with open(sys.argv[0], 'r', encoding='utf8') as f:
        my_code = f.read()

    r = translate(my_code)
    with open('textcode.vbs', 'w', encoding='utf8') as f:
        f.write(r)
