import os
import sys

from pypinyin import lazy_pinyin
import click

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
sep_switch_typing = 100
sep_close_typing = 150
sep_time = 5
zh_char_sep = 30
wait_notepad_time = 700


def switch_typing() -> list:
    snippet = [
        "' 打开输入法",
        f'WshShell.SendKeys "^+"',
        f'WScript.Sleep {sep_switch_typing}'
    ]
    return snippet


def my_close_typing() -> list:
    snippet = [
        f'WshShell.SendKeys "^+"',
        f'WScript.Sleep {sep_close_typing}'
    ]
    return snippet


def general_close_type() -> list:
    snippet = [
        f'WshShell.SendKeys "^ "',
        f'WScript.Sleep {sep_close_typing}'
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


def translate(code: str) -> str:
    vbs_header = [
        "''''''''''''''''''''''''''",
        "' author: Notype",
        "' github: http://www.github.com/notnotype",
        "' date: 2020/11/21 0:31",
        '\'whouse:            {}'.format('me' if close_typing == my_close_typing else 'you'),
        '\'sep_time:          {}'.format(str(sep_time)),
        '\'sep_switch_typing: {}'.format(str(sep_switch_typing)),
        '\'zh_char_sep:       {}'.format(str(zh_char_sep)),
        '\'wait_notepad_time: {}'.format(str(wait_notepad_time)),
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


@click.command()
@click.option('--whouse', '-w', default='you', help='谁用?')
@click.option('--fromfile', '-f', type=click.File(encoding='utf-8'), help='从文件加载输入')
@click.option('--text', '-t', required=True, help='要转换成vbs的文本')
@click.option('--outfile', '-o', type=click.File(encoding='utf8'), help='输入文件')
@click.option('--sep-time', '-s', help='每一个字母的间隔速度')
@click.option('--zh-char-sep', '-zcs', '-z', help='中文打字间隔速度')
@click.option('--sep-switch-typing', '-sst', '-s', help='打开输入法速度')
@click.option('--wait-notepad-time', '-wnt', help='等待记事本打开时间')
@click.option('--sep-close-typing', '-sct', help='关闭输入法时间')
@click.option('--alert', '-a', help='提示警告')
def main(whouse, fromfile, text, outfile, **kwargs):
    """
    功能: 把一串文本转换为vbs自动编写这段代码的vbs脚本\n
    示例: `python main.py --text 我是傻逼`
    """
    global close_typing
    global sep_time
    global sep_switch_typing
    global zh_char_sep
    global wait_notepad_time

    if kwargs['sep_time']:
        sep_time = kwargs['sep_time']
    if kwargs['sep_switch_typing']:
        sep_switch_typing = kwargs['sep_switch_typing']
    if kwargs['zh_char_sep']:
        zh_char_sep = kwargs['zh_char_sep']
    if kwargs['wait_notepad_time']:
        wait_notepad_time = kwargs['wait_notepad_time']

    print('======参数列表======')
    print('whouse:            {}' .format(str(whouse)))
    print('sep_time:          {}' .format(str(sep_time)))
    print('sep_switch_typing: {}' .format(str(sep_switch_typing)))
    print('zh_char_sep:       {}' .format(str(zh_char_sep)))
    print('wait_notepad_time: {}' .format(str(wait_notepad_time)))
    print('====================')

    my_code = text

    if whouse == 'me':
        close_typing = my_close_typing
    else:
        close_typing = general_close_type

    if fromfile:
        my_code = fromfile.read()
        print('转换前字数: {}'.format(len(my_code)))

    r = translate(my_code)
    print('转换后字数: {}'.format(len(r)))
    if not outfile:
        with open('textcode.vbs', 'w', encoding='utf8') as f:
            f.write(r)
    else:
        outfile.write(r)
    print('输出完成')


if __name__ == '__main__':
    main()
