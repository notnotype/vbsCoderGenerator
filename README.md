# vbsCoderGenerator
 python生成vbs写代码脚本
 
# 快速开始
```shell script
git https://github.com/notnotype/vbsCoderGenerator
cd vbsCoderGenerator
pip install -r requirements.txt
./readme.vbs
```
 

# 演示
```text
>python main.py --zh-char-sep 10 --sep-time 1 --text "this is a 测试" -f "main.py" --whouse you
======参数列表======
whouse:            me
sep_time:          1
sep_switch_typing: 100
zh_char_sep:       10
wait_notepad_time: 700
====================
转换前字数: 5949
转换后字数: 242323
输出完成
```

# 命令提醒
> 输入 `python main.py --help`查看
```
python main.py --help
Usage: main.py [OPTIONS]

  功能: 把一串文本转换为vbs自动编写这段代码的vbs脚本

  示例: `python main.py --text 我是傻逼`

Options:
  -w, --whouse TEXT               谁用?
  -f, --fromfile FILENAME         从文件加载输入
  -t, --text TEXT                 要转换成vbs的文本  [required]
  -o, --outfile FILENAME          输入文件
  -s, --sep-time TEXT             每一个字母的间隔速度
  -zcs, -z, --zh-char-sep TEXT    中文打字间隔速度
  -sst, -s, --sep-switch-typing TEXT
                                  打开输入法速度
  -wnt, --wait-notepad-time TEXT  等待记事本打开时间
  -sct, --sep-close-typing TEXT   关闭输入法时间
  -a, --alert TEXT                提示警告
  --help                          Show this message and exit.

```
