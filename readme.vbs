''''''''''''''''''''''''''
' author: Notype
' github: http://www.github.com/notnotype
' date: 2020/11/21 0:31
'whouse:            you
'sep_time:          1
'sep_switch_typing: 100
'zh_char_sep:       10
'wait_notepad_time: 700
''''''''''''''''''''''''''
''''''''''头文件''''''''''''''''
Set WshShell=WScript.CreateObject("WScript.Shell")
WshShell.AppActivate ""
WshShell.run "notepad"
WScript.Sleep 700
MsgBox "Code auto complete, Are you ready?"
WScript.Sleep 700
''''''''''''''''''''''''''
WshShell.SendKeys "#"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "v"
WScript.Sleep 1
WshShell.SendKeys "b"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "C"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
WshShell.SendKeys "d"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "r"
WScript.Sleep 1
WshShell.SendKeys "G"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "r"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
WshShell.SendKeys "r"
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "p"
WScript.Sleep 1
WshShell.SendKeys "y"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
' 打开输入法
WshShell.SendKeys "^+"
WScript.Sleep 100
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "g"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "c"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "g"
WScript.Sleep 1
WshShell.SendKeys " "
WshShell.SendKeys "^ "
WScript.Sleep 150
WshShell.SendKeys "v"
WScript.Sleep 1
WshShell.SendKeys "b"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
' 打开输入法
WshShell.SendKeys "^+"
WScript.Sleep 100
WshShell.SendKeys "x"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "d"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "m"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "j"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "b"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys " "
WshShell.SendKeys "^ "
WScript.Sleep 150
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys "#"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
' 打开输入法
WshShell.SendKeys "^+"
WScript.Sleep 100
WshShell.SendKeys "k"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "k"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys " "
WshShell.SendKeys "^ "
WScript.Sleep 150
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys "`"
WScript.Sleep 1
WshShell.SendKeys "`"
WScript.Sleep 1
WshShell.SendKeys "`"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "l"
WScript.Sleep 1
WshShell.SendKeys "l"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "c"
WScript.Sleep 1
WshShell.SendKeys "r"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "p"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys "g"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "p"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys ":"
WScript.Sleep 1
WshShell.SendKeys "/"
WScript.Sleep 1
WshShell.SendKeys "/"
WScript.Sleep 1
WshShell.SendKeys "g"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
WshShell.SendKeys "b"
WScript.Sleep 1
WshShell.SendKeys "."
WScript.Sleep 1
WshShell.SendKeys "c"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
WshShell.SendKeys "m"
WScript.Sleep 1
WshShell.SendKeys "/"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "y"
WScript.Sleep 1
WshShell.SendKeys "p"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "/"
WScript.Sleep 1
WshShell.SendKeys "v"
WScript.Sleep 1
WshShell.SendKeys "b"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "C"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
WshShell.SendKeys "d"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "r"
WScript.Sleep 1
WshShell.SendKeys "G"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "r"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
WshShell.SendKeys "r"
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys "c"
WScript.Sleep 1
WshShell.SendKeys "d"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "v"
WScript.Sleep 1
WshShell.SendKeys "b"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "C"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
WshShell.SendKeys "d"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "r"
WScript.Sleep 1
WshShell.SendKeys "G"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "r"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
WshShell.SendKeys "r"
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys "p"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "p"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "l"
WScript.Sleep 1
WshShell.SendKeys "l"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "r"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "r"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "q"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "r"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "m"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "."
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "x"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys "."
WScript.Sleep 1
WshShell.SendKeys "/"
WScript.Sleep 1
WshShell.SendKeys "r"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "d"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "."
WScript.Sleep 1
WshShell.SendKeys "v"
WScript.Sleep 1
WshShell.SendKeys "b"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys "`"
WScript.Sleep 1
WshShell.SendKeys "`"
WScript.Sleep 1
WshShell.SendKeys "`"
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys "#"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
' 打开输入法
WshShell.SendKeys "^+"
WScript.Sleep 100
WshShell.SendKeys "y"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys " "
WshShell.SendKeys "^ "
WScript.Sleep 150
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys "`"
WScript.Sleep 1
WshShell.SendKeys "`"
WScript.Sleep 1
WshShell.SendKeys "`"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "x"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys ">"
WScript.Sleep 1
WshShell.SendKeys "p"
WScript.Sleep 1
WshShell.SendKeys "y"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "m"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "."
WScript.Sleep 1
WshShell.SendKeys "p"
WScript.Sleep 1
WshShell.SendKeys "y"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "z"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "c"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "r"
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "p"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "1"
WScript.Sleep 1
WshShell.SendKeys "0"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "p"
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "m"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "1"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "x"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "1"
WScript.Sleep 1
WshShell.SendKeys "2"
WScript.Sleep 1
WshShell.SendKeys "3"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "f"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys Chr(34)
WScript.Sleep 1
WshShell.SendKeys "m"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "."
WScript.Sleep 1
WshShell.SendKeys "p"
WScript.Sleep 1
WshShell.SendKeys "y"
WScript.Sleep 1
WshShell.SendKeys Chr(34)
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "w"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "y"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys "="
WScript.Sleep 1
WshShell.SendKeys "="
WScript.Sleep 1
WshShell.SendKeys "="
WScript.Sleep 1
WshShell.SendKeys "="
WScript.Sleep 1
WshShell.SendKeys "="
WScript.Sleep 1
WshShell.SendKeys "="
WScript.Sleep 1
' 打开输入法
WshShell.SendKeys "^+"
WScript.Sleep 100
WshShell.SendKeys "c"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "l"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "b"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
WshShell.SendKeys " "
WshShell.SendKeys "^ "
WScript.Sleep 150
WshShell.SendKeys "="
WScript.Sleep 1
WshShell.SendKeys "="
WScript.Sleep 1
WshShell.SendKeys "="
WScript.Sleep 1
WshShell.SendKeys "="
WScript.Sleep 1
WshShell.SendKeys "="
WScript.Sleep 1
WshShell.SendKeys "="
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys "w"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys ":"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "m"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "p"
WScript.Sleep 1
WshShell.SendKeys "_"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "m"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys ":"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "1"
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "p"
WScript.Sleep 1
WshShell.SendKeys "_"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "w"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "c"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "_"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "y"
WScript.Sleep 1
WshShell.SendKeys "p"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "g"
WScript.Sleep 1
WshShell.SendKeys ":"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "1"
WScript.Sleep 1
WshShell.SendKeys "0"
WScript.Sleep 1
WshShell.SendKeys "0"
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys "z"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "_"
WScript.Sleep 1
WshShell.SendKeys "c"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "r"
WScript.Sleep 1
WshShell.SendKeys "_"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "p"
WScript.Sleep 1
WshShell.SendKeys ":"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "1"
WScript.Sleep 1
WshShell.SendKeys "0"
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys "w"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "_"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "p"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "d"
WScript.Sleep 1
WshShell.SendKeys "_"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "m"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys ":"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "7"
WScript.Sleep 1
WshShell.SendKeys "0"
WScript.Sleep 1
WshShell.SendKeys "0"
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys "="
WScript.Sleep 1
WshShell.SendKeys "="
WScript.Sleep 1
WshShell.SendKeys "="
WScript.Sleep 1
WshShell.SendKeys "="
WScript.Sleep 1
WshShell.SendKeys "="
WScript.Sleep 1
WshShell.SendKeys "="
WScript.Sleep 1
WshShell.SendKeys "="
WScript.Sleep 1
WshShell.SendKeys "="
WScript.Sleep 1
WshShell.SendKeys "="
WScript.Sleep 1
WshShell.SendKeys "="
WScript.Sleep 1
WshShell.SendKeys "="
WScript.Sleep 1
WshShell.SendKeys "="
WScript.Sleep 1
WshShell.SendKeys "="
WScript.Sleep 1
WshShell.SendKeys "="
WScript.Sleep 1
WshShell.SendKeys "="
WScript.Sleep 1
WshShell.SendKeys "="
WScript.Sleep 1
WshShell.SendKeys "="
WScript.Sleep 1
WshShell.SendKeys "="
WScript.Sleep 1
WshShell.SendKeys "="
WScript.Sleep 1
WshShell.SendKeys "="
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
' 打开输入法
WshShell.SendKeys "^+"
WScript.Sleep 100
WshShell.SendKeys "z"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "q"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "z"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
WshShell.SendKeys " "
WshShell.SendKeys "^ "
WScript.Sleep 150
WshShell.SendKeys ":"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "5"
WScript.Sleep 1
WshShell.SendKeys "9"
WScript.Sleep 1
WshShell.SendKeys "4"
WScript.Sleep 1
WshShell.SendKeys "9"
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
' 打开输入法
WshShell.SendKeys "^+"
WScript.Sleep 100
WshShell.SendKeys "z"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "z"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
WshShell.SendKeys " "
WshShell.SendKeys "^ "
WScript.Sleep 150
WshShell.SendKeys ":"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "2"
WScript.Sleep 1
WshShell.SendKeys "4"
WScript.Sleep 1
WshShell.SendKeys "2"
WScript.Sleep 1
WshShell.SendKeys "3"
WScript.Sleep 1
WshShell.SendKeys "2"
WScript.Sleep 1
WshShell.SendKeys "3"
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
' 打开输入法
WshShell.SendKeys "^+"
WScript.Sleep 100
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "c"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "w"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "c"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "g"
WScript.Sleep 1
WshShell.SendKeys " "
WshShell.SendKeys "^ "
WScript.Sleep 150
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys "`"
WScript.Sleep 1
WshShell.SendKeys "`"
WScript.Sleep 1
WshShell.SendKeys "`"
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys "#"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
' 打开输入法
WshShell.SendKeys "^+"
WScript.Sleep 100
WshShell.SendKeys "m"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "g"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "l"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "g"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "x"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "g"
WScript.Sleep 1
WshShell.SendKeys " "
WshShell.SendKeys "^ "
WScript.Sleep 150
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys ">"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
' 打开输入法
WshShell.SendKeys "^+"
WScript.Sleep 100
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "r"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
WshShell.SendKeys " "
WshShell.SendKeys "^ "
WScript.Sleep 150
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "`"
WScript.Sleep 1
WshShell.SendKeys "p"
WScript.Sleep 1
WshShell.SendKeys "y"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "m"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "."
WScript.Sleep 1
WshShell.SendKeys "p"
WScript.Sleep 1
WshShell.SendKeys "y"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "l"
WScript.Sleep 1
WshShell.SendKeys "p"
WScript.Sleep 1
WshShell.SendKeys "`"
WScript.Sleep 1
' 打开输入法
WshShell.SendKeys "^+"
WScript.Sleep 100
WshShell.SendKeys "c"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "k"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys " "
WshShell.SendKeys "^ "
WScript.Sleep 150
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys "`"
WScript.Sleep 1
WshShell.SendKeys "`"
WScript.Sleep 1
WshShell.SendKeys "`"
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys "p"
WScript.Sleep 1
WshShell.SendKeys "y"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "m"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "."
WScript.Sleep 1
WshShell.SendKeys "p"
WScript.Sleep 1
WshShell.SendKeys "y"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "l"
WScript.Sleep 1
WshShell.SendKeys "p"
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys "U"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "g"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys ":"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "m"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "."
WScript.Sleep 1
WshShell.SendKeys "p"
WScript.Sleep 1
WshShell.SendKeys "y"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "["
WScript.Sleep 1
WshShell.SendKeys "O"
WScript.Sleep 1
WshShell.SendKeys "P"
WScript.Sleep 1
WshShell.SendKeys "T"
WScript.Sleep 1
WshShell.SendKeys "I"
WScript.Sleep 1
WshShell.SendKeys "O"
WScript.Sleep 1
WshShell.SendKeys "N"
WScript.Sleep 1
WshShell.SendKeys "S"
WScript.Sleep 1
WshShell.SendKeys "]"
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
' 打开输入法
WshShell.SendKeys "^+"
WScript.Sleep 100
WshShell.SendKeys "g"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "g"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "g"
WScript.Sleep 1
WshShell.SendKeys " "
WshShell.SendKeys "^ "
WScript.Sleep 150
WshShell.SendKeys ":"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
' 打开输入法
WshShell.SendKeys "^+"
WScript.Sleep 100
WshShell.SendKeys "b"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "y"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "c"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "w"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "b"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "z"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "w"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys " "
WshShell.SendKeys "^ "
WScript.Sleep 150
WshShell.SendKeys "v"
WScript.Sleep 1
WshShell.SendKeys "b"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
' 打开输入法
WshShell.SendKeys "^+"
WScript.Sleep 100
WshShell.SendKeys "z"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "d"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "g"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "b"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "x"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "z"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "d"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "d"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "m"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "d"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys " "
WshShell.SendKeys "^ "
WScript.Sleep 150
WshShell.SendKeys "v"
WScript.Sleep 1
WshShell.SendKeys "b"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
' 打开输入法
WshShell.SendKeys "^+"
WScript.Sleep 100
WshShell.SendKeys "j"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "b"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys " "
WshShell.SendKeys "^ "
WScript.Sleep 150
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
' 打开输入法
WshShell.SendKeys "^+"
WScript.Sleep 100
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "l"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys " "
WshShell.SendKeys "^ "
WScript.Sleep 150
WshShell.SendKeys ":"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "`"
WScript.Sleep 1
WshShell.SendKeys "p"
WScript.Sleep 1
WshShell.SendKeys "y"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "m"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "."
WScript.Sleep 1
WshShell.SendKeys "p"
WScript.Sleep 1
WshShell.SendKeys "y"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "x"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
' 打开输入法
WshShell.SendKeys "^+"
WScript.Sleep 100
WshShell.SendKeys "w"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "b"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys " "
WshShell.SendKeys "^ "
WScript.Sleep 150
WshShell.SendKeys "`"
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys "O"
WScript.Sleep 1
WshShell.SendKeys "p"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys ":"
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "w"
WScript.Sleep 1
WshShell.SendKeys ","
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "w"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "T"
WScript.Sleep 1
WshShell.SendKeys "E"
WScript.Sleep 1
WshShell.SendKeys "X"
WScript.Sleep 1
WshShell.SendKeys "T"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
' 打开输入法
WshShell.SendKeys "^+"
WScript.Sleep 100
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "y"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "g"
WScript.Sleep 1
WshShell.SendKeys " "
WshShell.SendKeys "^ "
WScript.Sleep 150
WshShell.SendKeys "?"
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "f"
WScript.Sleep 1
WshShell.SendKeys ","
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "f"
WScript.Sleep 1
WshShell.SendKeys "r"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
WshShell.SendKeys "m"
WScript.Sleep 1
WshShell.SendKeys "f"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "l"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "F"
WScript.Sleep 1
WshShell.SendKeys "I"
WScript.Sleep 1
WshShell.SendKeys "L"
WScript.Sleep 1
WshShell.SendKeys "E"
WScript.Sleep 1
WshShell.SendKeys "N"
WScript.Sleep 1
WshShell.SendKeys "A"
WScript.Sleep 1
WshShell.SendKeys "M"
WScript.Sleep 1
WshShell.SendKeys "E"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
' 打开输入法
WshShell.SendKeys "^+"
WScript.Sleep 100
WshShell.SendKeys "c"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "g"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "w"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "j"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "j"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "z"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "r"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
WshShell.SendKeys " "
WshShell.SendKeys "^ "
WScript.Sleep 150
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys ","
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "x"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "T"
WScript.Sleep 1
WshShell.SendKeys "E"
WScript.Sleep 1
WshShell.SendKeys "X"
WScript.Sleep 1
WshShell.SendKeys "T"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
' 打开输入法
WshShell.SendKeys "^+"
WScript.Sleep 100
WshShell.SendKeys "y"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "z"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "c"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "g"
WScript.Sleep 1
WshShell.SendKeys " "
WshShell.SendKeys "^ "
WScript.Sleep 150
WshShell.SendKeys "v"
WScript.Sleep 1
WshShell.SendKeys "b"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
' 打开输入法
WshShell.SendKeys "^+"
WScript.Sleep 100
WshShell.SendKeys "d"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "w"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "b"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys " "
WshShell.SendKeys "^ "
WScript.Sleep 150
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "["
WScript.Sleep 1
WshShell.SendKeys "r"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "q"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "r"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "d"
WScript.Sleep 1
WshShell.SendKeys "]"
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
WshShell.SendKeys ","
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "f"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "l"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "F"
WScript.Sleep 1
WshShell.SendKeys "I"
WScript.Sleep 1
WshShell.SendKeys "L"
WScript.Sleep 1
WshShell.SendKeys "E"
WScript.Sleep 1
WshShell.SendKeys "N"
WScript.Sleep 1
WshShell.SendKeys "A"
WScript.Sleep 1
WshShell.SendKeys "M"
WScript.Sleep 1
WshShell.SendKeys "E"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
' 打开输入法
WshShell.SendKeys "^+"
WScript.Sleep 100
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "r"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "w"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "j"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys " "
WshShell.SendKeys "^ "
WScript.Sleep 150
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys ","
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "p"
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "m"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "T"
WScript.Sleep 1
WshShell.SendKeys "E"
WScript.Sleep 1
WshShell.SendKeys "X"
WScript.Sleep 1
WshShell.SendKeys "T"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
' 打开输入法
WshShell.SendKeys "^+"
WScript.Sleep 100
WshShell.SendKeys "m"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "y"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "g"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "z"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "m"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "d"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "j"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "g"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "d"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
WshShell.SendKeys " "
WshShell.SendKeys "^ "
WScript.Sleep 150
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "z"
WScript.Sleep 1
WshShell.SendKeys "c"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys ","
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "z"
WScript.Sleep 1
WshShell.SendKeys ","
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "z"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "c"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "r"
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "p"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "T"
WScript.Sleep 1
WshShell.SendKeys "E"
WScript.Sleep 1
WshShell.SendKeys "X"
WScript.Sleep 1
WshShell.SendKeys "T"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
' 打开输入法
WshShell.SendKeys "^+"
WScript.Sleep 100
WshShell.SendKeys "z"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "g"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "w"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "d"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "z"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "j"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "g"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "d"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
WshShell.SendKeys " "
WshShell.SendKeys "^ "
WScript.Sleep 150
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys ","
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys ","
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "p"
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "w"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "c"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "y"
WScript.Sleep 1
WshShell.SendKeys "p"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "g"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "T"
WScript.Sleep 1
WshShell.SendKeys "E"
WScript.Sleep 1
WshShell.SendKeys "X"
WScript.Sleep 1
WshShell.SendKeys "T"
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
' 打开输入法
WshShell.SendKeys "^+"
WScript.Sleep 100
WshShell.SendKeys "d"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "k"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "r"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "f"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "d"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
WshShell.SendKeys " "
WshShell.SendKeys "^ "
WScript.Sleep 150
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "w"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys ","
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "w"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "p"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "d"
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "m"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "T"
WScript.Sleep 1
WshShell.SendKeys "E"
WScript.Sleep 1
WshShell.SendKeys "X"
WScript.Sleep 1
WshShell.SendKeys "T"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
' 打开输入法
WshShell.SendKeys "^+"
WScript.Sleep 100
WshShell.SendKeys "d"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "g"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "d"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "j"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "b"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "d"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "k"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "j"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys " "
WshShell.SendKeys "^ "
WScript.Sleep 150
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "c"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys ","
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "p"
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "c"
WScript.Sleep 1
WshShell.SendKeys "l"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "y"
WScript.Sleep 1
WshShell.SendKeys "p"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "g"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "T"
WScript.Sleep 1
WshShell.SendKeys "E"
WScript.Sleep 1
WshShell.SendKeys "X"
WScript.Sleep 1
WshShell.SendKeys "T"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
' 打开输入法
WshShell.SendKeys "^+"
WScript.Sleep 100
WshShell.SendKeys "g"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "b"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "r"
WScript.Sleep 1
WshShell.SendKeys "u"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "f"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "j"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys " "
WshShell.SendKeys "^ "
WScript.Sleep 150
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys ","
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "l"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "r"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "T"
WScript.Sleep 1
WshShell.SendKeys "E"
WScript.Sleep 1
WshShell.SendKeys "X"
WScript.Sleep 1
WshShell.SendKeys "T"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
' 打开输入法
WshShell.SendKeys "^+"
WScript.Sleep 100
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "j"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "g"
WScript.Sleep 1
' 优化中文输入
WshShell.SendKeys "g"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
WshShell.SendKeys " "
WshShell.SendKeys "^ "
WScript.Sleep 150
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "-"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "l"
WScript.Sleep 1
WshShell.SendKeys "p"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "S"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "o"
WScript.Sleep 1
WshShell.SendKeys "w"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "h"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "m"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "s"
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "g"
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "a"
WScript.Sleep 1
WshShell.SendKeys "n"
WScript.Sleep 1
WshShell.SendKeys "d"
WScript.Sleep 1
WshShell.SendKeys " "
WScript.Sleep 1
WshShell.SendKeys "e"
WScript.Sleep 1
WshShell.SendKeys "x"
WScript.Sleep 1
WshShell.SendKeys "i"
WScript.Sleep 1
WshShell.SendKeys "t"
WScript.Sleep 1
WshShell.SendKeys "."
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1
WshShell.SendKeys "`"
WScript.Sleep 1
WshShell.SendKeys "`"
WScript.Sleep 1
WshShell.SendKeys "`"
WScript.Sleep 1
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1