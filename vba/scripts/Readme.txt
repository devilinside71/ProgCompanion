felbontás elemekre
declaration
Dim xxx As String|Double


Formatter:
line break 80 char
1., REM line
(\S+)

breakpoints:
.,;+-*/&
substr


operator except "" and ():
(>|<|=|\+|-|&|\/)(?=(?=(?:[^"]*"[^"]*")*[^"]*$)(?![^\(]*\)))
new RegExp('(>|<|=|\\+|-|&|\\/)(?=(?=(?:[^"]*"[^"]*")*[^"]*$)(?![^\\(]*\\)))', 'gi');


