import re
# \d : 任何数字
# <_sre.SRE_Match object; span=(4, 7), match='r4n'>
print(re.search(r"r\dn", "run r4n"))
# \D : 不是数字
# <_sre.SRE_Match object; span=(0, 3), match='run'>
print(re.search(r"r\Dn", "run r4n"))
# \s : 任何 white space, 如 [\t\n\r\f\v]
# <_sre.SRE_Match object; span=(0, 3), match='r\nn'>
print(re.search(r"r\sn", "r\nn r4n"))
# \S : 不是 white space
# <_sre.SRE_Match object; span=(4, 7), match='r4n'>
print(re.search(r"r\Sn", "r\nn r4n"))
# \w : [a-zA-Z0-9_]
# <_sre.SRE_Match object; span=(4, 7), match='r4n'>
print(re.search(r"r\wn", "r\nn r4n"))
# \W : 不是 \w
# <_sre.SRE_Match object; span=(0, 3), match='r\nn'>
print(re.search(r"r\Wn", "r\nn r4n"))
# \b : 空白字符 (只在某个字的开头或结尾)
# <_sre.SRE_Match object; span=(4, 8), match='runs'>
print(re.search(r"\bruns\b", "dog runs to cat"))
# \B : 空白字符 (不在某个字的开头或结尾)
# <_sre.SRE_Match object; span=(8, 14), match=' runs '>
print(re.search(r"\B runs \B", "dog   runs  to cat"))
# \\ : 匹配 \
print(re.search(r"runs\\", "runs\ to me")
      )     # <_sre.SRE_Match object; span=(0, 5), match='runs\\'>
# . : 匹配任何字符 (除了 \n)
# <_sre.SRE_Match object; span=(0, 3), match='r[n'>
print(re.search(r"r.n", "r[ns to me"))
# ^ : 匹配开头
# <_sre.SRE_Match object; span=(0, 3), match='dog'>
print(re.search(r"^dog", "dog runs to cat"))
# $ : 匹配结尾
# <_sre.SRE_Match object; span=(12, 15), match='cat'>
print(re.search(r"cat$", "dog runs to cat"))
# ? : 前面的字符可有可无
# <_sre.SRE_Match object; span=(0, 6), match='Monday'>
print(re.search(r"Mon(day)?", "Monday"))
# <_sre.SRE_Match object; span=(0, 3), match='Mon'>
print(re.search(r"Mon(day)?", "Mon"))
