import winreg
"""
key = "SYSTEM\CurrentControlSet\Control\Print\Printers\Microsoft Print to PDF"
open_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key)
countkey = winreg.QueryInfoKey(open_key)[0]
keylist = []
for i in range(int(countkey)):
    name = winreg.EnumKey(open_key, i)  # 获取子键名
    keylist.append(name)
winreg.CloseKey(open_key)
print(keylist)
"""
import winreg

key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\Print\Printers\Microsoft Print to PDF\PrinterDriverData")

"""
# 获取该键的所有键值，因为没有方法可以获取键值的个数，所以只能用这种方法进行遍历
try:
    i = 0
    while 1:
        # EnumValue方法用来枚举键值，EnumKey用来枚举子键
        name, value, type = winreg.EnumValue(key, i)
    print(repr(name))
    i += 1
except WindowsError:
    print('no')

# 如果知道键的名称，也可以直接取值
"""
countkey = winreg.QueryInfoKey(key)[1]
print(winreg.QueryInfoKey(key))
keylist = []
valuelist = []
for i in range(int(countkey)):
    value = winreg.EnumValue(key, i)  # 获取数据，【0】：名称 【1】类型 【2】：数值
    print('名称：%s, 数据：%' %(value[0], value[1]))
    #keylist.append(value[0])
    #valuelist.append(value[2])
#print(keylist)


