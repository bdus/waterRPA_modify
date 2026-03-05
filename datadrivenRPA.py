import pyautogui
import time
import xlrd
import pyperclip
import os
import re


# ==================== 数据驱动功能 ====================

def parseTemplate(template, data_dict):
    """
    将模板字符串中的 {{column_name}} 替换为 data_dict 中对应的值
    """
    if not isinstance(template, str):
        return template

    pattern = r'\{\{(\w+)\}\}'

    def replace(match):
        col_name = match.group(1)
        if col_name in data_dict:
            return str(data_dict[col_name])
        return match.group(0)  # 如果找不到列，保持原样

    return re.sub(pattern, replace, template)


def getDataRow(data_sheet, row_idx):
    """
    从 data sheet 中获取一行数据，返回字典格式
    row_idx: 行索引（从 1 开始，0 是表头）
    """
    header_row = data_sheet.row(0)
    data_row = data_sheet.row(row_idx)

    data_dict = {}
    for col_idx, header_cell in enumerate(header_row):
        if col_idx < len(data_row):
            col_name = str(header_cell.value).strip()
            data_dict[col_name] = data_row[col_idx].value
        else:
            col_name = str(header_cell.value).strip()
            data_dict[col_name] = ""

    return data_dict


def processDataDriven(cmd_sheet, data_sheet):
    """
    数据驱动模式：对 data sheet 中的每一行数据执行 cmd 中的所有命令
    """
    print("=" * 50)
    print("数据驱动模式启动")
    print("=" * 50)

    # 获取表头
    headers = [str(cell.value).strip() for cell in data_sheet.row(0)]
    print(f"数据列：{headers}")
    print(f"数据行数：{data_sheet.nrows - 1}")
    print("=" * 50)

    # 遍历每一行数据（跳过表头）
    for row_idx in range(1, data_sheet.nrows):
        data_row = getDataRow(data_sheet, row_idx)
        print(f"\n>>> 第 {row_idx} 轮执行，数据：{data_row}")
        print("-" * 50)

        success = mainWork(cmd_sheet, data_row)

        if success:
            print(f">>> 第 {row_idx} 轮执行完成")
        else:
            print(f">>> 第 {row_idx} 轮被跳过（条件判断失败）")

        print("-" * 50)
        time.sleep(0.5)  # 每轮之间稍作等待

    print("\n" + "=" * 50)
    print("所有数据执行完毕！")
    print("=" * 50)


# ==================== 原有功能 ====================

def mouseClick(clickTimes, lOrR, img, reTry, Checkeep):
    if reTry == 1:
        while True:
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
                break
            if Checkeep == "dontkeep":
                break
            print("未找到匹配图片，0.1 秒后重试")
            time.sleep(0.1)
    elif reTry == -1:
        while True:
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
            time.sleep(0.1)
    elif reTry > 1:
        i = 1
        while i < reTry + 1:
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
                print("重复第", i, "次")
                i += 1
                time.sleep(0.1)
                break
            if Checkeep == "dontkeep":
                break
            print("未找到匹配图片，0.1 秒后重试")
            time.sleep(0.1)


def hotkey_get(hk_g_inputValue):
    try:
        newinput = hk_g_inputValue.split(',')
        pyautogui.hotkey(*tuple(newinput))
    except:
        pyperclip.copy(hk_g_inputValue)
        pyautogui.hotkey('ctrl', 'v')


def hotkeyGroup(reTry, hkg_inputValue):
    if reTry == 1:
        hotkey_get(hkg_inputValue)
        print("执行了：", hkg_inputValue)
        time.sleep(0.1)
    elif reTry == -1:
        while True:
            hotkey_get(hkg_inputValue)
            print("执行了：", hkg_inputValue)
            time.sleep(0.1)
    elif reTry > 1:
        i = 1
        while i < reTry + 1:
            hotkey_get(hkg_inputValue)
            print("执行了：", hkg_inputValue)
            i += 1
            time.sleep(0.1)


def checkImageExist(img, timeout=5):
    """
    检查图片是否在屏幕上存在
    timeout: 最大等待时间（秒）
    返回：True/False
    """
    start_time = time.time()
    while time.time() - start_time < timeout:
        try:
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if location is not None:
                return True
        except pyautogui.ImageNotFoundException:
            pass  # 没找到图片，继续重试
        time.sleep(0.2)
    return False


# 数据检查
# cmdType.value  1.0 左键单击    2.0 左键双击  3.0 右键单击  4.0 输入  5.0 等待  6.0 滚轮
# 7.0 热键组合（最多 4 个）
# 8.0 粘贴当前时间
# 9.0 系统命令集
# 10.0 输出数据列文本（数据驱动模式专用）
# 11.0 条件判断（图像存在性检查）
# ctype     空：0
#           字符串：1
#           数字：2
#           日期：3
#           布尔：4
#           error：5
def dataCheck(sheet1, data_sheet=None):
    checkCmd = True
    errorCount = 0

    # 行数检查
    if sheet1.nrows < 2:
        print("[错误] cmd sheet 没有数据行（至少需要 1 行表头 +1 行数据）")
        return False

    # 有效的命令类型列表
    validCmdTypes = [1.0, 1.1, 2.0, 2.1, 3.0, 3.1, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0, 11.0]

    # 每行数据检查
    i = 1
    while i < sheet1.nrows:
        # 第 1 列 操作类型检查
        cmdType = sheet1.row(i)[0]
        cmdValue = sheet1.row(i)[1]

        if cmdType.ctype != 2 or cmdType.value not in validCmdTypes:
            print(f"[错误] 第 {i+1} 行：操作类型无效，当前值='{cmdType.value}'")
            errorCount += 1
            checkCmd = False

        # 读图点击类型指令，内容必须为字符串类型 (1.0, 2.0, 3.0, 11.0)
        elif cmdType.value in [1.0, 2.0, 3.0, 11.0]:
            if cmdValue.ctype != 1:
                print(f"[错误] 第 {i+1} 行：图片路径必须为字符串，当前值='{cmdValue.value}'")
                errorCount += 1
                checkCmd = False

        # 输入类型，内容不能为空 (数据驱动模式下可以是 {{column_name}} 格式)
        elif cmdType.value == 4.0 or cmdType.value == 10.0:
            if cmdValue.ctype == 0:
                print(f"[错误] 第 {i+1} 行：输入文本不能为空")
                errorCount += 1
                checkCmd = False

        # 等待类型，内容必须为数字
        elif cmdType.value == 5.0:
            if cmdValue.ctype != 2:
                print(f"[错误] 第 {i+1} 行：等待时间必须为数字，当前值='{cmdValue.value}'")
                errorCount += 1
                checkCmd = False

        # 滚轮事件，内容必须为数字
        elif cmdType.value == 6.0:
            if cmdValue.ctype != 2:
                print(f"[错误] 第 {i+1} 行：滚轮滑动必须为数字，当前值='{cmdValue.value}'")
                errorCount += 1
                checkCmd = False

        # 7.0 热键组合，内容不能为空
        elif cmdType.value == 7.0:
            if cmdValue.ctype == 0:
                print(f"[错误] 第 {i+1} 行：热键组合不能为空")
                errorCount += 1
                checkCmd = False

        # 8.0 时间，内容不能为空
        elif cmdType.value == 8.0:
            if cmdValue.ctype == 0:
                print(f"[错误] 第 {i+1} 行：时间格式不能为空")
                errorCount += 1
                checkCmd = False

        # 9.0 系统命令集模式，内容不能为空
        elif cmdType.value == 9.0:
            if cmdValue.ctype == 0:
                print(f"[错误] 第 {i+1} 行：系统命令不能为空")
                errorCount += 1
                checkCmd = False

        i += 1

    # 打印检查总结
    if errorCount > 0:
        print(f"\n[总结] 数据检查失败，共发现 {errorCount} 个错误，请修复后重试！")
    else:
        print(f"\n[总结] 数据检查通过，共检查 {sheet1.nrows-1} 行数据，未发现错误。")

    # 检查 data sheet（如果存在）
    if data_sheet is not None:
        if data_sheet.nrows < 2:
            print("[错误] data sheet 没有数据行（至少需要 1 行表头 +1 行数据）")
            checkCmd = False
        else:
            # 检查表头
            headers = data_sheet.row(0)
            if len(headers) == 0:
                print("[错误] data sheet 表头为空")
                checkCmd = False
            else:
                headerNames = [str(h.value).strip() for h in headers]
                print(f"data sheet 表头列名：{headerNames}")
                print(f"data sheet 数据行数：{data_sheet.nrows - 1}")

    return checkCmd


# 任务
def mainWork(sheet1, data_row=None):
    """
    执行 cmd sheet 中的所有命令
    data_row: 数据驱动模式下的当前行数据字典，如果为 None 则是固定模式
    返回：True-执行成功，False-需要跳过当前行（条件判断失败）
    """
    i = 1
    while i < sheet1.nrows:
        # 取本行指令的操作类型
        cmdType = sheet1.row(i)[0]

        # 1 代表点击左键
        if cmdType.value == 1.0:
            # 取图片名称
            img = sheet1.row(i)[1].value
            # 取重试次数
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(1, "left", img, reTry, "keep")
            print("左键", img)

        # 1.1 代表点击左键【尝试 1 次若无图片则跳过此命令】
        elif cmdType.value == 1.1:
            # 取图片名称
            img = sheet1.row(i)[1].value
            # 取重试次数
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(1, "left", img, reTry, "dontkeep")
            print("单击", img, "  ", "【此步仅为尝试】")

        # 2 代表双击左键
        elif cmdType.value == 2.0:
            # 取图片名称
            img = sheet1.row(i)[1].value
            # 取重试次数
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(2, "left", img, reTry, "keep")
            print("双击", img)

        # 2.1 代表双击左键【尝试 1 次若无图片则跳过此命令】
        elif cmdType.value == 2.1:
            # 取图片名称
            img = sheet1.row(i)[1].value
            # 取重试次数
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(2, "left", img, reTry, "dontkeep")
            print("双击", img, "  ", "【此步仅为尝试】")

        # 3 代表右键
        elif cmdType.value == 3.0:
            # 取图片名称
            img = sheet1.row(i)[1].value
            # 取重试次数
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(1, "right", img, reTry, "keep")
            print("右键", img)

        # 3.1 代表右键【尝试几次若无图片则跳过此命令】
        elif cmdType.value == 3.1:
            # 取图片名称
            img = sheet1.row(i)[1].value
            # 取重试次数
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(1, "right", img, reTry, "dontkeep")
            print("右键", img, "  ", "【此步仅为尝试】")

        # 4 代表输入
        elif cmdType.value == 4.0:
            inputValue = sheet1.row(i)[1].value
            # 数据驱动模式下，解析模板
            if data_row is not None:
                inputValue = parseTemplate(inputValue, data_row)
            pyperclip.copy(inputValue)
            pyautogui.hotkey('ctrl', 'v')
            print("输入:", inputValue)
            time.sleep(0.5)

        # 5 代表等待
        elif cmdType.value == 5.0:
            # 取等待时间
            waitTime = sheet1.row(i)[1].value
            time.sleep(waitTime)
            print("等待", waitTime, "秒")

        # 6 代表滚轮
        elif cmdType.value == 6.0:
            # 取滚动距离
            scroll = sheet1.row(i)[1].value
            pyautogui.scroll(int(scroll))
            print("滚轮滑动", int(scroll), "距离")

        # 7 代表热键组合
        elif cmdType.value == 7.0:
            # 取重试次数，并循环。
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            inputValue = sheet1.row(i)[1].value
            # 数据驱动模式下，解析模板
            if data_row is not None:
                inputValue = parseTemplate(inputValue, data_row)
            hotkeyGroup(reTry, inputValue)
            time.sleep(0.5)

        # 8 代表粘贴当前时间
        elif cmdType.value == 8.0:
            # 设置本机当前时间。
            localtime = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
            pyperclip.copy(localtime)
            pyautogui.hotkey('ctrl', 'v')
            print("粘贴了本机时间:", localtime)
            time.sleep(0.5)

        # 9 代表系统命令集模式
        elif cmdType.value == 9.0:
            wincmd = sheet1.row(i)[1].value
            os.system(wincmd)
            print("运行系统命令:", wincmd)
            time.sleep(0.5)

        # 10.0 代表输出数据列文本（新增）
        elif cmdType.value == 10.0:
            template = sheet1.row(i)[1].value
            if data_row is not None:
                outputValue = parseTemplate(template, data_row)
                pyperclip.copy(outputValue)
                pyautogui.hotkey('ctrl', 'v')
                print("输出数据:", outputValue)
            else:
                print("警告：10.0 命令只能在数据驱动模式下使用")
            time.sleep(0.5)

        # 11.0 代表条件判断（图像存在性检查）
        elif cmdType.value == 11.0:
            img = sheet1.row(i)[1].value
            # 数据驱动模式下，解析模板（支持{{column_name}}）
            if data_row is not None:
                img = parseTemplate(img, data_row)

            # 取超时时间
            timeout = 5  # 默认超时 5 秒
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                timeout = sheet1.row(i)[2].value

            found = checkImageExist(img, timeout)
            if not found:
                print(f"条件判断失败：图片 {img} 不存在，跳过当前数据行")
                return False  # 返回 False 表示需要跳过
            else:
                print(f"条件判断通过：图片 {img} 存在")

        i += 1

    return True  # 执行成功


# 主程序
if __name__ == '__main__':
    file = r'cmd.xls'
    # 打开文件
    wb = xlrd.open_workbook(filename=file)
    # 通过索引获取表格 sheet 页
    cmd_sheet = wb.sheet_by_index(0)

    # 检查是否有 data sheet
    data_sheet = None
    if wb.nsheets > 1:
        data_sheet = wb.sheet_by_index(1)
        print("检测到 data sheet，启用数据驱动模式")
    else:
        print("只检测到 cmd sheet，使用固定模式")

    print('数据驱动增强版 v240224')
    print('')

    # 避免多次循环导致的 ctrl+v 导入到，按 ESC 进行取消。
    pyautogui.hotkey('esc')

    # 数据检查
    checkCmd = dataCheck(cmd_sheet, data_sheet)

    # 输入选项实现功能
    if checkCmd:
        key = "1"
        # input('选择功能：1.做一次 2.循环几次 3.循环到死 0.退出程序\n特殊功能：c.清理屏幕显示\n———————————————————————————————————————\n')
        if key == '1':
            # 循环拿出每一行指令
            print("正在执行第 1 次命令")
            print("")
            if data_sheet is not None:
                # 数据驱动模式
                processDataDriven(cmd_sheet, data_sheet)
            else:
                # 固定模式
                mainWork(cmd_sheet)
            print("")
            print("已经完成第 1 次命令")
            print("——————————————————分割线——————————————————")
            print("")

        elif key == '2':
            print("")
            count = 0
            times = input('输入需要循环的次数，务必输入正整数。\n')
            times = int(times)
            if count < times:
                while count < times:
                    count += 1
                    print("正在执行第", count, "次", "命令")
                    print("")
                    if data_sheet is not None:
                        processDataDriven(cmd_sheet, data_sheet)
                    else:
                        mainWork(cmd_sheet)
                    time.sleep(0.1)
                    print("等待 0.1 秒")
                    print("")
                    print("已经完成第", count, "次", "命令")
                    print("——————————————————分割线——————————————————")
                    print("")
            else:
                print('输入有误或者已经退出!')
                os.system('pause')
                print("")
                print("——————————————————————————————————————————")

        elif key == '3':
            count = 0
            while True:
                count += 1
                print("正在执行第", count, "次", "命令")
                print("")
                if data_sheet is not None:
                    processDataDriven(cmd_sheet, data_sheet)
                else:
                    mainWork(cmd_sheet)
                time.sleep(0.1)
                print("等待 0.1 秒")
                print("")
                print("已经完成第", count, "次", "命令")
                print("——————————————————分割线——————————————————")
                print("")

        elif key == '0':
            print("正在清理缓存文件...")
            os.system('@echo off & for /d %i in (%temp%\\^_MEI*) do (rd /s /q "%i")>nul')
            exit("正在退出程序...")

        elif key == 'c':
            os.system('cls')

        else:
            print('输入有误或者已经退出!')
            os.system('pause')
            print("")
            print("——————————————————————————————————————————")

    os.system('pause')
