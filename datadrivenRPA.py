import pyautogui
import time
import xlrd
import pyperclip
import os
import re
import subprocess

# ==================== 全局配置 ====================
CMD_IF_START = 91.0      # IF (存在则执行)
CMD_ELSE_START = 92.0    # ELSE (配合91使用)
CMD_IF_NOT_START = 94.0  # IF NOT (不存在则执行，独立使用)
CMD_IF_END = 93.0        # END (所有分支的结束)

MAX_RETRY_LIMIT = 50

# ==================== 数据驱动功能 ====================

def parseTemplate(template, data_dict):
    if not isinstance(template, str):
        return template
    pattern = r'\{\{(\w+)\}\}'
    def replace(match):
        col_name = match.group(1)
        if col_name in data_dict:
            return str(data_dict[col_name])
        return match.group(0)
    return re.sub(pattern, replace, template)

def getDataRow(data_sheet, row_idx):
    header_row = data_sheet.row(0)
    data_row = data_sheet.row(row_idx)
    data_dict = {}
    for col_idx, header_cell in enumerate(header_row):
        col_name = str(header_cell.value).strip()
        if col_idx < len(data_row):
            data_dict[col_name] = data_row[col_idx].value
        else:
            data_dict[col_name] = ""
    return data_dict

def processDataDriven(cmd_sheet, data_sheet):
    print("=" * 60)
    print("数据驱动模式启动 (支持 94-IF NOT)")
    print("=" * 60)
    headers = [str(cell.value).strip() for cell in data_sheet.row(0)]
    print(f"数据列：{headers}")
    print(f"数据行数：{data_sheet.nrows - 1}")
    print("=" * 60)
    
    for row_idx in range(1, data_sheet.nrows):
        data_row = getDataRow(data_sheet, row_idx)
        print(f"\n>>> 第 {row_idx} 轮执行")
        print("-" * 30)
        try:
            success = mainWork(cmd_sheet, data_row)
            if success:
                print(f">>> 第 {row_idx} 轮执行完成")
            else:
                print(f">>> 第 {row_idx} 轮被跳过/终止")
        except Exception as e:
            print(f"!!! 第 {row_idx} 轮发生异常：{e}")
            import traceback
            traceback.print_exc()
        print("-" * 30)
        time.sleep(1.0)
    print("\n所有数据执行完毕！")

# ==================== 功能函数 ====================

def mouseClick(clickTimes, lOrR, img, reTry, Checkeep):
    confidence_levels = [0.9, 0.8, 0.7] 
    if reTry == 1 and Checkeep == "keep":
        limit = MAX_RETRY_LIMIT
    elif reTry == 1 and Checkeep == "dontkeep":
        limit = 1
    elif reTry > 1:
        limit = reTry
    else:
        limit = 999999

    count = 0
    found = False
    
    while count < limit:
        for conf in confidence_levels:
            try:
                location = pyautogui.locateCenterOnScreen(img, confidence=conf)
                if location is not None:
                    pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
                    found = True
                    if conf < 0.9:
                        print(f"  [注意] 低置信度 ({conf}) 找到：{os.path.basename(img)}")
                    break
            except Exception:
                pass
        if found: break
        count += 1
        if count >= limit: break
        if Checkeep == "dontkeep": break
        if count % 10 == 0:
            print(f"  [等待] 寻找 {os.path.basename(img)}... ({count}/{limit})")
        time.sleep(0.2)

    if not found:
        if Checkeep == "keep":
            raise Exception(f"[错误] 超时未找到图片：{os.path.basename(img)}")
        else:
            print(f"  [跳过] 未找到 (dontkeep)：{os.path.basename(img)}")

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
        time.sleep(0.1)
    elif reTry > 1:
        for _ in range(reTry):
            hotkey_get(hkg_inputValue)
            time.sleep(0.1)

def checkImageExist(img, timeout=5):
    start_time = time.time()
    while time.time() - start_time < timeout:
        try:
            for conf in [0.9, 0.8]:
                if pyautogui.locateCenterOnScreen(img, confidence=conf):
                    return True
        except: pass
        time.sleep(0.2)
    return False

# ==================== 辅助函数 ====================

def findNextCommandIndex(sheet1, start_idx, target_type):
    """查找下一个特定命令的行号（处理嵌套）"""
    level = 1
    i = start_idx + 1
    while i < sheet1.nrows:
        val = sheet1.row(i)[0].value
        # 处理嵌套：91, 92, 94 都视为开始块（增加层级），93 视为结束
        if val in [CMD_IF_START, CMD_ELSE_START, CMD_IF_NOT_START]:
            level += 1
        elif val == CMD_IF_END:
            level -= 1
            if level == 0 and target_type == CMD_IF_END:
                return i
        # 特殊处理：如果在找 92，且层级为1
        if target_type == CMD_ELSE_START and val == CMD_ELSE_START and level == 2: 
            # 注意：这里逻辑稍微复杂，因为 92 是在 91 内部的。
            # 简单起见，对于 94 和 91 找 93，上面的逻辑足够。
            # 对于 91 找 92，我们需要在 level==1 (相对于91) 时找到 92。
            # 上面的 level 初始是 1。遇到 91 变 2。遇到 92 (如果是目标) 应该在 level==2 时返回？
            # 不，findNextCommandIndex 通常用于从 91 找 92。
            # 让我们简化：只处理 93 的查找，92 的查找单独处理或不依赖此通用函数的高级嵌套逻辑
            pass 
            
        # 修正通用查找逻辑：只严格匹配 target_type 且层级归零（针对93）或层级为1（针对92）
        if target_type == CMD_IF_END:
            if val == CMD_IF_END:
                level -= 1 # 先减再判断，因为进入时level已经加了？不，这里是扫描
                # 重新梳理 level 逻辑：
                # 起始 level = 1 (当前命令层级)
                # 遇到 开始命令 (91,92,94) -> level++
                # 遇到 结束命令 (93) -> level--
                # 当 level == 0 时，找到配对的 93
                pass
        
        # 为了代码稳健，重写一个简单的查找器
        i += 1
    
    # --- 重写查找逻辑以确保准确 ---
    level = 1
    i = start_idx + 1
    while i < sheet1.nrows:
        val = sheet1.row(i)[0].value
        
        if val in [CMD_IF_START, CMD_IF_NOT_START]:
            level += 1
        elif val == CMD_ELSE_START:
            # 92 比较特殊，它属于 91 的一部分，不增加新的嵌套层级用于寻找 93
            # 但在寻找 92 本身时，它应该是目标
            pass 
        elif val == CMD_IF_END:
            level -= 1
            if level == 0:
                if target_type == CMD_IF_END:
                    return i
        
        if target_type == CMD_ELSE_START and val == CMD_ELSE_START and level == 1:
             return i

        i += 1
    return -1

def dataCheck(sheet1, data_sheet=None):
    i = 1
    while i < sheet1.nrows:
        val = sheet1.row(i)[0].value
        if val in [CMD_IF_START, CMD_IF_NOT_START]:
            if findNextCommandIndex(sheet1, i, CMD_IF_END) == -1:
                cmd_name = "91 (IF)" if val == 91.0 else "94 (IF NOT)"
                print(f"[错误] 第 {i+1} 行：{cmd_name} 缺少配对的 93 (END)")
                return False
        elif val == CMD_ELSE_START:
            if findNextCommandIndex(sheet1, i, CMD_IF_END) == -1:
                print(f"[错误] 第 {i+1} 行：92 (ELSE) 缺少配对的 93 (END)")
                return False
        i += 1
    return True

# ==================== 主执行引擎 (支持 94) ====================

def mainWork(sheet1, data_row=None):
    i = 1
    n_rows = sheet1.nrows
    
    skip_until_row = -1
    wait_for_else_to_skip = False # 仅用于 91-92-93 逻辑

    while i < n_rows:
        row = sheet1.row(i)
        cmdType = row[0]
        cmdValue = row[1]
        cmdRetry = row[2]
        val = cmdType.value
        
        # --- 1. 通用跳过检查 ---
        if skip_until_row != -1:
            if i < skip_until_row:
                i += 1
                continue
            elif i == skip_until_row:
                print(f"  -> [恢复] 到达行 {i+1}，结束跳过")
                skip_until_row = -1
                wait_for_else_to_skip = False
        
        # --- 2. 91-92 特殊状态处理 ---
        if wait_for_else_to_skip and val == CMD_ELSE_START:
            print("  -> [92] IF 成功，跳过 ELSE 块")
            end_idx = findNextCommandIndex(sheet1, i, CMD_IF_END)
            if end_idx != -1:
                skip_until_row = end_idx
                wait_for_else_to_skip = False
                i += 1
                continue
            else:
                wait_for_else_to_skip = False

        # --- 3. 控制命令处理 ---
        
        # === 91.0 IF (存在则执行) ===
        if val == CMD_IF_START:
            img = cmdValue.value
            if data_row: img = parseTemplate(img, data_row)
            timeout = cmdRetry.value if cmdRetry.ctype == 2 and cmdRetry.value != 0 else 5
            found = checkImageExist(img, timeout)
            
            if found:
                print(f"  -> [91] 条件满足 ({os.path.basename(img)})，执行 THEN 块")
                else_idx = findNextCommandIndex(sheet1, i, CMD_ELSE_START)
                if else_idx != -1:
                    wait_for_else_to_skip = True
            else:
                print(f"  -> [91] 条件不满足 ({os.path.basename(img)})，跳过 THEN 块")
                else_idx = findNextCommandIndex(sheet1, i, CMD_ELSE_START)
                if else_idx != -1:
                    skip_until_row = else_idx
                else:
                    end_idx = findNextCommandIndex(sheet1, i, CMD_IF_END)
                    if end_idx != -1:
                        skip_until_row = end_idx
            i += 1
            continue

        # === 94.0 IF NOT (不存在则执行) ===
        elif val == CMD_IF_NOT_START:
            img = cmdValue.value
            if data_row: img = parseTemplate(img, data_row)
            timeout = cmdRetry.value if cmdRetry.ctype == 2 and cmdRetry.value != 0 else 5
            found = checkImageExist(img, timeout)
            
            if not found:
                # 图片不存在 -> 条件成立 -> 执行块内代码
                print(f"  -> [94] 条件满足 (未发现 {os.path.basename(img)})，执行块内逻辑")
                # 什么都不做，继续向下执行
            else:
                # 图片存在 -> 条件不成立 -> 跳过块内代码
                print(f"  -> [94] 条件不满足 (已发现 {os.path.basename(img)})，跳过块内逻辑")
                end_idx = findNextCommandIndex(sheet1, i, CMD_IF_END)
                if end_idx != -1:
                    skip_until_row = end_idx
                    print(f"  -> [设置] 跳过至 93 (行 {end_idx+1})")
                else:
                    print("  -> [警告] 未找到配对的 93")
            i += 1
            continue

        # === 92.0 ELSE ===
        elif val == CMD_ELSE_START:
            # 仅作为 91 的跳转点，独立运行时忽略或视为普通行（这里选择忽略继续）
            if skip_until_row == -1 and not wait_for_else_to_skip:
                # 如果没有前序 91，直接当作普通标记跳过，或者执行？
                # 按照定义，92 必须配合 91。孤立 92 直接跳过内容直到 93 比较安全
                # 但为了灵活性，如果孤立出现，我们视为“执行”直到 93
                pass
            i += 1
            continue

        # === 93.0 END ===
        elif val == CMD_IF_END:
            if skip_until_row != -1 or wait_for_else_to_skip:
                print("  -> [93] 到达结束点，重置状态")
            skip_until_row = -1
            wait_for_else_to_skip = False
            i += 1
            continue

        # --- 4. 执行普通命令 ---
        if skip_until_row != -1 and i < skip_until_row:
            i += 1
            continue

        try:
            if val == 1.0:
                img = cmdValue.value
                reTry = cmdRetry.value if cmdRetry.ctype == 2 and cmdRetry.value != 0 else 1
                mouseClick(1, "left", img, reTry, "keep")
                print("  [OK] 左键", os.path.basename(img))
            elif val == 1.1:
                img = cmdValue.value
                mouseClick(1, "left", img, 1, "dontkeep")
                print("  [OK] 尝试单击", os.path.basename(img))
            elif val == 2.0:
                img = cmdValue.value
                reTry = cmdRetry.value if cmdRetry.ctype == 2 and cmdRetry.value != 0 else 1
                mouseClick(2, "left", img, reTry, "keep")
                print("  [OK] 双击", os.path.basename(img))
            elif val == 3.0:
                img = cmdValue.value
                reTry = cmdRetry.value if cmdRetry.ctype == 2 and cmdRetry.value != 0 else 1
                mouseClick(1, "right", img, reTry, "keep")
                print("  [OK] 右键", os.path.basename(img))
            elif val == 4.0:
                txt = cmdValue.value
                if data_row: txt = parseTemplate(txt, data_row)
                pyperclip.copy(txt)
                pyautogui.hotkey('ctrl', 'v')
                print(f"  [OK] 输入：{txt[:20]}")
            elif val == 5.0:
                wt = cmdValue.value
                time.sleep(wt)
                print(f"  [OK] 等待 {wt} 秒")
            elif val == 6.0:
                sc = cmdValue.value
                pyautogui.scroll(int(sc))
            elif val == 7.0:
                txt = cmdValue.value
                if data_row: txt = parseTemplate(txt, data_row)
                reTry = cmdRetry.value if cmdRetry.ctype == 2 and cmdRetry.value != 0 else 1
                hotkeyGroup(reTry, txt)
                print(f"  [OK] 热键：{txt}")
            elif val == 8.0:
                lt = time.strftime("%Y-%m-%d %H:%M:%S")
                pyperclip.copy(lt)
                pyautogui.hotkey('ctrl', 'v')
            elif val == 9.0:
                wincmd = cmdValue.value
                if data_row: wincmd = parseTemplate(wincmd, data_row)
                print(f"  [CMD] 执行：{wincmd}")
                subprocess.Popen(wincmd, shell=True)
                time.sleep(0.5)
            elif val == 10.0:
                txt = cmdValue.value
                if data_row: txt = parseTemplate(txt, data_row)
                pyperclip.copy(txt)
                pyautogui.hotkey('ctrl', 'v')
                print(f"  [OK] 输入：{txt[:20]}")
            elif val == 11.0:
                img = cmdValue.value
                if data_row: img = parseTemplate(img, data_row)
                timeout = cmdRetry.value if cmdRetry.ctype == 2 and cmdRetry.value != 0 else 5
                if not checkImageExist(img, timeout):
                    print(f"  [FAIL] 条件判断失败：{img}")
                    return False
        except Exception as e:
            print(f"\n!!! 命令执行失败 (行 {i+1}): {e}")
            raise
        
        time.sleep(0.5)
        i += 1

    return True

if __name__ == '__main__':
    file = r'cmd.xls'
    if not os.path.exists(file):
        if os.path.exists('cmd.xlsx'):
            print("错误：请将 cmd.xlsx 另存为 cmd.xls")
        else:
            print(f"错误：找不到 {file}")
        os.system('pause')
        exit()

    wb = xlrd.open_workbook(filename=file)
    cmd_sheet = wb.sheet_by_index(0)
    # 自动查找 data sheet
    data_sheet = None
    for name in wb.sheet_names():
        if 'data' in name.lower():
            data_sheet = wb.sheet_by_name(name)
            break
    
    print('运行中...')
    if not dataCheck(cmd_sheet, data_sheet):
        print("检查失败")
        os.system('pause')
        exit()

    try:
        if data_sheet:
            processDataDriven(cmd_sheet, data_sheet)
        else:
            mainWork(cmd_sheet)
    except KeyboardInterrupt:
        print("\n用户停止")
    except Exception as e:
        print(f"\n程序遇到错误...")
        import traceback
        traceback.print_exc()
    
    os.system('pause')
