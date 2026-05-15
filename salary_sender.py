import os
import pandas as pd
import requests
import json
import logging
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Border, Side
from copy import copy

# ===================== 【用户配置区】这里改成你自己的信息 =====================
# 企业微信核心配置
CORP_ID = "wwd5993be58eb1ecb3"
CORP_SECRET = "4NBtL1Oq_5iqF9fCGAbIJ6f78R5-posiAUjC_21GHCI"
AGENT_ID = 1000092  # 数字类型，比如1000001

# 工资表配置
TOTAL_SALARY_FILE = "f:\\项目\\salary sheet\\总工资表.xlsx"  # 您上传的总工资表文件路径
SAVE_DIR = "f:\\项目\\salary sheet\\工资条拆分文件"  # 单人工资条保存目录
LOG_FILE = "f:\\项目\\salary sheet\\工资条发送日志.txt"  # 发送日志文件路径

# 功能开关
RESPLIT_SALARY = False  # 是否重新拆分工资条（True=重新生成，False=用已有的）
RESEND_ALL = True  # 是否重新发送所有（True=全部重发，False=只发失败的）
SEND_MODE = "file"  # 发送模式："file"=发Excel文件，"text"=发文字工资条
MAX_RETRIES = 3  # 发送失败最大重试次数
USE_LOCAL_USER_MAP = True  # 是否使用本地用户映射表（True=从Excel读取，False=从API获取）
USER_MAP_FILE = "f:\\项目\\salary sheet\\总部人员信息表.xlsx"  # 本地员工信息表路径
# ==================================================================================

# 初始化日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler()
    ]
)

# 创建保存目录
os.makedirs(SAVE_DIR, exist_ok=True)

# 定义单元格样式（和原表保持一致）
thin_border = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)
center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

# ===================== 企业微信API核心函数 =====================
def get_access_token():
    """获取企业微信access_token"""
    url = f"https://qyapi.weixin.qq.com/cgi-bin/gettoken?corpid={CORP_ID}&corpsecret={CORP_SECRET}"
    try:
        resp = requests.get(url, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        if data.get("errcode") != 0:
            raise Exception(f"获取access_token失败: {data.get('errmsg')}")
        logging.info("✅ 获取企业微信access_token成功")
        return data["access_token"]
    except Exception as e:
        logging.error(f"❌ 获取access_token失败: {str(e)}")
        raise

def get_user_list(access_token):
    """获取企业微信通讯录用户列表，返回姓名:userid的映射"""
    url = f"https://qyapi.weixin.qq.com/cgi-bin/user/list?access_token={access_token}&department_id=1&fetch_child=1"
    try:
        resp = requests.get(url, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        if data.get("errcode") != 0:
            raise Exception(f"获取用户列表失败: {data.get('errmsg')}")
        user_map = {user["name"]: user["userid"] for user in data["userlist"]}
        logging.info(f"✅ 获取到企业微信用户总数: {len(user_map)}")
        return user_map
    except Exception as e:
        logging.error(f"❌ 获取用户列表失败: {str(e)}")
        raise

def get_user_list_from_local():
    """从本地Excel文件读取员工姓名和userid映射"""
    try:
        df = pd.read_excel(USER_MAP_FILE, header=None)
        user_map = {}
        for _, row in df.iterrows():
            if len(row) >= 2 and pd.notna(row[0]) and pd.notna(row[1]):
                name = str(row[0]).strip()
                # 从姓名中提取纯姓名（格式：部门-姓名 或 姓名）
                if '-' in name:
                    name_parts = name.split('-')
                    if len(name_parts) > 1:
                        name = name_parts[-1].strip()
                userid = str(row[1]).strip()
                user_map[name] = userid
        logging.info(f"✅ 从本地文件读取到 {len(user_map)} 个用户映射")
        return user_map
    except Exception as e:
        logging.error(f"❌ 从本地文件读取用户列表失败: {str(e)}")
        raise

def upload_media(access_token, file_path):
    """上传临时文件到企业微信，返回media_id"""
    url = f"https://qyapi.weixin.qq.com/cgi-bin/media/upload?access_token={access_token}&type=file"
    try:
        with open(file_path, "rb") as f:
            files = {"media": (os.path.basename(file_path), f, "application/octet-stream")}
            resp = requests.post(url, files=files, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        if data.get("errcode") != 0:
            raise Exception(f"上传文件失败: {data.get('errmsg')}")
        return data["media_id"]
    except Exception as e:
        logging.error(f"❌ 上传文件失败: {str(e)}")
        raise

def send_file_msg(access_token, userid, media_id, name):
    """发送企业微信文件消息"""
    url = f"https://qyapi.weixin.qq.com/cgi-bin/message/send?access_token={access_token}"
    payload = {
        "touser": userid,
        "msgtype": "file",
        "agentid": AGENT_ID,
        "file": {"media_id": media_id},
        "safe": 0
    }
    try:
        resp = requests.post(url, json=payload, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        if data.get("errcode") != 0:
            raise Exception(f"发送消息失败: {data.get('errmsg')}")
        logging.info(f"✅ {name} 工资条文件发送成功")
        return True
    except Exception as e:
        logging.error(f"❌ {name} 发送失败: {str(e)}")
        return False

def send_text_msg(access_token, userid, row_data, name):
    """发送企业微信结构化文字工资条"""
    url = f"https://qyapi.weixin.qq.com/cgi-bin/message/send?access_token={access_token}"
    
    # 提取核心工资字段，适配您的表格结构
    dept = row_data[1] if pd.notna(row_data[1]) else ""
    post = row_data[2] if pd.notna(row_data[2]) else ""
    real_attend = row_data[6] if pd.notna(row_data[6]) else ""
    base_salary = row_data[7] if pd.notna(row_data[7]) else ""
    post_salary = row_data[8] if pd.notna(row_data[8]) else ""
    performance = row_data[9] if pd.notna(row_data[9]) else ""
    overtime = row_data[10] if pd.notna(row_data[10]) else ""
    house_subsidy = row_data[11] if pd.notna(row_data[11]) else ""
    full_attend = row_data[12] if pd.notna(row_data[12]) else ""
    seniority = row_data[20] if pd.notna(row_data[20]) else ""
    should_pay = row_data[21] if pd.notna(row_data[21]) else ""
    social_security = row_data[22] if pd.notna(row_data[22]) else ""
    should_deduct = row_data[23] if pd.notna(row_data[23]) else ""
    real_pay = row_data[24] if pd.notna(row_data[24]) else ""

    # 生成文字内容
    month = datetime.now().strftime("%Y年%m月")
    text_content = f"""【{month} 工资条 - {name}】
部门：{dept}
岗位：{post}
实出勤：{real_attend} 天
-------------------
【工资明细】
基本工资：{base_salary} 元
岗位薪资：{post_salary} 元
绩效工资：{performance} 元
加班工资：{overtime} 元
房补：{house_subsidy} 元
全勤奖：{full_attend} 元
工龄补贴：{seniority} 元
-------------------
应发合计：{should_pay} 元
社保扣款：{social_security} 元
应扣合计：{should_deduct} 元
本月实发：{real_pay} 元
-------------------
如有疑问，请联系财务核对~"""

    payload = {
        "touser": userid,
        "msgtype": "text",
        "agentid": AGENT_ID,
        "text": {"content": text_content},
        "safe": 0
    }
    try:
        resp = requests.post(url, json=payload, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        if data.get("errcode") != 0:
            raise Exception(f"发送消息失败: {data.get('errmsg')}")
        logging.info(f"✅ {name} 文字工资条发送成功")
        return True
    except Exception as e:
        logging.error(f"❌ {name} 发送失败: {str(e)}")
        return False

# ===================== 工资条拆分核心函数 =====================
def split_all_salary():
    """批量拆分所有员工的工资条，和原表格式完全一致"""
    # 读取原表
    wb_original = load_workbook(TOTAL_SALARY_FILE)
    ws_original = wb_original["工资表"]

    # 提取双行表头
    header_row1 = [cell.value for cell in ws_original[1]]
    header_row2 = [cell.value for cell in ws_original[2]]

    # 提取所有有效数据行
    data_rows = []
    for row in ws_original.iter_rows(min_row=3, values_only=True):
        data_rows.append(row)

    # 提取姓名列索引
    name_col_idx = None
    for idx, cell in enumerate(header_row1):
        if cell == "姓名":
            name_col_idx = idx
            break
    if name_col_idx is None:
        for idx, cell in enumerate(header_row2):
            if cell == "姓名":
                name_col_idx = idx
                break

    if name_col_idx is None:
        raise Exception("❌ 未找到姓名列，请检查工资表结构")

    # 循环拆分每个员工
    success_count = 0
    for row in data_rows:
        name = row[name_col_idx]
        if pd.isna(name):
            logging.warning("⚠️ 跳过空姓名行")
            continue

        # 创建新工作簿
        wb_new = Workbook()
        ws_new = wb_new.active
        ws_new.title = f"{name}工资条"

        # 写入双行表头和数据
        ws_new.append(header_row1)
        ws_new.append(header_row2)
        ws_new.append(row)

        # 应用样式
        for row_cells in ws_new.iter_rows():
            for cell in row_cells:
                cell.border = thin_border
                cell.alignment = center_alignment

        # 自动调整列宽
        for col in ws_new.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length * 1.2 + 2, 20)
            ws_new.column_dimensions[column].width = adjusted_width

        # 保存文件
        month = datetime.now().strftime("%Y%m")
        file_name = f"工资条-{name}-{month}.xlsx"
        file_path = os.path.join(SAVE_DIR, file_name)
        wb_new.save(file_path)
        success_count += 1
        logging.info(f"📄 生成工资条：{file_path}")

    logging.info(f"✅ 工资条拆分完成，共生成 {success_count} 个文件")
    return data_rows, name_col_idx

# ===================== 主函数 =====================
def main():
    try:
        # 1. 拆分工资条
        if RESPLIT_SALARY:
            data_rows, name_col_idx = split_all_salary()
        else:
            # 读取已有的数据
            wb_original = load_workbook(TOTAL_SALARY_FILE)
            ws_original = wb_original["工资表"]
            data_rows = []
            for row in ws_original.iter_rows(min_row=3, values_only=True):
                data_rows.append(row)
            # 查找姓名列索引
            name_col_idx = None
            for idx, cell in enumerate([cell.value for cell in ws_original[1]]):
                if cell == "姓名":
                    name_col_idx = idx
                    break
            if name_col_idx is None:
                for idx, cell in enumerate([cell.value for cell in ws_original[2]]):
                    if cell == "姓名":
                        name_col_idx = idx
                        break
            logging.info("✅ 跳过拆分，使用已有的工资条文件")

        # 2. 获取企业微信信息和用户列表
        access_token = get_access_token()
        if USE_LOCAL_USER_MAP:
            user_map = get_user_list_from_local()
        else:
            user_map = get_user_list(access_token)

        # 3. 循环发送工资条
        success_send = 0
        fail_send = 0
        skip_send = 0
        month = datetime.now().strftime("%Y%m")

        for row in data_rows:
            name = row[name_col_idx]
            if pd.isna(name):
                skip_send += 1
                continue

            # 匹配企业微信用户
            if name not in user_map:
                logging.warning(f"⚠️ 企业微信中未找到用户：{name}，跳过发送")
                skip_send += 1
                continue
            userid = user_map[name]

            # 发送模式选择
            send_success = False
            if SEND_MODE == "file":
                # 发送文件模式
                file_name = f"工资条-{name}-{month}.xlsx"
                file_path = os.path.join(SAVE_DIR, file_name)
                if not os.path.exists(file_path):
                    logging.warning(f"⚠️ 未找到{name}的工资条文件，跳过")
                    skip_send += 1
                    continue

                # 上传文件并发送，带重试
                for i in range(MAX_RETRIES):
                    try:
                        media_id = upload_media(access_token, file_path)
                        send_success = send_file_msg(access_token, userid, media_id, name)
                        if send_success:
                            break
                    except Exception as e:
                        logging.warning(f"第{i+1}次发送失败，重试中：{str(e)}")

            elif SEND_MODE == "text":
                # 发送文字模式，带重试
                for i in range(MAX_RETRIES):
                    send_success = send_text_msg(access_token, userid, row, name)
                    if send_success:
                        break

            # 统计结果
            if send_success:
                success_send += 1
            else:
                fail_send += 1
                logging.error(f"❌ {name} 最终发送失败，已达最大重试次数")

        # 4. 输出最终结果
        logging.info("="*50)
        logging.info("📊 工资条发送任务最终结果")
        logging.info(f"✅ 发送成功：{success_send} 人")
        logging.info(f"❌ 发送失败：{fail_send} 人")
        logging.info(f"⚠️ 跳过发送：{skip_send} 人")
        logging.info(f"📄 总员工数：{len(data_rows)} 人")
        logging.info("="*50)

    except Exception as e:
        logging.error(f"❌ 任务执行失败：{str(e)}", exc_info=True)

if __name__ == "__main__":
    main()