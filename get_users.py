import requests
import time

CORP_ID = "wwd5993be58eb1ecb3"
CONTACT_SECRET = input("请输入通讯录同步secret: ").strip()

def get_access_token(corp_id, secret):
    url = f"https://qyapi.weixin.qq.com/cgi-bin/gettoken?corpid={corp_id}&corpsecret={secret}"
    resp = requests.get(url, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    if data.get("errcode") != 0:
        raise Exception(f"获取access_token失败: {data.get('errmsg')}")
    return data["access_token"]

def get_all_users(access_token):
    all_users = []
    
    url = f"https://qyapi.weixin.qq.com/cgi-bin/department/list?access_token={access_token}"
    resp = requests.get(url, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    
    if data.get("errcode") != 0:
        raise Exception(f"获取部门列表失败: {data.get('errmsg')}")
    
    departments = data.get("department", [])
    print(f"发现 {len(departments)} 个部门")
    
    for dept in departments:
        dept_id = dept["id"]
        dept_name = dept["name"]
        
        url = f"https://qyapi.weixin.qq.com/cgi-bin/user/list?access_token={access_token}&department_id={dept_id}&fetch_child=0"
        resp = requests.get(url, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        
        if data.get("errcode") == 0:
            users = data.get("userlist", [])
            print(f"  部门 [{dept_name}] 有 {len(users)} 人")
            all_users.extend(users)
            time.sleep(0.2)
    
    return all_users

def main():
    try:
        print("正在获取access_token...")
        token = get_access_token(CORP_ID, CONTACT_SECRET)
        print(f"获取成功: {token[:20]}...")
        
        print("\n正在获取所有用户列表...")
        users = get_all_users(token)
        
        print(f"\n✅ 共获取到 {len(users)} 个用户")
        
        with open("企业微信用户列表.txt", "w", encoding="utf-8") as f:
            f.write("姓名\tUserID\t部门\n")
            f.write("=" * 50 + "\n")
            for user in users:
                name = user.get("name", "未知")
                userid = user.get("userid", "未知")
                dept = user.get("department", [])
                dept_str = ",".join(map(str, dept))
                f.write(f"{name}\t{userid}\t{dept_str}\n")
                print(f"  {name} -> {userid}")
        
        print("\n📄 用户列表已保存到: 企业微信用户列表.txt")
        
        with open("用户映射配置.py", "w", encoding="utf-8") as f:
            f.write("MANUAL_USER_MAP = {\n")
            for user in users:
                name = user.get("name", "")
                userid = user.get("userid", "")
                if name and userid:
                    f.write(f'    "{name}": "{userid}",\n')
            f.write("}\n")
        
        print("📄 用户映射配置已保存到: 用户映射配置.py")
        
    except Exception as e:
        print(f"❌ 错误: {str(e)}")

if __name__ == "__main__":
    main()