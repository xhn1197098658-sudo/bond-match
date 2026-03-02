# -*- coding: utf-8 -*-
"""命令行测试 iFinD 登录，查看实际返回值。用法：python test_ifind_login.py [账号] [密码]"""
import sys
import os

def main():
    # 从环境变量或命令行读取
    user = os.environ.get("IFIND_USER", "")
    pwd = os.environ.get("IFIND_PASSWORD", "")
    if len(sys.argv) >= 3:
        user, pwd = sys.argv[1], sys.argv[2]
    elif len(sys.argv) == 2:
        user = sys.argv[1]
        pwd = input("请输入密码: ").strip()

    if not user or not pwd:
        print("用法: python test_ifind_login.py <账号> <密码>")
        print("或设置环境变量 IFIND_USER, IFIND_PASSWORD 后直接运行")
        return 1

    print("正在尝试连接 iFinD...")
    print("账号:", user[:3] + "***" if len(user) > 3 else "***")

    try:
        mod = __import__("iFinDPy")
    except ImportError:
        try:
            mod = __import__("iFinDAPI")
        except ImportError:
            print("未安装 iFinDAPI，请运行: pip install iFinDAPI")
            return 1

    login = getattr(mod, "THS_iFinDLogin", None)
    if not login:
        print("未找到 THS_iFinDLogin")
        return 1

    res = login(user, pwd)
    print("返回值:", res, "  类型:", type(res).__name__)
    if hasattr(res, "ErrorCode"):
        print("ErrorCode:", getattr(res, "ErrorCode", "无"))

    if res == 0 or res == 1:
        print(">>> 登录成功")
        return 0
    print(">>> 登录失败，请根据返回值检查账号、密码及数据接口权限")
    return 1

if __name__ == "__main__":
    sys.exit(main())
