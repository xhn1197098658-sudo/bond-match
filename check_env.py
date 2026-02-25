# -*- coding: utf-8 -*-
"""快速自检：Python 环境、依赖、数据库、iFinD 是否可用。运行：python check_env.py"""
import sys
import os

def main():
    print("=" * 50)
    print("债券买家匹配 - 环境自检")
    print("=" * 50)

    # 1. Python 版本
    print(f"\n[1] Python 版本: {sys.version.split()[0]}")
    if sys.version_info < (3, 8):
        print("    建议使用 Python 3.8 及以上")
    else:
        print("    通过")

    # 2. 项目目录
    root = os.path.dirname(os.path.abspath(__file__))
    os.chdir(root)
    print(f"\n[2] 项目目录: {root}")

    # 3. 依赖包
    print("\n[3] 依赖包:")
    for name in ("PyQt5", "pandas", "openpyxl"):
        try:
            __import__(name)
            print(f"    {name}: 已安装")
        except ImportError:
            print(f"    {name}: 未安装 -> pip install {name}")

    # 4. 数据库与 data_provider
    print("\n[4] 项目模块:")
    try:
        from database.db_manager import DatabaseManager
        db = DatabaseManager()
        db.get_connection()
        print("    database.db_manager: 正常")
        db.close_connection()
    except Exception as e:
        print(f"    database.db_manager: 失败 - {e}")

    try:
        from data_provider import get_data_api
        api = get_data_api()
        print("    data_provider.get_data_api: 正常（返回 iFinD API 实例）")
    except Exception as e:
        print(f"    data_provider: 失败 - {e}")

    # 5. iFinD 是否可用
    print("\n[5] iFinD 接口:")
    try:
        from ifind_api import IFIND_AVAILABLE, _load_ifind
        mod = _load_ifind()
        if mod is not None:
            print("    iFinDPy/iFinDAPI: 已安装，可加载")
            login = getattr(mod, 'THS_iFinDLogin', None)
            print(f"    登录函数 THS_iFinDLogin: {'可用' if login else '未找到'}")
        else:
            print("    iFinDPy/iFinDAPI: 未安装 -> pip install iFinDAPI")
    except Exception as e:
        print(f"    ifind_api: 异常 - {e}")

    # 6. 配置文件/环境变量提示
    print("\n[6] iFinD 账号配置:")
    try:
        from PyQt5.QtCore import QSettings
        s = QSettings("BondBuyerMatch", "App")
        user = s.value("iFind/username", "")
        if user:
            print("    已保存账号（来自上次「iFinD 设置」）")
        else:
            print("    未保存账号。可在运行 app.py 后通过「帮助 → iFinD 设置」填写，或设置环境变量 IFIND_USER / IFIND_PASSWORD")
    except ImportError:
        print("    （需安装 PyQt5 后才能在程序中保存账号；或直接设置环境变量 IFIND_USER / IFIND_PASSWORD）")

    print("\n" + "=" * 50)
    print("自检结束。若上述无报错，可运行: python app.py  或  python bond_lookup.py 132001.SH")
    print("=" * 50)

if __name__ == "__main__":
    main()
