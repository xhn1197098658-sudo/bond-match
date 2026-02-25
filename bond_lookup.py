# -*- coding: utf-8 -*-
"""债券发行人查询命令行工具，使用 iFinD（同花顺）数据源。"""
import sys
from data_provider import get_data_api


def lookup_bond(bond_code):
    """使用 iFinD 查询债券发行人信息。"""
    print("数据源: iFinD (同花顺)")
    print(f"正在查询债券: {bond_code}")
    data_api = get_data_api()
    try:
        if not data_api.connect():
            print("错误: 无法连接 iFinD，请配置账号密码或设置环境变量 IFIND_USER / IFIND_PASSWORD")
            return None
        bond_info = data_api.get_bond_info(bond_code)
        if not bond_info:
            print(f"未找到债券 {bond_code} 的信息")
            return None
        issuer_name = bond_info.get("issuer_name", "")
        if not issuer_name:
            print("债券信息中无发行人名称")
            return bond_info
        issuer_info = data_api.get_issuer_info(issuer_name)
        print("\n债券信息:")
        print(f"  债券代码: {bond_info.get('bond_code', '')}")
        print(f"  债券名称: {bond_info.get('bond_name', 'N/A')}")
        print(f"  发行人: {issuer_name}")
        if issuer_info:
            print(f"  发行人代码: {issuer_info.get('issuer_code', 'N/A')}")
            print(f"  行业: {issuer_info.get('industry', 'N/A')}")
            print(f"  评级: {issuer_info.get('credit_rating', 'N/A')}")
        return {"bond": bond_info, "issuer": issuer_info}
    finally:
        data_api.disconnect()


def main():
    print("债券买家匹配 - 债券查询工具 (iFinD)")
    print("===================================")
    if len(sys.argv) > 1:
        bond_code = sys.argv[1]
    else:
        bond_code = input("请输入债券代码: ")
    if not bond_code or len(bond_code) < 5:
        print("错误: 债券代码无效")
        return False
    result = lookup_bond(bond_code.strip())
    return result is not None


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
