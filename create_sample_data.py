# -*- coding: utf-8 -*-
"""在 data/ 下生成示例 Excel，用于测试「潜在买家」。运行：python create_sample_data.py"""
import os
import sys

def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)
    os.makedirs('data', exist_ok=True)

    try:
        import pandas as pd
    except ImportError:
        print("需要 pandas、openpyxl：pip install pandas openpyxl")
        return 1

    # 010107.SH 的发行人为「中华人民共和国财政部」（国债）
    issuer_name = "中华人民共和国财政部"
    bond_code = "010107.SH"
    bond_name = "21国债(7)"

    # 可买名单：哪些机构可以买该发行人
    can_buy = pd.DataFrame([
        {'company_name': '示例基金公司A', 'company_type': '基金', 'issuer_name': issuer_name},
        {'company_name': '示例资管公司B', 'company_type': '资管', 'issuer_name': issuer_name},
    ])
    path_can = os.path.join('data', 'sample_can_buy_lists.xlsx')
    can_buy.to_excel(path_can, index=False)
    print(f"已生成: {path_can}")

    # 持仓：机构持有哪些债券（发行人一致即可匹配为潜在买家）
    holdings = pd.DataFrame([
        {'company_name': '示例基金公司A', 'fund_name': '示例债券基金', 'issuer_name': issuer_name, 'bond_code': bond_code, 'bond_name': bond_name, 'amount': 100},
    ])
    path_hold = os.path.join('data', 'sample_holdings.xlsx')
    holdings.to_excel(path_hold, index=False)
    print(f"已生成: {path_hold}")

    # 联系人
    contacts = pd.DataFrame([
        {'company_name': '示例基金公司A', 'fund_name': '示例债券基金', 'name': '张三', 'position': '投资经理', 'email': 'zhangsan@example.com', 'phone': '13800138000', 'is_primary': '是'},
        {'company_name': '示例资管公司B', 'name': '李四', 'position': '固收总监', 'phone': '13900139000'},
    ])
    path_contact = os.path.join('data', 'sample_contacts.xlsx')
    contacts.to_excel(path_contact, index=False)
    print(f"已生成: {path_contact}")

    print()
    print("下一步：打开程序 → 帮助 → 更新数据 → 选择 data 文件夹，导入后搜索 010107.SH 即可看到潜在买家。")
    return 0

if __name__ == "__main__":
    sys.exit(main())
