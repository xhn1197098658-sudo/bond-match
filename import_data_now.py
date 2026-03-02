# -*- coding: utf-8 -*-
"""一键导入 data 文件夹内三个 Excel 到数据库（无需确认）。"""
import os
import sys

def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)
    data_dir = os.path.join(script_dir, 'data')
    holdings = os.path.join(data_dir, 'sample_holdings.xlsx')
    can_buy = os.path.join(data_dir, 'sample_can_buy_lists.xlsx')
    contacts = os.path.join(data_dir, 'sample_contacts.xlsx')
    for name, path in [('持仓', holdings), ('可买名单', can_buy), ('联系人', contacts)]:
        if not os.path.exists(path):
            print(f"未找到: {path}")
            return 1
    from database.db_manager import DatabaseManager
    db = DatabaseManager()
    ok, counts = db.import_from_excel(holdings, can_buy, contacts)
    db.close_connection()
    if ok:
        print("导入完成:", counts)
        return 0
    print("导入失败")
    return 1

if __name__ == "__main__":
    sys.exit(main())
