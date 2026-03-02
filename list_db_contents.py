# -*- coding: utf-8 -*-
"""查看当前数据库中有哪些数据。运行：python list_db_contents.py"""
import os
import sys

def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    sys.path.insert(0, script_dir)
    os.chdir(script_dir)

    from database.db_manager import DatabaseManager
    db = DatabaseManager()
    conn = db.get_connection()
    cursor = conn.cursor()

    tables = [
        ('issuers', '发行人'),
        ('bonds', '债券'),
        ('buyside_companies', '买方机构'),
        ('funds', '基金'),
        ('holdings', '持仓'),
        ('can_buy_lists', '可买名单'),
        ('contacts', '联系人'),
    ]
    print("=" * 50)
    print("数据库内容概览（database/bond_buyer_match.db）")
    print("=" * 50)
    for table, name in tables:
        cursor.execute(f"SELECT COUNT(*) FROM {table}")
        n = cursor.fetchone()[0]
        print(f"  {name} ({table}): {n} 条")
    print()
    print("说明：潜在买家来自「可买名单」和「持仓」。若两者均为 0，请先在 data/ 放入 Excel，再通过「帮助→更新数据」导入。")
    print("=" * 50)
    db.close_connection()

if __name__ == "__main__":
    main()
