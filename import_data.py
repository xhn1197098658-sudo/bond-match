# -*- coding: utf-8 -*-
"""从 Excel 导入持仓、可买名单、联系人到数据库。不依赖 Wind/iFinD。"""
import os
import sys
import glob
from database.db_manager import DatabaseManager


def find_excel_file(directory, base_name):
    """在目录中查找包含 base_name 的 Excel 文件（.xlsx 或 .xls）。"""
    for ext in ('.xlsx', '.xls'):
        path = os.path.join(directory, f"{base_name}{ext}")
        if os.path.exists(path):
            return path
    for ext in ('.xlsx', '.xls'):
        for path in glob.glob(os.path.join(directory, f"*{base_name}*{ext}")):
            return path
    return None


def validate_file_path(path):
    """校验路径存在且为 Excel 文件。"""
    if not path:
        return None
    if os.path.isdir(path):
        files = glob.glob(os.path.join(path, "*.xlsx")) + glob.glob(os.path.join(path, "*.xls"))
        if not files:
            print(f"目录下未找到 Excel 文件: {path}")
            return None
        print("找到的 Excel 文件:")
        for i, f in enumerate(files, 1):
            print(f"  {i}. {os.path.basename(f)}")
        try:
            idx = int(input("请输入序号 (0 取消): "))
            if 1 <= idx <= len(files):
                return files[idx - 1]
        except (ValueError, IndexError):
            pass
        return None
    if os.path.exists(path) and (path.endswith('.xlsx') or path.endswith('.xls')):
        return path
    print(f"文件不存在或不是 Excel: {path}")
    return None


def main():
    print("债券买家匹配 - 数据导入")
    print("===================================")
    os.makedirs('database', exist_ok=True)
    db = DatabaseManager()
    data_dir = 'data'
    os.makedirs(data_dir, exist_ok=True)
    holdings_file = find_excel_file(data_dir, 'holdings')
    can_buy_file = find_excel_file(data_dir, 'can_buy_lists')
    contacts_file = find_excel_file(data_dir, 'contacts')
    if not holdings_file:
        holdings_file = validate_file_path(input("请输入持仓 Excel 路径: "))
    if not can_buy_file:
        can_buy_file = validate_file_path(input("请输入可买名单 Excel 路径: "))
    if not contacts_file:
        contacts_file = validate_file_path(input("请输入联系人 Excel 路径: "))
    if not all((holdings_file, can_buy_file, contacts_file)):
        print("缺少必要文件，退出")
        return False
    print("\n将导入:")
    print(f"  持仓: {holdings_file}")
    print(f"  可买名单: {can_buy_file}")
    print(f"  联系人: {contacts_file}")
    if input("确认导入? (y/n): ").strip().lower() != 'y':
        print("已取消")
        return False
    ok, counts = db.import_from_excel(holdings_file, can_buy_file, contacts_file)
    if ok:
        print("导入完成:", counts)
    else:
        print("导入失败")
    return ok


if __name__ == "__main__":
    sys.exit(0 if main() else 1)
