# -*- coding: utf-8 -*-
"""
iFinD (同花顺) 数据接口封装，与 WindAPI 保持相同接口，便于切换数据源。

使用前请：
  1. pip install iFinDAPI
  2. 在同花顺数据接口平台申请账号：https://quantapi.51ifind.com/
  3. 在程序中「帮助 → 数据源设置」选择 iFinD 并填写账号、密码

适配说明（若你使用的 iFinD 版本接口不同）：
  - 登录函数多为 THS_iFinDLogin(用户名, 密码)，返回 0 表示成功。
  - 取数函数可能是 THS_Data、THS_BD 等，参数与指标名以官方「超级命令」或
    文档为准；可在本文件中搜索 THS_Data / THS_BD 并按你本地接口修改指标字符串。
  - 债券名称、发行人等指标名可能为 ths_sec_name、ths_bond_issuer 等，
    具体以同花顺提供的指标列表为准。
"""

import os
import re
from datetime import datetime

try:
    import pandas as pd
except ImportError:
    pd = None

IFIND_AVAILABLE = False
_ifind_module = None  # 延迟加载 iFinDPy


def _load_ifind():
    """尝试加载 iFinD Python 接口（pip install iFinDAPI）。支持 iFinDPy 或 iFinDAPI 包名。"""
    global IFIND_AVAILABLE, _ifind_module
    if _ifind_module is not None:
        return _ifind_module if _ifind_module else None
    for pkg_name in ('iFinDPy', 'iFinDAPI'):
        try:
            mod = __import__(pkg_name)
            # 优先取 THS_iFinDLogin, THS_Data 等
            login = getattr(mod, 'THS_iFinDLogin', None)
            ths_data = getattr(mod, 'THS_Data', None)
            ths_bd = getattr(mod, 'THS_BD', None)
            if login is not None or ths_data is not None or ths_bd is not None:
                _ifind_module = mod
                IFIND_AVAILABLE = True
                return _ifind_module
            _ifind_module = mod
            IFIND_AVAILABLE = True
            return _ifind_module
        except ImportError:
            continue
    _ifind_module = False
    return None


class iFindAPI:
    """iFinD 数据接口，与 WindAPI 保持相同对外接口。"""

    def __init__(self, username=None, password=None, ifind_path=""):
        self.connected = False
        self._username = username or os.environ.get("IFIND_USER", "")
        self._password = password or os.environ.get("IFIND_PASSWORD", "")
        self._ifind_path = ifind_path
        if _load_ifind() is not None:
            self.connect()

    def connect(self):
        """连接 iFinD 数据服务（登录）。"""
        mod = _load_ifind()
        if mod is None:
            return False
        if self.connected:
            return True
        try:
            login = getattr(mod, 'THS_iFinDLogin', None)
            if login is None:
                self.connected = False
                return False
            if not self._username or not self._password:
                self.connected = False
                return False
            res = login(self._username, self._password)
            if res == 0 or (hasattr(res, 'ErrorCode') and getattr(res, 'ErrorCode', -1) == 0):
                self.connected = True
            else:
                self.connected = False
        except Exception:
            self.connected = False
        return self.connected

    def disconnect(self):
        """断开 iFinD 连接（如有对应接口）。"""
        self.connected = False
        # iFinD 部分版本无显式 close，仅清状态即可

    def _ensure_connected(self):
        if _load_ifind() is None:
            return False
        if not self.connected and not self.connect():
            return False
        return True

    def _preprocess_bond_code(self, bond_code):
        """与 Wind 一致：缺后缀时补 .SH。"""
        if re.match(r'^\d{6}$', bond_code) and '.' not in bond_code:
            bond_code = bond_code + '.SH'
        return bond_code

    def get_bond_info(self, bond_code):
        """
        获取债券信息，返回与 WindAPI.get_bond_info 相同结构：
        bond_code, bond_name, issuer_name, bond_type, maturity_date, coupon_rate, issue_amount
        """
        if not self._ensure_connected():
            return None
        bond_code = self._preprocess_bond_code(bond_code)
        mod = _load_ifind()
        if mod is None:
            return None
        try:
            ths_bd = getattr(mod, 'THS_BD', None)
            ths_data = getattr(mod, 'THS_Data', None)

            bond_name = ""
            issuer_name = ""
            maturity_date = None
            coupon_rate = None
            issue_amount = None

            # 方式1：债券专用 THS_BD（若存在）
            if ths_bd is not None:
                try:
                    # 示例：THS_BD('code','date','indicator')，具体参数见 iFinD 文档
                    r = ths_bd(bond_code, datetime.now().strftime('%Y-%m-%d'), 'ths_sec_name;ths_bond_issuer;ths_maturity_date;ths_coupon_rate;ths_issue_amount')
                    if r and (r == 0 or (hasattr(r, 'ErrorCode') and getattr(r, 'ErrorCode', 0) == 0)):
                        if hasattr(r, 'Data') and r.Data:
                            d = r.Data
                            bond_name = d[0][0] if len(d) > 0 and d[0] else ""
                            issuer_name = d[1][0] if len(d) > 1 and d[1] else ""
                            maturity_date = d[2][0] if len(d) > 2 and d[2] else None
                            coupon_rate = d[3][0] if len(d) > 3 and d[3] else None
                            issue_amount = d[4][0] if len(d) > 4 and d[4] else None
                except Exception:
                    pass

            # 方式2：通用 THS_Data 截面（指标名需与 iFinD 超级命令一致）
            if (not bond_name or not issuer_name) and ths_data is not None:
                try:
                    # 常见 iFinD 债券指标名示例，实际以官方文档为准
                    r = ths_data(bond_code, 'ths_sec_name;ths_bond_issuer;ths_maturity_date;ths_coupon_rate;ths_issue_amount', 'None', 'D', datetime.now().strftime('%Y-%m-%d'))
                    if r and hasattr(r, 'data') and r.data:
                        arr = r.data
                        if isinstance(arr, (list, tuple)) and len(arr) > 0:
                            bond_name = bond_name or (arr[0][0] if arr[0] else "")
                            issuer_name = issuer_name or (arr[1][0] if len(arr) > 1 and arr[1] else "")
                            if len(arr) > 2 and arr[2]:
                                maturity_date = arr[2][0]
                            if len(arr) > 3 and arr[3]:
                                coupon_rate = arr[3][0]
                            if len(arr) > 4 and arr[4]:
                                issue_amount = arr[4][0]
                except Exception:
                    pass

            if bond_name or issuer_name:
                return {
                    'bond_code': bond_code,
                    'bond_name': bond_name or '',
                    'issuer_name': issuer_name or '',
                    'bond_type': '',
                    'maturity_date': maturity_date,
                    'coupon_rate': coupon_rate,
                    'issue_amount': issue_amount,
                }
            return None
        except Exception:
            return None

    def get_issuer_info(self, issuer_name):
        """
        获取发行人信息，返回与 WindAPI.get_issuer_info 相同结构：
        issuer_name, issuer_code, industry, credit_rating
        """
        if not self._ensure_connected():
            return None
        mod = _load_ifind()
        if mod is None:
            return None
        try:
            ths_data = getattr(mod, 'THS_Data', None)
            industry = None
            credit_rating = None
            issuer_code = None
            if ths_data is not None:
                # 按发行人名称查行业、评级等，具体指标以 iFinD 文档为准
                try:
                    r = ths_data(issuer_name, 'ths_industry;ths_rating', 'None', 'D', datetime.now().strftime('%Y-%m-%d'))
                    if r and hasattr(r, 'data') and r.data and len(r.data) >= 2:
                        industry = r.data[0][0] if r.data[0] else None
                        credit_rating = r.data[1][0] if r.data[1] else None
                    issuer_code = issuer_name  # iFinD 可能用名称即代码
                except Exception:
                    pass
            return {
                'issuer_name': issuer_name,
                'issuer_code': issuer_code,
                'industry': industry,
                'credit_rating': credit_rating,
            }
        except Exception:
            return {
                'issuer_name': issuer_name,
                'issuer_code': None,
                'industry': None,
                'credit_rating': None,
            }
