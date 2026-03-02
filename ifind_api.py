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
        self._last_error = ""
        mod = _load_ifind()
        if mod is None:
            self._last_error = "未安装 iFinDAPI，请运行: pip install iFinDAPI"
            return False
        if self.connected:
            return True
        try:
            login = getattr(mod, 'THS_iFinDLogin', None)
            if login is None:
                self._last_error = "iFinD 接口中未找到 THS_iFinDLogin 函数"
                self.connected = False
                return False
            if not self._username or not self._password:
                self._last_error = "请先在「帮助 → iFinD 设置」中填写账号和密码"
                self.connected = False
                return False
            # 转为字符串并去除首尾空格；Windows 下 iFinD 使用 GBK，仅传 ASCII 可避免编码问题
            user = str(self._username).strip()
            pwd = str(self._password).strip()
            res = login(user, pwd)
            # 0=成功, 1=部分版本成功, -201=重复登录(已登录，可正常取数)
            err_code = getattr(res, 'ErrorCode', res)
            if err_code is None:
                err_code = res
            if res in (0, 1, -201) or err_code in (0, 1, -201):
                self.connected = True
            else:
                self._last_error = (
                    f"登录失败，返回值: {err_code}。"
                    "请确认：1) 账号密码为 quantapi.51ifind.com 数据接口账号；"
                    "2) 已开通权限；3) 网络可访问同花顺接口。"
                )
                self.connected = False
        except Exception as e:
            self._last_error = str(e)
            self.connected = False
        return self.connected

    def get_last_error(self):
        """返回最近一次连接失败的原因。"""
        return getattr(self, '_last_error', '') or '未知错误'

    def get_last_bond_error(self):
        """返回最近一次债券查询失败时 iFinD 返回的说明（若有）。"""
        return getattr(self, '_last_bond_error', '') or ''

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

            # 记录最后一次债券查询错误，便于界面提示
            self._last_bond_error = ""

            def parse_bd_result(r, bond_name, issuer_name, maturity_date, coupon_rate, issue_amount):
                """从 THS_BD 返回的 THSData 解析债券名称、发行人等。"""
                import pandas as pd
                if r is None or getattr(r, 'errorcode', -1) != 0 or getattr(r, 'data', None) is None:
                    return bond_name, issuer_name, maturity_date, coupon_rate, issue_amount
                df = r.data
                if isinstance(df, pd.DataFrame) and not df.empty:
                    for col in ['ths_sec_name', 'ths_sec_name_stock', '证券名称', '债券名称', '证券简称', '名称']:
                        if col in df.columns and pd.notna(df[col].iloc[0]):
                            bond_name = str(df[col].iloc[0]).strip()
                            break
                    for col in ['ths_bond_issuer', 'ths_float_holder_name_stock', '发行人', '发行人名称', '发行主体']:
                        if col in df.columns and pd.notna(df[col].iloc[0]):
                            issuer_name = str(df[col].iloc[0]).strip()
                            break
                    for col in ['ths_maturity_date', '到期日', '到期日期']:
                        if col in df.columns and pd.notna(df[col].iloc[0]):
                            maturity_date = df[col].iloc[0]
                            break
                    for col in ['ths_coupon_rate', '票面利率', '当期票面利率(%)']:
                        if col in df.columns and pd.notna(df[col].iloc[0]):
                            coupon_rate = df[col].iloc[0]
                            break
                    for col in ['ths_issue_amount', '发行总额', '发行规模']:
                        if col in df.columns and pd.notna(df[col].iloc[0]):
                            issue_amount = df[col].iloc[0]
                            break
                    if not bond_name and len(df.columns) >= 1:
                        v = df.iloc[0, 0]
                        if pd.notna(v) and str(v).strip():
                            bond_name = str(v).strip()
                    if not issuer_name and len(df.columns) >= 2:
                        v = df.iloc[0, 1]
                        if pd.notna(v) and str(v).strip():
                            issuer_name = str(v).strip()
                elif isinstance(r.data, (list, tuple)) and len(r.data) > 0:
                    tbl = r.data[0] if r.data else {}
                    if isinstance(tbl, dict):
                        bond_name = bond_name or str(tbl.get('ths_sec_name', tbl.get('证券名称', '')) or '').strip()
                        issuer_name = issuer_name or str(tbl.get('ths_bond_issuer', tbl.get('发行人', '')) or '').strip()
                    elif isinstance(tbl, (list, tuple)) and len(tbl) >= 2:
                        bond_name = bond_name or str(tbl[0] or '').strip()
                        issuer_name = issuer_name or str(tbl[1] or '').strip()
                return bond_name, issuer_name, maturity_date, coupon_rate, issue_amount

            # 多种 (指标, 参数) 组合，避免 "the params are invalid"
            # 文档示例为股票: THS_BD('600843.SH','ths_stock_short_name_stock;ths_float_holder_name_stock',';2021-12-31,0','format:json')
            today = datetime.now().strftime('%Y-%m-%d')
            past_date = '2024-01-15'  # 使用已过去的日期，部分接口不接受未来日期
            trials = [
                ('ths_stock_short_name_stock;ths_float_holder_name_stock', ';%s,0' % past_date),
                ('ths_stock_short_name_stock;ths_float_holder_name_stock', ';%s,0' % today),
                ('ths_sec_name;ths_bond_issuer', ';%s,0' % past_date),
                ('ths_sec_name;ths_bond_issuer', ';%s,0' % today),
                ('ths_sec_name;ths_bond_issuer', '0'),
                ('ths_sec_name;ths_bond_issuer', ''),
                ('ths_sec_name', ';%s,0' % past_date),
                ('ths_sec_name', '0'),
            ]
            if ths_bd is not None:
                for indicators, param_option in trials:
                    if bond_name and issuer_name:
                        break
                    try:
                        r = ths_bd(bond_code, indicators, param_option, 'format:dataframe')
                        if r is not None and getattr(r, 'errorcode', -1) != 0:
                            self._last_bond_error = getattr(r, 'errmsg', '') or ("iFinD 错误码: %s" % getattr(r, 'errorcode', ''))
                            continue
                        bond_name, issuer_name, maturity_date, coupon_rate, issue_amount = parse_bd_result(
                            r, bond_name, issuer_name, maturity_date, coupon_rate, issue_amount
                        )
                    except Exception:
                        continue
            if (not bond_name or not issuer_name) and ths_bd is not None:
                for indicators, param_option in [('ths_stock_short_name_stock;ths_float_holder_name_stock', ';%s,0' % past_date), ('ths_sec_name;ths_bond_issuer', '0')]:
                    if bond_name and issuer_name:
                        break
                    try:
                        r = ths_bd(bond_code, indicators, param_option, 'format:list')
                        if r is not None and getattr(r, 'errorcode', -1) != 0 and not self._last_bond_error:
                            self._last_bond_error = getattr(r, 'errmsg', '') or ("iFinD 错误码: %s" % getattr(r, 'errorcode', ''))
                        if r is not None and getattr(r, 'errorcode', -1) == 0 and getattr(r, 'data', None):
                            tbl = r.data
                            if isinstance(tbl, (list, tuple)) and len(tbl) > 0:
                                row = tbl[0]
                                if isinstance(row, dict):
                                    bond_name = bond_name or str(row.get('ths_sec_name', row.get('证券名称', '')) or '').strip()
                                    issuer_name = issuer_name or str(row.get('ths_bond_issuer', row.get('发行人', '')) or '').strip()
                                elif isinstance(row, (list, tuple)) and len(row) >= 2:
                                    bond_name = bond_name or str(row[0] or '').strip()
                                    issuer_name = issuer_name or str(row[1] or '').strip()
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
            if not self._last_bond_error and (bond_code or "").endswith(".IB"):
                self._last_bond_error = "银行间债券(.IB) 可能需在 iFinD 中使用不同代码或接口查询，请核对代码或改用交易所债券(.SH/.SZ) 试一下。"
            return None
        except Exception as e:
            self._last_bond_error = str(e)
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
