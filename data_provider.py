# -*- coding: utf-8 -*-
"""数据源：仅使用 iFinD（同花顺）。"""

import os


def get_data_api(settings=None, **kwargs):
    """
    返回 iFinD 数据 API 实例。
    - settings: QSettings，用于读取 iFind/username、iFind/password
    - kwargs: 可传 ifind_user, ifind_password, ifind_path 覆盖
    返回对象：connect(), disconnect(), get_bond_info(code), get_issuer_info(name)
    """
    from ifind_api import iFindAPI
    username = kwargs.get("ifind_user") or (settings.value("iFind/username", "") if settings else "") or os.environ.get("IFIND_USER", "")
    password = kwargs.get("ifind_password") or (settings.value("iFind/password", "") if settings else "") or os.environ.get("IFIND_PASSWORD", "")
    path = kwargs.get("ifind_path") or (settings.value("iFind/path", "") if settings else "")
    return iFindAPI(username=username, password=password, ifind_path=path)
