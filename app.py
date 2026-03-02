import sys
import os
import re
from pathlib import Path


def _parse_year_from_filename(name):
    """从文件名解析年度，如 23年表三、2023年、表三23、2023表 → 2023"""
    if not name:
        return ""
    s = os.path.splitext(str(name))[0]
    # 优先匹配完整年份 20XX
    m = re.search(r"(20\d{2})", s)
    if m:
        return m.group(1)
    # XX年（前后均可）
    m = re.search(r"(\d{2})年", s)
    if m:
        y = int(m.group(1))
        return str(2000 + y) if y < 50 else str(1900 + y)
    # 开头两位数字（如 23表三）
    m = re.search(r"^(\d{2})[^\d]", s)
    if m:
        y = int(m.group(1))
        return str(2000 + y) if y < 50 else str(1900 + y)
    # 结尾两位数字（如 表三23、持仓23）
    m = re.search(r"[^\d](\d{2})(?:\D|$)", s)
    if m:
        y = int(m.group(1))
        return str(2000 + y) if y < 50 else str(1900 + y)
    return ""


# #region agent log
def _dbg(hid, msg, data=None):
    try:
        import time as _t
        import json as _j
        _lp = r'C:\Users\86159\Documents\bond-match\debug-c1ea61.log'
        if getattr(sys, 'frozen', False):
            _lp = os.path.join(os.path.dirname(sys.executable), 'debug-c1ea61.log')
        with open(_lp, 'a', encoding='utf-8') as _f:
            _f.write(_j.dumps({"hypothesisId": hid, "location": "app.py", "message": msg, "data": data or {}, "timestamp": int(_t.time() * 1000)}, ensure_ascii=False) + '\n')
    except Exception:
        pass
try:
    _dbg("H1", "pre-import env", {"executable": getattr(sys, 'executable', '?'), "prefix": getattr(sys, 'prefix', '?'), "frozen": getattr(sys, 'frozen', False), "cwd": os.getcwd()})
except Exception:
    pass
try:
    import _ctypes as _ct
    _dbg("H2", "_ctypes import OK", {})
except Exception as _e:
    _dbg("H2", "_ctypes import FAIL", {"error": str(_e), "type": type(_e).__name__})
# #endregion

import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QLabel, QLineEdit, QPushButton, 
                            QTableWidget, QTableWidgetItem, QTabWidget, 
                            QGroupBox, QFormLayout, QMessageBox, QHeaderView,
                            QStatusBar, QSplitter, QTextEdit, QFileDialog,
                            QDialog, QDialogButtonBox)
from PyQt5.QtCore import Qt, QSettings, QTimer, QObject, pyqtSignal, QEvent
from PyQt5.QtGui import QFont, QIcon
import traceback
import glob
import time
import uuid

from database.db_manager import DatabaseManager
from data_provider import get_data_api


class iFinDSettingsDialog(QDialog):
    """iFinD（同花顺）账号设置"""
    def __init__(self, parent=None, settings=None):
        super().__init__(parent)
        self.settings = settings or QSettings("BondBuyerMatch", "App")
        self.setWindowTitle("iFinD 设置")
        self.setMinimumWidth(450)
        layout = QVBoxLayout(self)
        desc = QLabel("请填写同花顺数据接口账号与密码。需先安装：pip install iFinDAPI")
        desc.setWordWrap(True)
        layout.addWidget(desc)
        form = QFormLayout()
        self.user_input = QLineEdit(self.settings.value("iFind/username", ""))
        self.user_input.setPlaceholderText("同花顺数据接口账号")
        self.pass_input = QLineEdit(self.settings.value("iFind/password", ""))
        self.pass_input.setEchoMode(QLineEdit.Password)
        self.pass_input.setPlaceholderText("密码")
        form.addRow("账号:", self.user_input)
        form.addRow("密码:", self.pass_input)
        layout.addLayout(form)
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self._on_accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    def _on_accept(self):
        self.settings.setValue("iFind/username", self.user_input.text().strip())
        self.settings.setValue("iFind/password", self.pass_input.text().strip())
        self.settings.sync()  # 立即写入，确保保存成功
        self.accept()


class DeleteBatchDialog(QDialog):
    """删除指定导入批次"""
    TYPE_LABELS = {"holdings": "持仓", "can_buy_lists": "可买名单", "contacts": "联系人"}

    def __init__(self, parent=None, db_manager=None):
        super().__init__(parent)
        self.db_manager = db_manager or DatabaseManager()
        self.deleted_count = None
        self.setWindowTitle("删除指定导入")
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("选择要删除的导入批次（仅删除该批次的记录）："))
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["文件名", "类型", "条数", "导入时间"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        layout.addWidget(self.table)
        btn_layout = QHBoxLayout()
        delete_btn = QPushButton("删除所选批次")
        delete_btn.clicked.connect(self._do_delete)
        btn_layout.addStretch()
        btn_layout.addWidget(delete_btn)
        layout.addLayout(btn_layout)
        button_box = QDialogButtonBox(QDialogButtonBox.Close)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        self._load_batches()

    def _load_batches(self):
        batches = self.db_manager.get_import_batches()
        self.table.setRowCount(len(batches))
        for i, b in enumerate(batches):
            self.table.setItem(i, 0, QTableWidgetItem(str(b.get("filename", ""))))
            self.table.setItem(i, 1, QTableWidgetItem(self.TYPE_LABELS.get(b.get("data_type", ""), b.get("data_type", ""))))
            self.table.setItem(i, 2, QTableWidgetItem(str(b.get("row_count", 0))))
            self.table.setItem(i, 3, QTableWidgetItem(str(b.get("imported_at", ""))))
            self.table.item(i, 0).setData(Qt.UserRole, b.get("batch_id"))
        if not batches:
            layout = self.layout()
            lbl = QLabel("暂无导入记录（此前导入的数据无批次信息）。")
            layout.insertWidget(1, lbl)

    def _do_delete(self):
        row = self.table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "未选择", "请先选择要删除的批次。")
            return
        batch_id = self.table.item(row, 0).data(Qt.UserRole)
        fn = self.table.item(row, 0).text()
        dt = self.table.item(row, 1).text()
        r = QMessageBox.warning(
            self,
            "确认删除",
            f"确定删除该批次的导入数据？\n\n文件：{fn}\n类型：{dt}\n\n此操作不可恢复。",
            QMessageBox.Ok | QMessageBox.Cancel,
            QMessageBox.Cancel
        )
        if r != QMessageBox.Ok:
            return
        self.deleted_count = self.db_manager.delete_batch(batch_id)
        self.accept()


class BondBuyerMatchApp(QMainWindow):
    """Main application window for Bond Buyer Match"""
    
    def __init__(self):
        super().__init__()
        
        # Load settings
        self.settings = QSettings("BondBuyerMatch", "App")
        
        # Initialize API and database
        self.data_api = None
        self.init_data_api()
        
        self.db_manager = DatabaseManager()
        
        # Store the current search results
        self.current_issuer_id = None
        self.current_bond_info = None
        self.current_issuer_info = None
        
        # Remarks and bond_times per (bond_code, company_name) for results table
        self._remarks = {}
        self._bond_times = {}
        self._data_year = ""  # 数据年度，从导入文件名解析（如 23年表三 → 2023）

        # Selection handling
        self.last_selection_time = 0
        self.debounce_timer = QTimer()
        self.debounce_timer.setSingleShot(True)
        self.debounce_timer.timeout.connect(self._handle_delayed_selection)
        self.pending_selection = None
        self.selection_count = 0
        self.display_count = 0
        
        # Set up the UI
        self.init_ui()
        
        # Load other settings
        self.load_settings()
    
    def init_data_api(self):
        """初始化 iFinD 数据 API。"""
        try:
            self.data_api = get_data_api(settings=self.settings)
            if not self.data_api.connected:
                err = getattr(self.data_api, 'get_last_error', lambda: '')()
                msg = "无法连接 iFinD。\n\n"
                if err:
                    msg += f"原因：{err}\n\n"
                msg += "请确认：\n"
                msg += "1. 已在「帮助 → iFinD 设置」中填写同花顺数据接口的账号与密码；\n"
                msg += "2. 账号为 quantapi.51ifind.com 申请的数据接口账号（非炒股软件账号）；\n"
                msg += "3. 已安装：pip install iFinDAPI"
                QMessageBox.warning(self, "iFinD 未连接", msg)
                return True
        except Exception as e:
            QMessageBox.warning(
                self,
                "iFinD 连接错误",
                f"无法连接 iFinD: {str(e)}\n\n{traceback.format_exc()}"
            )
            self.data_api = None
            return False

    def init_ui(self):
        """Initialize the user interface"""
        self.setWindowTitle("债券买家匹配")
        self.setGeometry(100, 100, 1200, 800)
        
        # Create the central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Main layout
        main_layout = QVBoxLayout(central_widget)
        
        # Create a horizontal layout for search and bond info
        top_layout = QHBoxLayout()
        
        # Create the search section
        search_group = QGroupBox("债券搜索")
        search_layout = QHBoxLayout(search_group)
        
        # Bond code label and input
        self.bond_code_label = QLabel("债券代码:")
        self.bond_code_input = QLineEdit()
        self.bond_code_input.setPlaceholderText("输入债券代码 (例如: 132001.SH)")
        # Connect Enter key press to search function
        self.bond_code_input.returnPressed.connect(self.search_bond)
        self.search_button = QPushButton("搜索")
        self.search_button.clicked.connect(self.search_bond)
        
        # Add to search layout
        search_layout.addWidget(self.bond_code_label)
        search_layout.addWidget(self.bond_code_input)
        search_layout.addWidget(self.search_button)
        
        # Bond info group
        bond_info_group = QGroupBox("债券信息")
        bond_info_layout = QFormLayout(bond_info_group)
        
        # Bond info fields
        self.bond_name_label = QLabel("-")
        self.issuer_name_label = QLabel("-")
        
        # Add fields to form layout
        bond_info_layout.addRow("债券名称:", self.bond_name_label)
        bond_info_layout.addRow("发行人名称:", self.issuer_name_label)
        
        # Add search and bond info to the top layout
        top_layout.addWidget(search_group, 7)  # search takes 70% of width
        top_layout.addWidget(bond_info_group, 3)  # bond info takes 30% of width
        
        # Add the top layout to the main layout
        main_layout.addLayout(top_layout, 1)  # top section takes 10% of height
        
        # Create a splitter for the results
        results_splitter = QSplitter(Qt.Horizontal)
        
        # Create the results section
        results_group = QGroupBox("潜在买家")
        results_layout = QVBoxLayout(results_group)
        
        # Create the results table
        self.results_table = QTableWidget()
        self.results_table.setColumnCount(4)
        self.results_table.setHorizontalHeaderLabels(["公司名称", "匹配类型", "债券时间", "备注"])
        self.results_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.results_table.setEditTriggers(QTableWidget.DoubleClicked)
        self.results_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.results_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.results_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.results_table.horizontalHeader().setMinimumSectionSize(70)
        self.results_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.results_table.itemChanged.connect(self._on_results_cell_edited)
        # Set alternate row colors to improve readability
        self.results_table.setAlternatingRowColors(True)
        # Connect selection change after everything is set up
        self.results_table.selectionModel().selectionChanged.connect(self.show_selected_company_details)
        # Style the header
        self.results_table.horizontalHeader().setStyleSheet("QHeaderView::section { background-color: #f0f0f0; padding: 6px; border: 1px solid #ddd; }")
        
        # Add the results table to the layout
        results_layout.addWidget(self.results_table)
        
        # Add results group to the splitter
        results_splitter.addWidget(results_group)
        
        # Create the details section
        details_group = QGroupBox("联系人详情")
        details_layout = QVBoxLayout(details_group)
        
        # Create the contacts table
        self.contacts_table = QTableWidget()
        self.contacts_table.setColumnCount(8)
        self.contacts_table.setHorizontalHeaderLabels(["姓名", "职位", "电子邮件", "电话", "QQ", "QT", "微信", "手机"])
        self.contacts_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.contacts_table.horizontalHeader().setStretchLastSection(False)
        self.contacts_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        # Set alternate row colors for better visibility
        self.contacts_table.setAlternatingRowColors(True)
        # Set default row height to ensure proper spacing
        self.contacts_table.verticalHeader().setDefaultSectionSize(30)
        # Hide vertical header (row numbers)
        self.contacts_table.verticalHeader().setVisible(False)
        # Set wider default column widths
        column_widths = [120, 120, 180, 140, 120, 120, 140, 140]
        for i, width in enumerate(column_widths):
            self.contacts_table.setColumnWidth(i, width)
        # Style the header
        self.contacts_table.horizontalHeader().setStyleSheet("QHeaderView::section { background-color: #f0f0f0; padding: 6px; border: 1px solid #ddd; }")
        
        # Add the contacts table to the layout
        details_layout.addWidget(self.contacts_table)
        
        # Fund holdings section (if applicable)
        self.fund_holdings_group = QGroupBox("基金持仓")
        fund_holdings_layout = QVBoxLayout(self.fund_holdings_group)
        
        # Create the fund holdings table
        self.fund_holdings_table = QTableWidget()
        self.fund_holdings_table.setColumnCount(3)
        self.fund_holdings_table.setHorizontalHeaderLabels(["基金名称", "债券", "金额"])
        self.fund_holdings_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.fund_holdings_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        # Set alternate row colors
        self.fund_holdings_table.setAlternatingRowColors(True)
        # Hide vertical header (row numbers)
        self.fund_holdings_table.verticalHeader().setVisible(False)
        # Style the header
        self.fund_holdings_table.horizontalHeader().setStyleSheet("QHeaderView::section { background-color: #f0f0f0; padding: 6px; border: 1px solid #ddd; }")
        
        # Add the fund holdings table to the layout
        fund_holdings_layout.addWidget(self.fund_holdings_table)
        
        # Add fund holdings to details layout
        details_layout.addWidget(self.fund_holdings_group)
        
        # Add details group to the splitter
        results_splitter.addWidget(details_group)
        
        # Set the initial sizes for the splitter
        results_splitter.setSizes([500, 500])
        
        # Add the splitter to the main layout
        main_layout.addWidget(results_splitter, 9)  # results section takes 90% of height
        
        # Set up the status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("就绪")
        
        # Set up menu
        self.setup_menu()
    
    def setup_menu(self):
        """Set up the application menu"""
        menubar = self.menuBar()
        
        # Help menu
        help_menu = menubar.addMenu("帮助")
        
        ifind_action = help_menu.addAction("iFinD 设置")
        ifind_action.triggered.connect(self.configure_ifind)
        data_year_action = help_menu.addAction("设置数据年度...")
        data_year_action.triggered.connect(self._set_data_year)
        help_menu.addSeparator()
        # 一个一个导入
        import_holdings_action = help_menu.addAction("导入持仓...")
        import_holdings_action.triggered.connect(lambda: self.import_single("holdings"))
        import_canbuy_action = help_menu.addAction("导入可买名单...")
        import_canbuy_action.triggered.connect(lambda: self.import_single("can_buy_lists"))
        import_contacts_action = help_menu.addAction("导入联系人...")
        import_contacts_action.triggered.connect(lambda: self.import_single("contacts"))
        help_menu.addSeparator()
        # 批量更新
        update_data_action = help_menu.addAction("更新数据（批量）...")
        update_data_action.triggered.connect(self.update_data)
        help_menu.addSeparator()
        clear_holdings_action = help_menu.addAction("清空持仓数据...")
        clear_holdings_action.triggered.connect(lambda: self._clear_data("holdings"))
        clear_canbuy_action = help_menu.addAction("清空可买名单...")
        clear_canbuy_action.triggered.connect(lambda: self._clear_data("can_buy_lists"))
        clear_contacts_action = help_menu.addAction("清空联系人...")
        clear_contacts_action.triggered.connect(lambda: self._clear_data("contacts"))
        clear_all_action = help_menu.addAction("清空所有数据...")
        clear_all_action.triggered.connect(lambda: self._clear_data("all"))
        delete_batch_action = help_menu.addAction("删除指定导入...")
        delete_batch_action.triggered.connect(self._delete_batch_dialog)
        help_menu.addSeparator()
        about_action = help_menu.addAction("关于")
        about_action.triggered.connect(self.show_about)
        contact_action = help_menu.addAction("联系作者")
        contact_action.triggered.connect(self.show_contact)
    
    def search_bond(self):
        """Search for bond information and matching buyside companies"""
        bond_code = self.bond_code_input.text().strip()
        
        if not bond_code:
            QMessageBox.warning(self, "输入错误", "请输入债券代码")
            return
        
        if not self.data_api:
            QMessageBox.warning(
                self,
                "iFinD 未连接",
                "iFinD 未连接，无法搜索债券信息。请从帮助菜单选择「iFinD 设置」填写账号密码。"
            )
            return
        
        # Clear previous results
        self.clear_results()
        
        # Update status
        self.status_bar.showMessage(f"正在搜索债券 {bond_code}...")
        
        try:
            self.current_bond_info = self.data_api.get_bond_info(bond_code)
            if not self.current_bond_info:
                detail = getattr(self.data_api, 'get_last_bond_error', lambda: '')()
                msg = f"无法找到债券 {bond_code} 的信息。\n\n请确认 iFinD 已连接并正常运行；若已连接仍无结果，可能是该代码在 iFinD 中无数据或指标名不匹配。"
                if detail:
                    msg += f"\n\n详情：{detail}"
                QMessageBox.warning(self, "搜索错误", msg)
                self.status_bar.showMessage("搜索失败")
                return
            
            # Display bond information
            self.display_bond_info(self.current_bond_info)
            
            # Get issuer information
            issuer_name = self.current_bond_info.get('issuer_name', '')
            if not issuer_name:
                QMessageBox.warning(self, "搜索错误", f"债券 {bond_code} 缺少发行人信息")
                self.status_bar.showMessage("搜索不完整 - 缺少发行人信息")
                return
                
            self.current_issuer_info = self.data_api.get_issuer_info(issuer_name)
            if not self.current_issuer_info:
                QMessageBox.warning(self, "搜索错误", f"无法获取发行人 {issuer_name} 的信息，请确认 iFinD 已连接并正常运行。")
                self.status_bar.showMessage("搜索不完整 - 无法获取发行人信息")
                return
                
            # Add/update bond and issuer in database
            self.current_issuer_id = self.db_manager.add_bond_and_issuer(
                bond_code=bond_code,
                bond_name=self.current_bond_info.get('bond_name', ''),
                issuer_name=issuer_name,
                issuer_code=self.current_issuer_info.get('issuer_code'),
                credit_rating=self.current_issuer_info.get('credit_rating'),
                industry=self.current_issuer_info.get('industry'),
                issue_date=self.current_bond_info.get('issue_date'),
                maturity_date=self.current_bond_info.get('maturity_date'),
                coupon_rate=self.current_bond_info.get('coupon_rate')
            )
            
            # Find matching buyside companies（传入 bond_code 以支持按债券代码匹配，如发行人缺失或名称不一致）
            if self.current_issuer_id:
                self.find_matching_companies(self.current_issuer_id, bond_code=bond_code)
                self.status_bar.showMessage("搜索完成")
            else:
                QMessageBox.warning(self, "数据库错误", "无法将发行人信息保存到数据库")
                self.status_bar.showMessage("搜索失败 - 数据库错误")
                
        except Exception as e:
            QMessageBox.critical(self, "搜索错误", f"搜索过程中发生错误: {str(e)}")
            self.status_bar.showMessage("搜索失败")
            print(f"Search error: {str(e)}")
            if "iFinD" in str(e) or "数据" in str(e):
                retry = QMessageBox.question(
                    self,
                    "iFinD 连接问题",
                    "iFinD 连接出现问题。是否打开「iFinD 设置」重新配置？",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.Yes
                )
                if retry == QMessageBox.Yes:
                    self.configure_ifind()
    
    def display_bond_info(self, bond_info):
        """Display bond information in the UI"""
        self.bond_name_label.setText(bond_info.get('bond_name', '-'))
        self.issuer_name_label.setText(bond_info.get('issuer_name', '-'))
    
    def find_matching_companies(self, issuer_id, bond_code=None):
        """Find and display buyside companies that can buy bonds from this issuer"""
        companies = self.db_manager.get_can_buy_companies(issuer_id, bond_code=bond_code)
        
        if not companies:
            self.status_bar.showMessage("未找到匹配的买方公司。请通过「帮助」→「更新数据」导入持仓、可买名单 Excel。")
            # 显示一行提示，告知用户如何导入数据
            self.results_table.setRowCount(1)
            hint = QTableWidgetItem("（暂无潜在买家。请通过「帮助」→「更新数据」导入持仓、可买名单 Excel）")
            hint.setForeground(Qt.gray)
            self.results_table.setItem(0, 0, hint)
            for c in range(1, 4):
                self.results_table.setItem(0, c, QTableWidgetItem(""))
            return
        
        # Display the companies in the results table
        self.results_table.setRowCount(len(companies))
        
        # Map for match type translations
        match_type_map = {
            'both': {"text": "两者", "color": Qt.green},
            'explicit': {"text": "白名单", "color": Qt.yellow},
            'holdings': {"text": "过去持仓", "color": Qt.cyan}
        }
        
        # 债券时间默认值：优先用数据年度（导入文件名如 23年表三 → 2023），否则用债券到期/发行日年份
        bond_date_default = str(self._data_year).strip() if self._data_year else ""
        if not bond_date_default and self.current_bond_info:
            for k in ('maturity_date', 'issue_date'):
                v = self.current_bond_info.get(k)
                if v is not None and str(v).strip():
                    if hasattr(v, 'strftime'):
                        bond_date_default = v.strftime('%Y')
                    else:
                        s = str(v)[:10]
                        bond_date_default = s[:4] if len(s) >= 4 and s[:4].isdigit() else ""
                    if bond_date_default:
                        break
        # 无文档年度且无债券日期时留空，用户可手动填写（不再用当前年覆盖）

        # Disconnect signals during table population to prevent unwanted selections
        self.results_table.blockSignals(True)
        try:
            self.results_table.itemChanged.disconnect(self._on_results_cell_edited)
        except Exception:
            pass

        for row, company in enumerate(companies):
            company_name = company.get('company_name', '')
            # Company name (read-only)
            nm = QTableWidgetItem(company_name)
            nm.setFlags(nm.flags() & ~Qt.ItemIsEditable)
            self.results_table.setItem(row, 0, nm)

            # Match type (read-only)
            match_type = company.get('match_type', '')
            match_info = match_type_map.get(match_type, {"text": "未知", "color": None})
            match_type_item = QTableWidgetItem(match_info["text"])
            match_type_item.setFlags(match_type_item.flags() & ~Qt.ItemIsEditable)
            if match_info["color"]:
                match_type_item.setBackground(match_info["color"])
            self.results_table.setItem(row, 1, match_type_item)

            # 债券时间 (editable，默认数据年度或债券年份)
            key = (self.current_bond_info.get('bond_code', '') if self.current_bond_info else '', company_name)
            bond_time_val = self._bond_times.get(key, bond_date_default)
            dt = QTableWidgetItem(bond_time_val)
            self.results_table.setItem(row, 2, dt)

            # 备注 (editable)
            key = (self.current_bond_info.get('bond_code', '') if self.current_bond_info else '', company_name)
            remark = self._remarks.get(key, '')
            rm = QTableWidgetItem(remark)
            self.results_table.setItem(row, 3, rm)

        self.results_table.blockSignals(False)
        self.results_table.itemChanged.connect(self._on_results_cell_edited)
        
        # 有潜在买家但尚未点击时，在右侧提示“请点击左侧公司”
        self._show_click_company_hint()
        
        self.status_bar.showMessage(f"找到 {len(companies)} 个匹配的买方公司")
    
    def _on_results_cell_edited(self, item):
        """保存用户对债券时间(列2)、备注(列3)的编辑"""
        col = item.column()
        if col not in (2, 3):
            return
        row = item.row()
        company_item = self.results_table.item(row, 0)
        if not company_item:
            return
        company_name = company_item.text()
        bond_code = self.current_bond_info.get('bond_code', '') if self.current_bond_info else ''
        key = (bond_code, company_name)
        if col == 2:
            self._bond_times[key] = item.text().strip()
        else:
            self._remarks[key] = item.text().strip()

    def _show_click_company_hint(self):
        """在联系人区域显示：请点击左侧潜在买家中的公司"""
        try:
            if self.contacts_table.rowCount() > 0:
                self.contacts_table.setRowCount(0)
            self.contacts_table.setRowCount(1)
            hint = QTableWidgetItem("请点击左侧「潜在买家」中的一家公司，查看其联系人详情与基金持仓。")
            hint.setForeground(Qt.gray)
            self.contacts_table.setItem(0, 0, hint)
            self.contacts_table.setSpan(0, 0, 1, self.contacts_table.columnCount())
            if self.fund_holdings_table.rowCount() > 0:
                self.fund_holdings_table.setRowCount(0)
            self.fund_holdings_group.setVisible(False)
        except Exception as e:
            print(f"Error in _show_click_company_hint: {e}")
    
    def find_company_by_name(self, company_name):
        """Find a company in the database by its name using multiple search strategies"""
        if not company_name:
            return None
            
        conn = self.db_manager.get_connection()
        cursor = conn.cursor()
        company = None
        
        # Search strategies in order of precision
        search_strategies = [
            # Exact match
            ("SELECT company_id, company_name FROM buyside_companies WHERE company_name = ?", 
             (company_name,), "exact match"),
            # LIKE match with trailing spaces
            ("SELECT company_id, company_name FROM buyside_companies WHERE company_name LIKE ?", 
             (f"{company_name}%",), "LIKE match"),
            # TRIM match
            ("SELECT company_id, company_name FROM buyside_companies WHERE TRIM(company_name) = TRIM(?)", 
             (company_name,), "TRIM match"),
            # Partial name match
            ("SELECT company_id, company_name FROM buyside_companies WHERE company_name LIKE ?", 
             (f"%{company_name.split()[0] if ' ' in company_name else company_name[:min(len(company_name), 10)]}%",), 
             "partial match")
        ]
        
        # Try each strategy until we find a match
        for query, params, strategy_name in search_strategies:
            cursor.execute(query, params)
            result = cursor.fetchone()
            if result:
                company = result
                break
                
        return company
    
    def show_selected_company_details(self):
        """Show details for the selected company"""
        try:
            # Get the selected rows
            selected_rows = self.results_table.selectionModel().selectedRows()
            if not selected_rows:
                return
            
            row = selected_rows[0].row()
            company_name = self.results_table.item(row, 0).text()
            
            # Increment selection counter
            self.selection_count += 1
            current_selection_id = self.selection_count
            
            # Cancel any pending timer to avoid multiple selections
            if self.debounce_timer.isActive():
                self.debounce_timer.stop()
            
            # Calculate time since last selection
            current_time = time.time()
            
            # Store this selection pending processing
            self.pending_selection = {
                'company_name': company_name,
                'row': row,
                'selection_id': current_selection_id
            }
            
            # Update the last selection time
            self.last_selection_time = current_time
            
            # Start debounce timer (200ms delay)
            self.debounce_timer.start(200)
            
        except Exception as e:
            print(f"Error in show_selected_company_details: {str(e)}")
            traceback.print_exc()
            self.status_bar.showMessage(f"显示详情时出错: {str(e)}")
    
    def _handle_delayed_selection(self):
        """Process delayed selection after debounce timer expires"""
        if not self.pending_selection:
            return
            
        selection = self.pending_selection
        company_name = selection['company_name']
        row = selection['row']
        
        # Increment display counter for this operation
        self.display_count += 1
        
        # Clear previous contacts and holdings
        if self.contacts_table.rowCount() > 0:
            self.contacts_table.setRowCount(0)
        
        if self.fund_holdings_table.rowCount() > 0:
            self.fund_holdings_table.setRowCount(0)
            
        try:
            # Look up the company in the database
            company = self.find_company_by_name(company_name)
            
            if not company:
                self.display_message_in_table(f"无法找到公司 '{company_name}'")
                self.status_bar.showMessage(f"无法找到公司 '{company_name}'")
                return
            
            # Block all signals during display operations to prevent cascading events
            self.blockSignals(True)
            
            # Display contacts
            self.display_company_contacts(company['company_id'], company_name)
            
            # Display fund holdings if available
            self.display_fund_holdings(company['company_id'])
            
            # Restore signals
            self.blockSignals(False)
            
            # Clear pending selection as it's now processed
            self.pending_selection = None
            
        except Exception as e:
            # Restore signals in case of error
            self.blockSignals(False)
            
            print(f"Error during display: {str(e)}")
            traceback.print_exc()
            self.status_bar.showMessage(f"显示详情时出错: {str(e)}")
    
    def display_company_contacts(self, company_id, company_name):
        """Display contacts for a company"""
        try:
            # Check if table already has content
            if self.contacts_table.rowCount() > 0:
                self.contacts_table.setRowCount(0)
            
            conn = self.db_manager.get_connection()
            cursor = conn.cursor()
            
            # First, fetch ONLY contacts without any joins to avoid duplication
            contacts_query = """
            SELECT contact_id, name, position, email, phone, 
                   qt, qq, wechat, mobil, is_primary, fund_id
            FROM contacts
            WHERE company_id = ?
            ORDER BY name, is_primary DESC
            """
            
            cursor.execute(contacts_query, (company_id,))
            contacts = cursor.fetchall()
            
            if not contacts:
                self.display_message_in_table(f"未找到 '{company_name}' 的联系人")
                self.status_bar.showMessage(f"未找到 '{company_name}' 的联系人")
                return
                
            # Get all funds for this company (one query instead of many)
            funds_query = """
            SELECT fund_id, fund_name
            FROM funds
            WHERE company_id = ?
            """
            
            cursor.execute(funds_query, (company_id,))
            funds = {fund['fund_id']: fund['fund_name'] for fund in cursor.fetchall()}
            
            # Group contacts by name to handle duplicates
            contacts_by_name = {}
            for contact in contacts:
                # Skip empty names
                name = contact['name'] if 'name' in contact.keys() and contact['name'] is not None else ""
                name = str(name).strip()
                if not name:
                    continue
                    
                # Use name as key to group identical names
                if name not in contacts_by_name:
                    contacts_by_name[name] = {
                        'contact': dict(contact),  # Use the first contact as base
                        'fund_ids': set(),
                        'is_primary': contact['is_primary'] if 'is_primary' in contact.keys() else False
                    }
                # If this is a primary contact, update the is_primary flag and use this as the base
                elif contact['is_primary'] and 'is_primary' in contact.keys():
                    contacts_by_name[name]['is_primary'] = True
                    contacts_by_name[name]['contact'] = dict(contact)
                
                # Add the fund ID for this contact if it exists
                if 'fund_id' in contact.keys() and contact['fund_id'] is not None:
                    contacts_by_name[name]['fund_ids'].add(contact['fund_id'])
            
            # Create list of merged contacts
            merged_contacts = []
            for name, data in contacts_by_name.items():
                contact = data['contact']
                contact['fund_ids'] = list(data['fund_ids'])
                contact['is_primary'] = data['is_primary']
                merged_contacts.append(contact)
            
            # Now set up the display table
            self.contacts_table.setRowCount(len(merged_contacts))
            
            for row, contact in enumerate(merged_contacts):
                # Safe getter for contact fields that works with sqlite3.Row
                def get_field(field, default=''):
                    if field in contact.keys() and contact[field] is not None:
                        return str(contact[field])
                    return default
                
                # Get contact data
                contact_name = get_field('name')
                position = get_field('position')
                email = get_field('email')
                phone = get_field('phone')
                qt = get_field('qt')
                qq = get_field('qq')
                wechat = get_field('wechat')
                mobil = get_field('mobil')
                is_primary = contact.get('is_primary', False)
                
                # Add fund affiliations 
                if 'fund_ids' in contact and contact['fund_ids']:
                    fund_names = []
                    for fund_id in contact['fund_ids']:
                        if fund_id in funds:
                            fund_names.append(funds[fund_id])
                    
                    if fund_names:
                        contact_name = f"{contact_name} ({', '.join(fund_names)})"
                
                # Create table items
                name_item = QTableWidgetItem(contact_name)
                position_item = QTableWidgetItem(position)
                email_item = QTableWidgetItem(email)
                phone_item = QTableWidgetItem(phone)
                qq_item = QTableWidgetItem(qq)
                qt_item = QTableWidgetItem(qt)
                wechat_item = QTableWidgetItem(wechat)
                mobil_item = QTableWidgetItem(mobil)
                
                # Set items in table
                self.contacts_table.setItem(row, 0, name_item)
                self.contacts_table.setItem(row, 1, position_item)
                self.contacts_table.setItem(row, 2, email_item)
                self.contacts_table.setItem(row, 3, phone_item)
                self.contacts_table.setItem(row, 4, qq_item)
                self.contacts_table.setItem(row, 5, qt_item)
                self.contacts_table.setItem(row, 6, wechat_item)
                self.contacts_table.setItem(row, 7, mobil_item)
                
                # Highlight primary contact
                if is_primary:
                    for col in range(8):
                        item = self.contacts_table.item(row, col)
                        font = item.font()
                        font.setBold(True)
                        item.setFont(font)
                    
                    # Add "领导" label
                    name_item.setText(f"{name_item.text()} (领导)")
            
            # Force the table to repaint
            self.contacts_table.repaint()
            
            self.status_bar.showMessage(f"显示 '{company_name}' 的 {len(merged_contacts)} 个联系人")
            
        except Exception as e:
            print(f"Error displaying contacts: {str(e)}")
            traceback.print_exc()
            self.status_bar.showMessage(f"显示联系人时出错: {str(e)}")
            self.display_message_in_table(f"加载联系人时出错: {str(e)}")
    
    def display_message_in_table(self, message):
        """Helper method to display a message across the contacts table"""
        try:
            self.contacts_table.clearSpans()
            self.contacts_table.setRowCount(1)
            message_item = QTableWidgetItem(message)
            message_item.setTextAlignment(Qt.AlignCenter)
            self.contacts_table.setItem(0, 0, message_item)
            self.contacts_table.setSpan(0, 0, 1, 8)  # Span across all columns
        except Exception as e:
            print(f"Error in display_message_in_table: {str(e)}")
            traceback.print_exc()
    
    def display_fund_holdings(self, company_id):
        """Display fund holdings for a company（先按发行人查，若无结果再按债券代码查，避免发行人名称不一致时查不到）"""
        try:
            fund_holdings = []
            if self.current_issuer_id:
                fund_holdings = self.db_manager.get_fund_holdings(company_id, self.current_issuer_id)
            if not fund_holdings and self.current_bond_info and self.current_bond_info.get('bond_code'):
                fund_holdings = self.db_manager.get_fund_holdings_by_bond_code(
                    company_id, self.current_bond_info['bond_code']
                )
            if not fund_holdings:
                self.fund_holdings_group.setVisible(False)
                return
            
            self.fund_holdings_table.setRowCount(len(fund_holdings))
            self.fund_holdings_group.setVisible(True)
            for row, holding in enumerate(fund_holdings):
                def get_field(field, default=''):
                    return holding[field] if field in holding.keys() else default
                fund_name = get_field('fund_name')
                bond_name = get_field('bond_name')
                amount = str(get_field('amount'))
                self.fund_holdings_table.setItem(row, 0, QTableWidgetItem(fund_name))
                self.fund_holdings_table.setItem(row, 1, QTableWidgetItem(bond_name))
                self.fund_holdings_table.setItem(row, 2, QTableWidgetItem(amount))
        except Exception as e:
            print(f"Error displaying fund holdings: {e}")
            traceback.print_exc()
            self.fund_holdings_group.setVisible(False)
    
    def clear_results(self):
        """Clear all result displays"""
        # Clear bond info
        self.bond_name_label.setText("-")
        self.issuer_name_label.setText("-")
        
        # Clear results table
        self.results_table.setRowCount(0)
        
        # Clear details tables
        self.contacts_table.setRowCount(0)
        self.fund_holdings_table.setRowCount(0)
        
        # Reset current data
        self.current_issuer_id = None
        self.current_bond_info = None
        self.current_issuer_info = None
    
    def show_about(self):
        """Show the about dialog"""
        QMessageBox.about(self, "关于债券买家匹配",
                         """<h2>债券买家匹配</h2>
                         <p>版本 1.0</p>
                         <p>一个基于债券发行人匹配潜在买家的工具。</p>""")
    
    def show_contact(self):
        """Show contact information"""
        QMessageBox.information(self, "联系作者",
                            """<h3>联系作者</h3>
                            <p>微信：RealAustinSun</p>""")
    
    def load_settings(self):
        """Load application settings"""
        self._data_year = self.settings.value("data_year", "") or ""
        geometry = self.settings.value("geometry")
        if geometry:
            self.restoreGeometry(geometry)
    
    def save_settings(self):
        """Save application settings"""
        self.settings.setValue("geometry", self.saveGeometry())
        self.settings.setValue("data_year", self._data_year)
    
    def closeEvent(self, event):
        self.save_settings()
        if self.data_api:
            self.data_api.disconnect()
        self.db_manager.close_connection()
        event.accept()

    def _set_data_year(self):
        """手动设置数据年度（债券时间默认显示该值，如导入文件名无法解析年度时可在此设置）"""
        from PyQt5.QtWidgets import QInputDialog
        year, ok = QInputDialog.getText(
            self, "设置数据年度",
            "请输入数据年度（如 2023），将作为债券时间列的默认显示：",
            QLineEdit.Normal,
            self._data_year or str(time.localtime().tm_year - 1)
        )
        if ok and year:
            y = str(year).strip()
            if y.isdigit() and len(y) == 4:
                self._data_year = y
                self.save_settings()
                QMessageBox.information(self, "设置成功", f"数据年度已设为 {y}，下次搜索时将作为债券时间默认值。")
            else:
                QMessageBox.warning(self, "输入错误", "请输入 4 位年份（如 2023）")

    def _clear_data(self, data_type):
        """清空错误导入的数据，需确认"""
        titles = {
            "holdings": ("清空持仓数据", "确定清空所有持仓数据？导入错误的表格后可用此功能删除。此操作不可恢复。"),
            "can_buy_lists": ("清空可买名单", "确定清空所有可买名单？此操作不可恢复。"),
            "contacts": ("清空联系人", "确定清空所有联系人？此操作不可恢复。"),
            "all": ("清空所有数据", "确定清空所有数据（持仓、可买名单、联系人、公司等）？此操作不可恢复。"),
        }
        title, msg = titles.get(data_type, ("确认", "确定执行？"))
        r = QMessageBox.warning(self, title, msg, QMessageBox.Ok | QMessageBox.Cancel, QMessageBox.Cancel)
        if r != QMessageBox.Ok:
            return
        n = self.db_manager.clear_data(data_type)
        QMessageBox.information(self, "完成", f"已删除 {n} 条记录。")
        self.status_bar.showMessage("数据已清空")
        self.clear_results()

    def _delete_batch_dialog(self):
        """打开删除指定导入批次对话框"""
        dialog = DeleteBatchDialog(self, self.db_manager)
        if dialog.exec_() == QDialog.Accepted and dialog.deleted_count is not None:
            QMessageBox.information(self, "完成", f"已删除该批次的 {dialog.deleted_count} 条记录。")
            self.status_bar.showMessage("已删除指定导入")
            self.clear_results()

    def configure_ifind(self):
        """打开 iFinD 设置对话框，保存后重新初始化 API"""
        dialog = iFinDSettingsDialog(self, self.settings)
        if dialog.exec_() == QDialog.Accepted:
            self.init_data_api()
            # 保存后始终提示账号已保存；若已连接则一并提示
            if self.data_api and self.data_api.connected:
                QMessageBox.information(self, "iFinD 设置", "账号已保存，iFinD 已连接。")
            else:
                QMessageBox.information(self, "iFinD 设置", "账号已保存。若仍无法连接，请检查网络与数据接口权限。")

    def find_data_file(self, folder_path, base_name):
        """在文件夹中查找数据文件，支持英文/中文名（如 holdings/持仓、can_buy_lists/可买名单、contacts/联系人）"""
        name_variants = {
            'holdings': ['holdings', '持仓'],
            'can_buy_lists': ['can_buy', 'can_buy_lists', '可买名单', '白名单'],
            'contacts': ['contacts', '联系人'],
        }
        patterns = name_variants.get(base_name, [base_name])
        for ext in ['.xlsx', '.xls', '.csv']:
            exact_path = os.path.join(folder_path, f"sample_{base_name}{ext}")
            if os.path.exists(exact_path):
                return exact_path
            for p in patterns:
                if p == base_name:
                    continue
                exact_path = os.path.join(folder_path, f"sample_{p}{ext}")
                if os.path.exists(exact_path):
                    return exact_path
        for ext in ['.xlsx', '.xls', '.csv']:
            for p in patterns:
                pattern = os.path.join(folder_path, f"*{p}*{ext}")
                matches = glob.glob(pattern)
                matches = [m for m in matches if not os.path.basename(m).startswith('~$')]
                if matches:
                    return matches[0]
        return None

    def update_data(self):
        """导入数据：支持选择文件夹（自动识别 Excel）或逐个选择三个 Excel 文件"""
        choice = QMessageBox.question(
            self,
            "更新数据",
            "请选择导入方式：\n\n"
            "【是】选择文件夹（自动查找持仓、可买名单、联系人 Excel）\n"
            "【否】逐个选择三个 Excel 文件（持仓 → 可买名单 → 联系人）",
            QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel,
            QMessageBox.Yes
        )
        if choice == QMessageBox.Cancel:
            return
        if choice == QMessageBox.Yes:
            self._update_data_from_folder()
        else:
            self._update_data_from_files()

    def import_single(self, data_type):
        """只导入某一类数据：持仓 / 可买名单 / 联系人（任选一个 Excel 文件）"""
        titles = {"holdings": "选择【持仓】Excel 或 CSV（50 万+ 行推荐 CSV）", "can_buy_lists": "选择【可买名单】Excel 或 CSV", "contacts": "选择【联系人】Excel 或 CSV"}
        start_dir = os.path.abspath("data") if os.path.isdir("data") else ""
        filter_excel = "数据文件 (*.xlsx *.xls *.csv);;Excel (*.xlsx *.xls);;CSV (*.csv);;所有文件 (*.*)"
        path, _ = QFileDialog.getOpenFileName(self, titles[data_type], start_dir, filter_excel)
        if not path:
            return
        self.status_bar.showMessage(f"正在导入 {titles[data_type]}...")
        try:
            if data_type == "holdings":
                success, counts = self.db_manager.import_from_excel(path, None, None)
                n = counts.get("holdings", 0)
            elif data_type == "can_buy_lists":
                success, counts = self.db_manager.import_from_excel(None, path, None)
                n = counts.get("can_buy_lists", 0)
            else:
                success, counts = self.db_manager.import_from_excel(None, None, path)
                n = counts.get("contacts", 0)
            if success and n > 0:
                if data_type == "holdings":
                    yr = _parse_year_from_filename(os.path.basename(path))
                    if yr:
                        self._data_year = yr
                QMessageBox.information(self, "导入成功", f"已导入 {n} 条记录：{os.path.basename(path)}")
                self.status_bar.showMessage("导入成功")
            elif success and n == 0:
                req = {"holdings": "公司名称、发行人、债券代码", "can_buy_lists": "公司名称、发行人", "contacts": "公司名称（可选：公司类型、发行人、姓名）"}
                hint = "联系人支持仅含「公司类型、公司名称、发行人」的表格（无姓名时自动占位）。" if data_type == "contacts" else "支持中文表头：公司名称/机构名称、发行人、债券代码、姓名/联系人 等。"
                QMessageBox.warning(
                    self, "未导入任何数据",
                    f"Excel 已读取但未插入任何记录。\n\n请检查：\n"
                    f"1) 是否有数据行（表头下方有内容）\n"
                    f"2) 表头列名是否包含：{req.get(data_type, '')}\n"
                    f"{hint}"
                )
                self.status_bar.showMessage("未导入任何数据")
            else:
                QMessageBox.critical(self, "导入失败", "导入时出错，请查看控制台。")
                self.status_bar.showMessage("导入失败")
        except Exception as e:
            QMessageBox.critical(self, "导入错误", f"导入时发生错误：\n{str(e)}")
            self.status_bar.showMessage("导入失败")
            traceback.print_exc()

    def _update_data_from_folder(self):
        """从选择的文件夹中查找并导入三个 Excel"""
        default_dir = os.path.abspath("data") if os.path.isdir("data") else ""
        folder_path = QFileDialog.getExistingDirectory(
            self, "选择包含数据文件的文件夹", default_dir, QFileDialog.ShowDirsOnly
        )
        if not folder_path:
            return
        required_files = ['holdings', 'can_buy_lists', 'contacts']
        found_files = {}
        missing_files = []
        for base_name in required_files:
            file_path = self.find_data_file(folder_path, base_name)
            if file_path:
                found_files[base_name] = file_path
            else:
                missing_files.append(base_name)
        if missing_files:
            QMessageBox.warning(
                self,
                "文件缺失",
                f"在所选文件夹中未找到以下类型的 Excel：\n"
                f"{', '.join(missing_files)}\n\n"
                f"支持的文件名示例：*holdings* 或 *持仓*、*can_buy* 或 *可买名单*、*contacts* 或 *联系人*"
            )
            return
        self._do_import(found_files['holdings'], found_files['can_buy_lists'], found_files['contacts'])

    def _update_data_from_files(self):
        """让用户依次选择三个 Excel 文件（支持任意文件名）"""
        start_dir = os.path.abspath("data") if os.path.isdir("data") else ""
        filter_excel = "数据文件 (*.xlsx *.xls *.csv);;Excel (*.xlsx *.xls);;CSV (*.csv);;所有文件 (*.*)"
        holdings_path, _ = QFileDialog.getOpenFileName(
            self, "选择【持仓】Excel", start_dir, filter_excel
        )
        if not holdings_path:
            return
        can_buy_path, _ = QFileDialog.getOpenFileName(
            self, "选择【可买名单】Excel", os.path.dirname(holdings_path), filter_excel
        )
        if not can_buy_path:
            return
        contacts_path, _ = QFileDialog.getOpenFileName(
            self, "选择【联系人】Excel", os.path.dirname(can_buy_path), filter_excel
        )
        if not contacts_path:
            return
        self._do_import(holdings_path, can_buy_path, contacts_path)

    def _do_import(self, holdings_path, can_buy_path, contacts_path):
        """执行导入并提示结果"""
        file_list = "\n".join([
            f"• 持仓: {os.path.basename(holdings_path)}",
            f"• 可买名单: {os.path.basename(can_buy_path)}",
            f"• 联系人: {os.path.basename(contacts_path)}",
        ])
        confirm = QMessageBox.question(
            self,
            "确认导入",
            f"将导入以下文件：\n\n{file_list}\n\n是否导入？将替换现有数据。",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if confirm == QMessageBox.No:
            return
        self.status_bar.showMessage("正在导入数据...")
        try:
            success, counts = self.db_manager.import_from_excel(holdings_path, can_buy_path, contacts_path)
            if success:
                yr = _parse_year_from_filename(os.path.basename(holdings_path))
                if yr:
                    self._data_year = yr
                h, c, t = counts.get("holdings", 0), counts.get("can_buy_lists", 0), counts.get("contacts", 0)
                msg = f"导入完成。持仓 {h} 条、可买名单 {c} 条、联系人 {t} 条。"
                if h == 0 and c == 0 and t == 0:
                    msg += "\n\n未插入任何记录，请检查 Excel 是否有数据行且表头列名正确（支持中文：公司名称、发行人、债券代码、姓名等）。"
                    QMessageBox.warning(self, "导入完成", msg)
                else:
                    QMessageBox.information(self, "导入成功", msg)
                self.status_bar.showMessage("数据导入完成")
            else:
                QMessageBox.critical(self, "导入失败", "导入时出错，请查看控制台。")
                self.status_bar.showMessage("数据导入失败")
        except Exception as e:
            QMessageBox.critical(self, "导入错误", f"导入时发生错误：\n{str(e)}")
            self.status_bar.showMessage("数据导入失败")
            traceback.print_exc()


if __name__ == "__main__":
    # Create directories if they don't exist
    os.makedirs('database', exist_ok=True)
    
    # Create the application
    app = QApplication(sys.argv)
    
    # Set application style
    app.setStyle("Fusion")
    
    # Create and show the main window
    window = BondBuyerMatchApp()
    window.show()
    
    # Run the application
    sys.exit(app.exec_()) 