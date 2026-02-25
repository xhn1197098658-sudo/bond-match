import sys
import os
from pathlib import Path
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
                QMessageBox.warning(
                    self,
                    "iFinD 未连接",
                    "无法连接 iFinD。请在「帮助 → iFinD 设置」中填写同花顺数据接口账号与密码，\n"
                    "并确保已安装：pip install iFinDAPI"
                )
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
        self.results_table.setColumnCount(2)
        self.results_table.setHorizontalHeaderLabels(["公司名称", "匹配类型"])
        self.results_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.results_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.results_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.results_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
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
        
        # Update data option
        update_data_action = help_menu.addAction("更新数据")
        update_data_action.triggered.connect(self.update_data)
        
        # About option
        about_action = help_menu.addAction("关于")
        about_action.triggered.connect(self.show_about)
        
        # Contact author option
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
                QMessageBox.warning(self, "搜索错误", f"无法找到债券 {bond_code} 的信息，请确认 iFinD 已连接并正常运行。")
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
            
            # Find matching buyside companies
            if self.current_issuer_id:
                self.find_matching_companies(self.current_issuer_id)
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
    
    def find_matching_companies(self, issuer_id):
        """Find and display buyside companies that can buy bonds from this issuer"""
        companies = self.db_manager.get_can_buy_companies(issuer_id)
        
        if not companies:
            self.status_bar.showMessage("未找到匹配的买方公司")
            return
        
        # Display the companies in the results table
        self.results_table.setRowCount(len(companies))
        
        # Map for match type translations
        match_type_map = {
            'both': {"text": "两者", "color": Qt.green},
            'explicit': {"text": "白名单", "color": Qt.yellow},
            'holdings': {"text": "过去持仓", "color": Qt.cyan}
        }
        
        # Disconnect signals during table population to prevent unwanted selections
        old_state = self.results_table.blockSignals(True)
        
        for row, company in enumerate(companies):
            # Company name
            self.results_table.setItem(row, 0, QTableWidgetItem(company.get('company_name', '')))
            
            # Match type
            match_type = company.get('match_type', '')
            match_info = match_type_map.get(match_type, {"text": "未知", "color": None})
            
            match_type_item = QTableWidgetItem(match_info["text"])
            if match_info["color"]:
                match_type_item.setBackground(match_info["color"])
            
            self.results_table.setItem(row, 1, match_type_item)
        
        # Restore signals
        self.results_table.blockSignals(old_state)
        
        self.status_bar.showMessage(f"找到 {len(companies)} 个匹配的买方公司")
    
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
        """Display fund holdings for a company"""
        if not self.current_issuer_id:
            self.fund_holdings_group.setVisible(False)
            return
            
        try:
            fund_holdings = self.db_manager.get_fund_holdings(company_id, self.current_issuer_id)
            
            if fund_holdings:
                self.fund_holdings_table.setRowCount(len(fund_holdings))
                self.fund_holdings_group.setVisible(True)
                
                for row, holding in enumerate(fund_holdings):
                    # Access sqlite3.Row objects safely
                    def get_field(field, default=''):
                        return holding[field] if field in holding.keys() else default
                    
                    fund_name = get_field('fund_name')
                    bond_name = get_field('bond_name')
                    amount = str(get_field('amount'))
                    
                    self.fund_holdings_table.setItem(row, 0, QTableWidgetItem(fund_name))
                    self.fund_holdings_table.setItem(row, 1, QTableWidgetItem(bond_name))
                    self.fund_holdings_table.setItem(row, 2, QTableWidgetItem(amount))
            else:
                self.fund_holdings_group.setVisible(False)
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
        geometry = self.settings.value("geometry")
        if geometry:
            self.restoreGeometry(geometry)
    
    def save_settings(self):
        """Save application settings"""
        self.settings.setValue("geometry", self.saveGeometry())
    
    def closeEvent(self, event):
        self.save_settings()
        if self.data_api:
            self.data_api.disconnect()
        self.db_manager.close_connection()
        event.accept()

    def configure_ifind(self):
        """打开 iFinD 设置对话框，保存后重新初始化 API"""
        dialog = iFinDSettingsDialog(self, self.settings)
        if dialog.exec_() == QDialog.Accepted:
            self.init_data_api()

    def find_data_file(self, folder_path, base_name):
        """Find data file in folder based on name pattern"""
        # Check for exact files first
        for ext in ['.xlsx', '.xls']:
            exact_path = os.path.join(folder_path, f"sample_{base_name}{ext}")
            if os.path.exists(exact_path):
                return exact_path
                
        # Search for files containing the base name
        for ext in ['.xlsx', '.xls']:
            pattern = os.path.join(folder_path, f"*{base_name}*{ext}")
            matches = glob.glob(pattern)
            if matches:
                return matches[0]  # Return first match
                
        return None

    def update_data(self):
        """Allow user to select a folder containing data files and import them to the database"""
        # Open folder selection dialog
        folder_path = QFileDialog.getExistingDirectory(self, "选择包含数据文件的文件夹", "", QFileDialog.ShowDirsOnly)
        
        if not folder_path:
            return  # User canceled
        
        # Check if the folder contains the required files
        required_files = ['holdings', 'can_buy_lists', 'contacts']
        found_files = {}
        missing_files = []
        
        for base_name in required_files:
            file_path = self.find_data_file(folder_path, base_name)
            if file_path:
                found_files[base_name] = file_path
            else:
                missing_files.append(base_name)
        
        # Check if all required files were found
        if missing_files:
            QMessageBox.warning(
                self, 
                "文件缺失", 
                f"以下文件找不到: {', '.join([f'sample_{f}.xlsx' for f in missing_files])}\n"
                f"请确保文件夹包含所有必需的数据文件。"
            )
            return
        
        # Confirm import
        file_list = "\n".join([f"• {os.path.basename(path)}" for path in found_files.values()])
        confirm = QMessageBox.question(
            self,
            "确认导入",
            f"找到以下文件:\n{file_list}\n\n"
            f"是否将这些数据导入数据库？这将替换现有数据。",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if confirm == QMessageBox.No:
            return
        
        # Update status
        self.status_bar.showMessage("正在导入数据...")
        
        # Import data
        try:
            success = self.db_manager.import_from_excel(
                found_files['holdings'],
                found_files['can_buy_lists'],
                found_files['contacts']
            )
            
            if success:
                QMessageBox.information(self, "导入成功", "数据已成功导入数据库。")
                self.status_bar.showMessage("数据导入成功")
            else:
                QMessageBox.critical(self, "导入失败", "导入数据时出错，请查看控制台获取详细信息。")
                self.status_bar.showMessage("数据导入失败")
        except Exception as e:
            QMessageBox.critical(self, "导入错误", f"导入数据时发生错误:\n{str(e)}")
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