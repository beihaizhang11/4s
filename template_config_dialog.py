"""
模板配置对话框
用于添加和编辑模板配置
"""
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, 
    QPushButton, QTableWidget, QTableWidgetItem, QFileDialog,
    QMessageBox, QGroupBox, QTextEdit, QHeaderView
)
from PyQt6.QtCore import Qt
from typing import Dict, List


class TemplateConfigDialog(QDialog):
    """模板配置对话框"""
    
    def __init__(self, parent=None, template_name: str = "", template_config: Dict = None):
        super().__init__(parent)
        self.template_name = template_name
        self.template_config = template_config if template_config else {}
        self.is_edit_mode = bool(template_name)
        
        self.init_ui()
        self.load_config()
    
    def init_ui(self):
        """初始化UI"""
        self.setWindowTitle("编辑模板配置" if self.is_edit_mode else "添加模板配置")
        self.setMinimumSize(900, 700)
        
        layout = QVBoxLayout()
        
        # 模板名称
        name_layout = QHBoxLayout()
        name_layout.addWidget(QLabel("模板名称:"))
        self.name_edit = QLineEdit()
        self.name_edit.setText(self.template_name)
        if self.is_edit_mode:
            self.name_edit.setReadOnly(True)
        name_layout.addWidget(self.name_edit)
        layout.addLayout(name_layout)
        
        # Excel模板路径
        excel_layout = QHBoxLayout()
        excel_layout.addWidget(QLabel("Excel模板:"))
        self.excel_path_edit = QLineEdit()
        excel_layout.addWidget(self.excel_path_edit)
        self.excel_browse_btn = QPushButton("浏览...")
        self.excel_browse_btn.clicked.connect(self.browse_excel_template)
        excel_layout.addWidget(self.excel_browse_btn)
        layout.addLayout(excel_layout)
        
        # 字段映射表
        mapping_group = QGroupBox("字段映射配置")
        mapping_layout = QVBoxLayout()
        
        # 说明文字
        help_text = QLabel(
            "配置数据库字段到Excel单元格的映射。可以将一个字段映射到多个单元格。\n"
            "单元格格式: A1, B2, Sheet1!C3 等。多个单元格用逗号分隔。"
        )
        help_text.setWordWrap(True)
        mapping_layout.addWidget(help_text)
        
        # 字段映射表格
        self.mapping_table = QTableWidget()
        self.mapping_table.setColumnCount(2)
        self.mapping_table.setHorizontalHeaderLabels(["数据库字段", "Excel单元格 (用逗号分隔)"])
        self.mapping_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.mapping_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        
        # 默认字段
        default_fields = [
            "task_id", "sender_name", "sender_mail", "sender_phone",
            "VIN", "task_dept", "sent_time", "task_description", "bg_description", "serial_number"
        ]
        self.mapping_table.setRowCount(len(default_fields))
        for i, field in enumerate(default_fields):
            self.mapping_table.setItem(i, 0, QTableWidgetItem(field))
            self.mapping_table.setItem(i, 1, QTableWidgetItem(""))
        
        mapping_layout.addWidget(self.mapping_table)
        
        # 添加/删除行按钮
        mapping_btn_layout = QHBoxLayout()
        self.add_mapping_btn = QPushButton("添加字段")
        self.add_mapping_btn.clicked.connect(self.add_mapping_row)
        mapping_btn_layout.addWidget(self.add_mapping_btn)
        
        self.remove_mapping_btn = QPushButton("删除选中行")
        self.remove_mapping_btn.clicked.connect(self.remove_mapping_row)
        mapping_btn_layout.addWidget(self.remove_mapping_btn)
        mapping_btn_layout.addStretch()
        
        mapping_layout.addLayout(mapping_btn_layout)
        mapping_group.setLayout(mapping_layout)
        layout.addWidget(mapping_group)
        
        # 邮件配置
        email_group = QGroupBox("邮件配置")
        email_layout = QVBoxLayout()
        
        # TO列表第一个固定收件人
        to_first_layout = QHBoxLayout()
        to_first_layout.addWidget(QLabel("TO首位收件人:"))
        self.to_first_edit = QLineEdit()
        self.to_first_edit.setPlaceholderText("邮箱地址")
        to_first_layout.addWidget(self.to_first_edit)
        email_layout.addLayout(to_first_layout)
        
        # TO列表固定收件人（在动态收件人之后）
        email_layout.addWidget(QLabel("TO固定收件人列表 (每行一个邮箱):"))
        self.to_fixed_edit = QTextEdit()
        self.to_fixed_edit.setMaximumHeight(80)
        self.to_fixed_edit.setPlaceholderText("在动态收件人之后添加的固定收件人\n每行一个邮箱地址")
        email_layout.addWidget(self.to_fixed_edit)
        
        # CC列表
        email_layout.addWidget(QLabel("CC抄送人列表 (每行一个邮箱):"))
        self.cc_edit = QTextEdit()
        self.cc_edit.setMaximumHeight(80)
        self.cc_edit.setPlaceholderText("每行一个邮箱地址")
        email_layout.addWidget(self.cc_edit)
        
        email_group.setLayout(email_layout)
        layout.addWidget(email_group)
        
        # 按钮
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        
        self.save_btn = QPushButton("保存")
        self.save_btn.clicked.connect(self.save_config)
        btn_layout.addWidget(self.save_btn)
        
        self.cancel_btn = QPushButton("取消")
        self.cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(self.cancel_btn)
        
        layout.addLayout(btn_layout)
        
        self.setLayout(layout)
    
    def load_config(self):
        """加载配置到界面"""
        if not self.template_config:
            return
        
        # 加载Excel路径
        self.excel_path_edit.setText(self.template_config.get("excel_template_path", ""))
        
        # 加载字段映射
        field_mappings = self.template_config.get("field_mappings", {})
        for i in range(self.mapping_table.rowCount()):
            field_item = self.mapping_table.item(i, 0)
            if field_item:
                field_name = field_item.text()
                if field_name in field_mappings:
                    cells = field_mappings[field_name]
                    self.mapping_table.setItem(i, 1, QTableWidgetItem(", ".join(cells)))
        
        # 加载邮件配置
        self.to_first_edit.setText(self.template_config.get("email_to_first", ""))
        
        to_fixed = self.template_config.get("email_to_fixed", [])
        self.to_fixed_edit.setText("\n".join(to_fixed))
        
        cc = self.template_config.get("email_cc", [])
        self.cc_edit.setText("\n".join(cc))
    
    def browse_excel_template(self):
        """浏览Excel模板文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择Excel模板文件",
            "",
            "Excel文件 (*.xlsx *.xls)"
        )
        if file_path:
            self.excel_path_edit.setText(file_path)
    
    def add_mapping_row(self):
        """添加字段映射行"""
        row = self.mapping_table.rowCount()
        self.mapping_table.insertRow(row)
        self.mapping_table.setItem(row, 0, QTableWidgetItem(""))
        self.mapping_table.setItem(row, 1, QTableWidgetItem(""))
    
    def remove_mapping_row(self):
        """删除选中的字段映射行"""
        current_row = self.mapping_table.currentRow()
        if current_row >= 0:
            self.mapping_table.removeRow(current_row)
    
    def save_config(self):
        """保存配置"""
        # 验证模板名称
        template_name = self.name_edit.text().strip()
        if not template_name:
            QMessageBox.warning(self, "错误", "请输入模板名称")
            return
        
        # 验证Excel路径
        excel_path = self.excel_path_edit.text().strip()
        if not excel_path:
            QMessageBox.warning(self, "错误", "请选择Excel模板文件")
            return
        
        # 获取字段映射
        field_mappings = {}
        for i in range(self.mapping_table.rowCount()):
            field_item = self.mapping_table.item(i, 0)
            cells_item = self.mapping_table.item(i, 1)
            
            if field_item and cells_item:
                field_name = field_item.text().strip()
                cells_text = cells_item.text().strip()
                
                if field_name and cells_text:
                    # 分割单元格列表
                    cells = [c.strip() for c in cells_text.split(",") if c.strip()]
                    if cells:
                        field_mappings[field_name] = cells
        
        # 获取邮件配置
        email_to_first = self.to_first_edit.text().strip()
        
        to_fixed_text = self.to_fixed_edit.toPlainText().strip()
        email_to_fixed = [line.strip() for line in to_fixed_text.split("\n") if line.strip()]
        
        cc_text = self.cc_edit.toPlainText().strip()
        email_cc = [line.strip() for line in cc_text.split("\n") if line.strip()]
        
        # 构建配置字典
        self.result_config = {
            "template_name": template_name,
            "excel_template_path": excel_path,
            "field_mappings": field_mappings,
            "email_to_first": email_to_first,
            "email_to_fixed": email_to_fixed,
            "email_cc": email_cc
        }
        
        self.accept()
    
    def get_config(self) -> Dict:
        """获取配置结果"""
        return getattr(self, 'result_config', None)
