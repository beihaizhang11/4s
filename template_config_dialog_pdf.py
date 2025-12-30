"""
模板配置对话框 - PDF版本
简化配置，不需要Excel模板和字段映射
"""
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, 
    QPushButton, QMessageBox, QGroupBox, QTextEdit
)
from PyQt6.QtCore import Qt
from typing import Dict


class TemplateConfigDialogPDF(QDialog):
    """模板配置对话框 - PDF版本"""
    
    def __init__(self, parent=None, template_name: str = "", template_config: Dict = None):
        super().__init__(parent)
        self.template_name = template_name
        self.template_config = template_config if template_config else {}
        self.is_edit_mode = bool(template_name)
        
        self.init_ui()
        self.load_config()
    
    def init_ui(self):
        """初始化UI"""
        self.setWindowTitle("编辑模板配置 (PDF版本)" if self.is_edit_mode else "添加模板配置 (PDF版本)")
        self.setMinimumSize(700, 600)
        
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
        
        # 邮件配置
        email_group = QGroupBox("邮件配置")
        email_layout = QVBoxLayout()
        
        # TO列表首位收件人组
        email_layout.addWidget(QLabel("TO首位收件人列表 (每行一个邮箱):"))
        self.to_first_edit = QTextEdit()
        self.to_first_edit.setMaximumHeight(80)
        self.to_first_edit.setPlaceholderText("在最前面的固定收件人\n每行一个邮箱地址")
        email_layout.addWidget(self.to_first_edit)
        
        # TO列表固定收件人
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
        
        # 邮件正文模板
        body_group = QGroupBox("邮件正文模板")
        body_layout = QVBoxLayout()
        
        body_help = QLabel(
            "可用变量：{task_description}, {bg_description}, {task_ids}\n"
            "多个Task ID时，{task_ids}会自动替换为所有Task ID"
        )
        body_help.setWordWrap(True)
        body_layout.addWidget(body_help)
        
        self.body_template_edit = QTextEdit()
        self.body_template_edit.setPlaceholderText(
            "邮件正文模板，支持变量替换\n"
            "例如：\n"
            "hello\n"
            "请参考附件服务 {task_description}\n"
            "\n"
            "背景：{bg_description}\n"
            "\n"
            "请2个工作日内确认是否可以提供服务。"
        )
        body_layout.addWidget(self.body_template_edit)
        
        body_group.setLayout(body_layout)
        layout.addWidget(body_group)
        
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
        
        # 加载邮件配置
        to_first = self.template_config.get("email_to_first", [])
        if isinstance(to_first, str):
            to_first = [to_first] if to_first else []
        self.to_first_edit.setText("\n".join(to_first))
        
        to_fixed = self.template_config.get("email_to_fixed", [])
        self.to_fixed_edit.setText("\n".join(to_fixed))
        
        cc = self.template_config.get("email_cc", [])
        self.cc_edit.setText("\n".join(cc))
        
        # 加载邮件正文模板
        body_template = self.template_config.get("email_body_template", "")
        if not body_template:
            body_template = """hello
请参考附件服务 {task_description}

背景：{bg_description}

请2个工作日内确认是否可以提供服务，如果可以，请补充报价信息回复询价单。
预估到货时间，并附上报价依据截图。"""
        self.body_template_edit.setText(body_template)
    
    def save_config(self):
        """保存配置"""
        # 验证模板名称
        template_name = self.name_edit.text().strip()
        if not template_name:
            QMessageBox.warning(self, "错误", "请输入模板名称")
            return
        
        # 获取邮件配置
        to_first_text = self.to_first_edit.toPlainText().strip()
        email_to_first = [line.strip() for line in to_first_text.split("\n") if line.strip()]
        
        to_fixed_text = self.to_fixed_edit.toPlainText().strip()
        email_to_fixed = [line.strip() for line in to_fixed_text.split("\n") if line.strip()]
        
        cc_text = self.cc_edit.toPlainText().strip()
        email_cc = [line.strip() for line in cc_text.split("\n") if line.strip()]
        
        # 获取邮件正文模板
        email_body_template = self.body_template_edit.toPlainText().strip()
        if not email_body_template:
            QMessageBox.warning(self, "错误", "请输入邮件正文模板")
            return
        
        # 构建配置字典（PDF版本不需要excel_template_path和field_mappings）
        self.result_config = {
            "template_name": template_name,
            "email_to_first": email_to_first,
            "email_to_fixed": email_to_fixed,
            "email_cc": email_cc,
            "email_body_template": email_body_template
        }
        
        self.accept()
    
    def get_config(self) -> Dict:
        """获取配置结果"""
        return getattr(self, 'result_config', None)
