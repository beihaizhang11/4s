"""
主窗口界面 - PDF变体版本
不使用Excel模板，扫描PDF文件并作为邮件附件
"""
import os
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
    QLineEdit, QPushButton, QComboBox, QTextEdit, QMessageBox,
    QFileDialog, QGroupBox, QListWidget
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from config_manager import ConfigManager
from database_manager import DatabaseManager
from email_manager import EmailManager
from pdf_scanner import PDFScanner


class EmailGeneratorThread(QThread):
    """邮件生成线程 - PDF版本"""
    
    finished = pyqtSignal(bool, str)  # 完成信号
    progress = pyqtSignal(str)  # 进度信号
    
    def __init__(self, task_id, template_name, config_manager):
        super().__init__()
        self.task_id = task_id
        self.template_name = template_name
        self.config_manager = config_manager
    
    def run(self):
        """执行邮件生成流程"""
        try:
            # 解析task_id
            task_ids = [tid.strip() for tid in self.task_id.split(',') if tid.strip()]
            if not task_ids:
                self.finished.emit(False, "Task ID不能为空")
                return
            
            is_batch = len(task_ids) > 1
            if is_batch:
                self.progress.emit(f"检测到批量处理模式，共 {len(task_ids)} 个Task ID")
            
            # 1. 复制数据库到本地
            self.progress.emit("正在复制数据库到本地...")
            db_path = self.config_manager.get_database_path()
            local_db_path = self.config_manager.get_local_database_path()
            
            if not db_path:
                self.finished.emit(False, "未配置数据库路径，请在设置中配置")
                return
            
            db_manager = DatabaseManager(db_path, local_db_path)
            if not db_manager.copy_database_to_local():
                self.finished.emit(False, "复制数据库失败")
                return
            
            # 2. 连接数据库并查询数据
            self.progress.emit("正在查询数据库...")
            if not db_manager.connect():
                self.finished.emit(False, "连接数据库失败")
                return
            
            # 查询数据
            if is_batch:
                task_data = db_manager.prepare_data_for_excel_batch(task_ids)
                staff_emails = db_manager.get_staff_emails_by_task_ids(task_ids)
            else:
                task_data = db_manager.prepare_data_for_excel(task_ids[0])
                staff_emails = db_manager.get_staff_emails_by_task_id(task_ids[0])
            
            if not task_data:
                db_manager.disconnect()
                self.finished.emit(False, f"未找到Task ID数据")
                return
            
            # 生成Task ID显示文本
            if is_batch:
                task_id_display = ', '.join(task_ids)
            else:
                task_id_display = task_ids[0]
            
            db_manager.disconnect()
            
            # 3. 扫描PDF文件
            self.progress.emit("正在扫描PDF文件...")
            pdf_directory = self.config_manager.get_pdf_directory()
            
            if not pdf_directory:
                self.finished.emit(False, "未配置PDF文件目录，请在设置中配置")
                return
            
            if not PDFScanner.validate_directory(pdf_directory):
                self.finished.emit(False, f"PDF目录无效或无法访问: {pdf_directory}")
                return
            
            pdf_files = PDFScanner.scan_pdf_files(pdf_directory)
            
            if not pdf_files:
                self.finished.emit(False, f"在目录 {pdf_directory} 中未找到PDF文件")
                return
            
            self.progress.emit(f"找到 {len(pdf_files)} 个PDF文件")
            
            # 4. 获取模板配置
            template_config = self.config_manager.get_template(self.template_name)
            if not template_config:
                self.finished.emit(False, f"未找到模板配置: {self.template_name}")
                return
            
            # 5. 创建Outlook邮件草稿
            self.progress.emit("正在创建Outlook邮件草稿...")
            email_manager = EmailManager()
            
            # 构建收件人列表
            to_recipients = []
            
            # 首位固定收件人组
            to_first = template_config.get("email_to_first", [])
            if isinstance(to_first, str):
                to_first = [to_first] if to_first else []
            to_recipients.extend(to_first)
            
            # 动态收件人
            sender_mail = task_data.get('sender_mail', '')
            if sender_mail and sender_mail.strip():
                to_recipients.append(sender_mail.strip())
            to_recipients.extend(staff_emails)
            
            # 末尾固定收件人组
            to_fixed = template_config.get("email_to_fixed", [])
            to_recipients.extend(to_fixed)
            
            # 抄送人列表
            cc_recipients = template_config.get("email_cc", [])
            
            # 邮件主题 - PDF版本格式
            subject = f"已收货 - Audi China 汽车维修服务询价：Task ID: {task_id_display}"
            
            # 邮件正文
            task_description = task_data.get('task_description', '')
            bg_description = task_data.get('bg_description', '')
            
            email_body_template = template_config.get("email_body_template", "")
            if not email_body_template:
                email_body_template = """hello
请参考附件服务 {task_description}

背景：{bg_description}

请2个工作日内确认是否可以提供服务，如果可以，请补充报价信息回复询价单。
预估到货时间，并附上报价依据截图。"""
            
            body = EmailManager.generate_email_body_from_template(
                email_body_template,
                task_description,
                bg_description,
                task_ids if is_batch else None
            )
            
            # 附件列表：所有PDF文件
            attachments = pdf_files.copy()
            
            # 添加教程附件
            tutorial_path = self.config_manager.get_tutorial_attachment_path()
            if tutorial_path and os.path.exists(tutorial_path):
                attachments.append(tutorial_path)
            
            if not email_manager.create_draft_email(
                subject, to_recipients, cc_recipients, body, attachments
            ):
                self.finished.emit(False, "创建邮件草稿失败")
                return
            
            # 成功
            message = f"邮件草稿创建成功！\n\n已添加 {len(pdf_files)} 个PDF附件\n\nOutlook邮件草稿已打开。"
            self.finished.emit(True, message)
            
        except Exception as e:
            self.finished.emit(False, f"处理过程中发生错误: {str(e)}")


class MainWindow(QMainWindow):
    """主窗口 - PDF版本"""
    
    def __init__(self):
        super().__init__()
        self.config_manager = ConfigManager()
        self.generator_thread = None
        
        self.init_ui()
        self.load_templates()
    
    def init_ui(self):
        """初始化UI"""
        self.setWindowTitle("汽车维修服务询价邮件生成器 (PDF版本)")
        self.setMinimumSize(800, 600)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QVBoxLayout()
        
        # 标题
        title_label = QLabel("汽车维修服务询价邮件生成器 (PDF版本)")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; padding: 10px;")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(title_label)
        
        # 基本设置组
        settings_group = QGroupBox("基本设置")
        settings_layout = QVBoxLayout()
        
        # 数据库路径
        db_layout = QHBoxLayout()
        db_layout.addWidget(QLabel("数据库路径:"))
        self.db_path_edit = QLineEdit()
        self.db_path_edit.setText(self.config_manager.get_database_path())
        db_layout.addWidget(self.db_path_edit)
        db_browse_btn = QPushButton("浏览...")
        db_browse_btn.clicked.connect(self.browse_database)
        db_layout.addWidget(db_browse_btn)
        settings_layout.addLayout(db_layout)
        
        # PDF文件目录
        pdf_layout = QHBoxLayout()
        pdf_layout.addWidget(QLabel("PDF文件目录:"))
        self.pdf_path_edit = QLineEdit()
        self.pdf_path_edit.setText(self.config_manager.get_pdf_directory())
        pdf_layout.addWidget(self.pdf_path_edit)
        pdf_browse_btn = QPushButton("浏览...")
        pdf_browse_btn.clicked.connect(self.browse_pdf_directory)
        pdf_layout.addWidget(pdf_browse_btn)
        settings_layout.addLayout(pdf_layout)
        
        # 教程附件路径
        tutorial_layout = QHBoxLayout()
        tutorial_layout.addWidget(QLabel("教程附件:"))
        self.tutorial_path_edit = QLineEdit()
        self.tutorial_path_edit.setText(self.config_manager.get_tutorial_attachment_path())
        tutorial_layout.addWidget(self.tutorial_path_edit)
        tutorial_browse_btn = QPushButton("浏览...")
        tutorial_browse_btn.clicked.connect(self.browse_tutorial)
        tutorial_layout.addWidget(tutorial_browse_btn)
        settings_layout.addLayout(tutorial_layout)
        
        # 保存设置按钮
        save_settings_btn = QPushButton("保存设置")
        save_settings_btn.clicked.connect(self.save_settings)
        settings_layout.addWidget(save_settings_btn)
        
        settings_group.setLayout(settings_layout)
        main_layout.addWidget(settings_group)
        
        # 模板管理组
        template_group = QGroupBox("模板管理")
        template_layout = QVBoxLayout()
        
        self.template_list = QListWidget()
        self.template_list.setMaximumHeight(150)
        template_layout.addWidget(self.template_list)
        
        template_btn_layout = QHBoxLayout()
        add_template_btn = QPushButton("添加模板")
        add_template_btn.clicked.connect(self.add_template)
        template_btn_layout.addWidget(add_template_btn)
        
        edit_template_btn = QPushButton("编辑模板")
        edit_template_btn.clicked.connect(self.edit_template)
        template_btn_layout.addWidget(edit_template_btn)
        
        delete_template_btn = QPushButton("删除模板")
        delete_template_btn.clicked.connect(self.delete_template)
        template_btn_layout.addWidget(delete_template_btn)
        
        template_btn_layout.addStretch()
        template_layout.addLayout(template_btn_layout)
        
        template_group.setLayout(template_layout)
        main_layout.addWidget(template_group)
        
        # 生成邮件组
        generate_group = QGroupBox("生成邮件")
        generate_layout = QVBoxLayout()
        
        task_id_layout = QHBoxLayout()
        task_id_layout.addWidget(QLabel("Task ID:"))
        self.task_id_edit = QLineEdit()
        self.task_id_edit.setPlaceholderText("请输入Task ID (支持多个，用逗号分隔)")
        task_id_layout.addWidget(self.task_id_edit)
        generate_layout.addLayout(task_id_layout)
        
        template_select_layout = QHBoxLayout()
        template_select_layout.addWidget(QLabel("选择模板:"))
        self.template_combo = QComboBox()
        template_select_layout.addWidget(self.template_combo)
        generate_layout.addLayout(template_select_layout)
        
        self.generate_btn = QPushButton("生成询价邮件 (附PDF)")
        self.generate_btn.setStyleSheet("font-size: 14px; padding: 10px; background-color: #4CAF50; color: white;")
        self.generate_btn.clicked.connect(self.generate_email)
        generate_layout.addWidget(self.generate_btn)
        
        generate_group.setLayout(generate_layout)
        main_layout.addWidget(generate_group)
        
        # 日志输出
        log_group = QGroupBox("处理日志")
        log_layout = QVBoxLayout()
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(150)
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        main_layout.addWidget(log_group)
        
        main_layout.addStretch()
        central_widget.setLayout(main_layout)
    
    def load_templates(self):
        """加载模板列表"""
        self.template_list.clear()
        self.template_combo.clear()
        
        templates = self.config_manager.get_templates()
        for template_name in templates.keys():
            self.template_list.addItem(template_name)
            self.template_combo.addItem(template_name)
    
    def browse_database(self):
        """浏览数据库文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择数据库文件", "", "数据库文件 (*.db *.sqlite)"
        )
        if file_path:
            self.db_path_edit.setText(file_path)
    
    def browse_pdf_directory(self):
        """浏览PDF文件目录"""
        folder_path = QFileDialog.getExistingDirectory(self, "选择PDF文件目录")
        if folder_path:
            self.pdf_path_edit.setText(folder_path)
    
    def browse_tutorial(self):
        """浏览教程附件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择教程附件", "", "所有文件 (*.*)"
        )
        if file_path:
            self.tutorial_path_edit.setText(file_path)
    
    def save_settings(self):
        """保存设置"""
        self.config_manager.set_database_path(self.db_path_edit.text())
        self.config_manager.set_pdf_directory(self.pdf_path_edit.text())
        self.config_manager.set_tutorial_attachment_path(self.tutorial_path_edit.text())
        
        QMessageBox.information(self, "成功", "设置已保存")
        self.log("设置已保存")
    
    def add_template(self):
        """添加模板 - PDF版本简化配置"""
        from template_config_dialog_pdf import TemplateConfigDialogPDF
        dialog = TemplateConfigDialogPDF(self)
        if dialog.exec():
            config = dialog.get_config()
            if config:
                template_name = config.pop("template_name")
                self.config_manager.add_template(template_name, config)
                self.load_templates()
                QMessageBox.information(self, "成功", f"模板 '{template_name}' 已添加")
                self.log(f"添加模板: {template_name}")
    
    def edit_template(self):
        """编辑模板"""
        from template_config_dialog_pdf import TemplateConfigDialogPDF
        current_item = self.template_list.currentItem()
        if not current_item:
            QMessageBox.warning(self, "警告", "请先选择要编辑的模板")
            return
        
        template_name = current_item.text()
        template_config = self.config_manager.get_template(template_name)
        
        dialog = TemplateConfigDialogPDF(self, template_name, template_config)
        if dialog.exec():
            config = dialog.get_config()
            if config:
                config.pop("template_name")
                self.config_manager.update_template(template_name, config)
                QMessageBox.information(self, "成功", f"模板 '{template_name}' 已更新")
                self.log(f"更新模板: {template_name}")
    
    def delete_template(self):
        """删除模板"""
        current_item = self.template_list.currentItem()
        if not current_item:
            QMessageBox.warning(self, "警告", "请先选择要删除的模板")
            return
        
        template_name = current_item.text()
        reply = QMessageBox.question(
            self, "确认删除", f"确定要删除模板 '{template_name}' 吗？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            self.config_manager.delete_template(template_name)
            self.load_templates()
            QMessageBox.information(self, "成功", f"模板 '{template_name}' 已删除")
            self.log(f"删除模板: {template_name}")
    
    def generate_email(self):
        """生成邮件"""
        task_id_input = self.task_id_edit.text().strip()
        if not task_id_input:
            QMessageBox.warning(self, "错误", "请输入Task ID")
            return
        
        template_name = self.template_combo.currentText()
        if not template_name:
            QMessageBox.warning(self, "错误", "请选择模板")
            return
        
        task_ids = [tid.strip() for tid in task_id_input.split(',') if tid.strip()]
        if len(task_ids) > 1:
            self.log(f"检测到批量模式: {len(task_ids)} 个Task ID")
        
        self.generate_btn.setEnabled(False)
        self.log_text.clear()
        self.log(f"开始生成询价邮件 - Task ID: {task_id_input}, 模板: {template_name}")
        
        self.generator_thread = EmailGeneratorThread(task_id_input, template_name, self.config_manager)
        self.generator_thread.progress.connect(self.log)
        self.generator_thread.finished.connect(self.on_generate_finished)
        self.generator_thread.start()
    
    def on_generate_finished(self, success: bool, message: str):
        """生成完成回调"""
        self.generate_btn.setEnabled(True)
        
        if success:
            self.log("✓ " + message)
        else:
            self.log("✗ " + message)
            QMessageBox.critical(self, "错误", message)
    
    def log(self, message: str):
        """添加日志"""
        self.log_text.append(message)
