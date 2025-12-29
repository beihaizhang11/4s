"""
邮件管理模块
负责通过Outlook创建邮件草稿
"""
import os
from typing import List
import win32com.client


class EmailManager:
    """邮件管理器类"""
    
    def __init__(self):
        self.outlook = None
    
    def initialize_outlook(self) -> bool:
        """初始化Outlook应用"""
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            return True
        except Exception as e:
            print(f"初始化Outlook失败: {e}")
            return False
    
    def create_draft_email(
        self,
        subject: str,
        to_recipients: List[str],
        cc_recipients: List[str],
        body: str,
        attachments: List[str]
    ) -> bool:
        """创建邮件草稿
        
        Args:
            subject: 邮件主题
            to_recipients: 收件人列表
            cc_recipients: 抄送人列表
            body: 邮件正文
            attachments: 附件路径列表
        
        Returns:
            是否成功
        """
        try:
            if not self.outlook:
                if not self.initialize_outlook():
                    return False
            
            # 创建邮件对象
            mail = self.outlook.CreateItem(0)  # 0 = olMailItem
            
            # 设置邮件主题
            mail.Subject = subject
            
            # 设置收件人
            if to_recipients:
                mail.To = "; ".join(to_recipients)
            
            # 设置抄送人
            if cc_recipients:
                mail.CC = "; ".join(cc_recipients)
            
            # 设置邮件正文
            mail.Body = body
            
            # 添加附件
            for attachment_path in attachments:
                if os.path.exists(attachment_path):
                    mail.Attachments.Add(os.path.abspath(attachment_path))
                else:
                    print(f"警告: 附件不存在: {attachment_path}")
            
            # 显示邮件窗口（作为草稿）
            mail.Display()
            
            print("邮件草稿已创建")
            return True
            
        except Exception as e:
            print(f"创建邮件草稿失败: {e}")
            return False
    
    def send_email(
        self,
        subject: str,
        to_recipients: List[str],
        cc_recipients: List[str],
        body: str,
        attachments: List[str]
    ) -> bool:
        """直接发送邮件（不显示）
        
        Args:
            subject: 邮件主题
            to_recipients: 收件人列表
            cc_recipients: 抄送人列表
            body: 邮件正文
            attachments: 附件路径列表
        
        Returns:
            是否成功
        """
        try:
            if not self.outlook:
                if not self.initialize_outlook():
                    return False
            
            # 创建邮件对象
            mail = self.outlook.CreateItem(0)  # 0 = olMailItem
            
            # 设置邮件主题
            mail.Subject = subject
            
            # 设置收件人
            if to_recipients:
                mail.To = "; ".join(to_recipients)
            
            # 设置抄送人
            if cc_recipients:
                mail.CC = "; ".join(cc_recipients)
            
            # 设置邮件正文
            mail.Body = body
            
            # 添加附件
            for attachment_path in attachments:
                if os.path.exists(attachment_path):
                    mail.Attachments.Add(os.path.abspath(attachment_path))
                else:
                    print(f"警告: 附件不存在: {attachment_path}")
            
            # 发送邮件
            mail.Send()
            
            print("邮件已发送")
            return True
            
        except Exception as e:
            print(f"发送邮件失败: {e}")
            return False
    
    @staticmethod
    def generate_email_body(task_description: str, bg_description: str = None) -> str:
        """生成邮件正文
        
        Args:
            task_description: 任务描述
            bg_description: 背景描述（可选）
        
        Returns:
            邮件正文
        """
        body = f"""hello
请参考附件服务 {task_description}
"""
        
        # 如果bg_description不为空，添加背景信息
        if bg_description and bg_description.strip():
            body += f"""
背景：{bg_description}
"""
        
        body += """
请2个工作日内确认是否可以提供服务，如果可以，请补充报价信息回复询价单。
预估到货时间，并附上报价依据截图。"""
        
        return body
    
    @staticmethod
    def generate_email_body_batch(task_ids: list, task_description: str, bg_description: str = None) -> str:
        """生成批量处理的邮件正文
        
        Args:
            task_ids: Task ID列表
            task_description: 任务描述（使用第一个task的）
            bg_description: 背景描述（可选）
        
        Returns:
            邮件正文
        """
        task_ids_text = ', '.join(task_ids)
        
        body = f"""hello
请参考附件服务 {task_description}

本次询价包含多个Task ID: {task_ids_text}
"""
        
        # 如果bg_description不为空，添加背景信息
        if bg_description and bg_description.strip():
            body += f"""
背景：{bg_description}
"""
        
        body += """
请2个工作日内确认是否可以提供服务，如果可以，请补充报价信息回复询价单。
预估到货时间，并附上报价依据截图。"""
        
        return body
