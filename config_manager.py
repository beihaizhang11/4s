"""
配置管理模块
负责管理模板配置、数据库路径、邮件设置等
"""
import json
import os
from typing import Dict, List, Any


class ConfigManager:
    """配置管理器类"""
    
    def __init__(self, config_file: str = "config.json"):
        self.config_file = config_file
        self.config = self._load_config()
    
    def _load_config(self) -> Dict[str, Any]:
        """加载配置文件"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                print(f"加载配置文件失败: {e}")
                return self._get_default_config()
        else:
            return self._get_default_config()
    
    def _get_default_config(self) -> Dict[str, Any]:
        """获取默认配置"""
        return {
            "database_path": "",  # 共享文件夹数据库路径
            "local_database_path": "./local_db.db",  # 本地数据库路径
            "output_folder": "./output",  # 输出文件夹（标准版本用）
            "pdf_directory": "",  # PDF文件目录路径（PDF版本用）
            "tutorial_attachment_path": "",  # 教程附件路径
            "templates": {}  # 模板配置
        }
    
    def save_config(self):
        """保存配置到文件"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=4)
            return True
        except Exception as e:
            print(f"保存配置文件失败: {e}")
            return False
    
    def get_database_path(self) -> str:
        """获取数据库路径"""
        return self.config.get("database_path", "")
    
    def set_database_path(self, path: str):
        """设置数据库路径"""
        self.config["database_path"] = path
        self.save_config()
    
    def get_local_database_path(self) -> str:
        """获取本地数据库路径"""
        return self.config.get("local_database_path", "./local_db.db")
    
    def get_output_folder(self) -> str:
        """获取输出文件夹"""
        return self.config.get("output_folder", "./output")
    
    def set_output_folder(self, path: str):
        """设置输出文件夹"""
        self.config["output_folder"] = path
        self.save_config()
    
    def get_pdf_directory(self) -> str:
        """获取PDF文件目录"""
        return self.config.get("pdf_directory", "")
    
    def set_pdf_directory(self, path: str):
        """设置PDF文件目录"""
        self.config["pdf_directory"] = path
        self.save_config()
    
    def get_tutorial_attachment_path(self) -> str:
        """获取教程附件路径"""
        return self.config.get("tutorial_attachment_path", "")
    
    def set_tutorial_attachment_path(self, path: str):
        """设置教程附件路径"""
        self.config["tutorial_attachment_path"] = path
        self.save_config()
    
    def get_templates(self) -> Dict[str, Dict]:
        """获取所有模板配置"""
        return self.config.get("templates", {})
    
    def get_template(self, template_name: str) -> Dict:
        """获取指定模板配置"""
        return self.config.get("templates", {}).get(template_name, None)
    
    def add_template(self, template_name: str, template_config: Dict):
        """添加新模板"""
        if "templates" not in self.config:
            self.config["templates"] = {}
        self.config["templates"][template_name] = template_config
        self.save_config()
    
    def update_template(self, template_name: str, template_config: Dict):
        """更新模板配置"""
        if "templates" not in self.config:
            self.config["templates"] = {}
        self.config["templates"][template_name] = template_config
        self.save_config()
    
    def delete_template(self, template_name: str):
        """删除模板"""
        if "templates" in self.config and template_name in self.config["templates"]:
            del self.config["templates"][template_name]
            self.save_config()
    
    def get_template_excel_path(self, template_name: str) -> str:
        """获取模板Excel路径"""
        template = self.get_template(template_name)
        if template:
            return template.get("excel_template_path", "")
        return ""
    
    def get_template_field_mappings(self, template_name: str) -> Dict[str, List[str]]:
        """获取模板字段映射
        返回格式: {"field_name": ["A1", "B2"], ...}
        """
        template = self.get_template(template_name)
        if template:
            return template.get("field_mappings", {})
        return {}
    
    def get_template_email_to_first(self, template_name: str) -> List[str]:
        """获取邮件TO列表的首位固定收件人列表"""
        template = self.get_template(template_name)
        if template:
            email_to_first = template.get("email_to_first", [])
            # 兼容旧版本配置（如果是字符串，转换为列表）
            if isinstance(email_to_first, str):
                return [email_to_first] if email_to_first else []
            return email_to_first
        return []
    
    def get_template_email_to_fixed(self, template_name: str) -> List[str]:
        """获取邮件TO列表的固定收件人（在动态收件人之后）"""
        template = self.get_template(template_name)
        if template:
            return template.get("email_to_fixed", [])
        return []
    
    def get_template_email_cc(self, template_name: str) -> List[str]:
        """获取邮件CC列表"""
        template = self.get_template(template_name)
        if template:
            return template.get("email_cc", [])
        return []
    
    def get_template_email_body_template(self, template_name: str) -> str:
        """获取邮件正文模板"""
        template = self.get_template(template_name)
        if template:
            return template.get("email_body_template", "")
        return ""
    
    @staticmethod
    def create_template_config(
        excel_template_path: str,
        field_mappings: Dict[str, List[str]],
        email_to_first: List[str],
        email_to_fixed: List[str],
        email_cc: List[str],
        email_body_template: str
    ) -> Dict:
        """创建模板配置字典"""
        return {
            "excel_template_path": excel_template_path,
            "field_mappings": field_mappings,
            "email_to_first": email_to_first,
            "email_to_fixed": email_to_fixed,
            "email_cc": email_cc,
            "email_body_template": email_body_template
        }
