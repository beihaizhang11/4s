"""
PDF扫描模块
负责扫描指定目录下的所有PDF文件
"""
import os
from typing import List


class PDFScanner:
    """PDF扫描器类"""
    
    def __init__(self):
        pass
    
    @staticmethod
    def scan_pdf_files(directory: str) -> List[str]:
        """扫描指定目录下的所有PDF文件
        
        Args:
            directory: 要扫描的目录路径
        
        Returns:
            PDF文件的完整路径列表
        """
        if not directory or not os.path.exists(directory):
            print(f"目录不存在: {directory}")
            return []
        
        pdf_files = []
        
        try:
            # 遍历目录下的所有文件
            for filename in os.listdir(directory):
                # 检查是否是PDF文件（不区分大小写）
                if filename.lower().endswith('.pdf'):
                    full_path = os.path.join(directory, filename)
                    # 确保是文件而不是目录
                    if os.path.isfile(full_path):
                        pdf_files.append(full_path)
            
            # 按文件名排序
            pdf_files.sort()
            
            print(f"在 {directory} 中找到 {len(pdf_files)} 个PDF文件")
            for pdf_file in pdf_files:
                print(f"  - {os.path.basename(pdf_file)}")
            
            return pdf_files
            
        except Exception as e:
            print(f"扫描PDF文件失败: {e}")
            return []
    
    @staticmethod
    def validate_directory(directory: str) -> bool:
        """验证目录是否存在且可访问
        
        Args:
            directory: 目录路径
        
        Returns:
            是否有效
        """
        if not directory:
            return False
        
        if not os.path.exists(directory):
            return False
        
        if not os.path.isdir(directory):
            return False
        
        # 检查是否有读取权限
        if not os.access(directory, os.R_OK):
            return False
        
        return True
