"""
数据库管理模块
负责从共享文件夹复制数据库、查询数据等操作
"""
import sqlite3
import shutil
import os
from typing import Dict, List, Optional


class DatabaseManager:
    """数据库管理器类"""
    
    def __init__(self, source_db_path: str, local_db_path: str):
        self.source_db_path = source_db_path
        self.local_db_path = local_db_path
        self.conn = None
    
    def copy_database_to_local(self) -> bool:
        """从共享文件夹复制数据库到本地"""
        try:
            if not os.path.exists(self.source_db_path):
                raise FileNotFoundError(f"源数据库文件不存在: {self.source_db_path}")
            
            # 创建本地数据库目录
            local_dir = os.path.dirname(self.local_db_path)
            if local_dir and not os.path.exists(local_dir):
                os.makedirs(local_dir)
            
            # 复制数据库文件
            shutil.copy2(self.source_db_path, self.local_db_path)
            print(f"数据库已复制到: {self.local_db_path}")
            return True
        except Exception as e:
            print(f"复制数据库失败: {e}")
            return False
    
    def connect(self) -> bool:
        """连接到本地数据库"""
        try:
            if not os.path.exists(self.local_db_path):
                print(f"本地数据库不存在: {self.local_db_path}")
                return False
            
            self.conn = sqlite3.connect(self.local_db_path)
            self.conn.row_factory = sqlite3.Row  # 使查询结果可以通过列名访问
            return True
        except Exception as e:
            print(f"连接数据库失败: {e}")
            return False
    
    def disconnect(self):
        """断开数据库连接"""
        if self.conn:
            self.conn.close()
            self.conn = None
    
    def get_task_data(self, task_id: str) -> Optional[Dict]:
        """根据task_id获取任务数据"""
        if not self.conn:
            print("数据库未连接")
            return None
        
        try:
            cursor = self.conn.cursor()
            query = """
                SELECT 
                    task_id, sender_name, sender_mail, sender_phone, 
                    VIN, task_dept, sent_time, task_description, bg_description,
                    car_model, carid
                FROM tasks
                WHERE task_id = ?
            """
            cursor.execute(query, (task_id,))
            row = cursor.fetchone()
            
            if row:
                return dict(row)
            else:
                print(f"未找到task_id: {task_id}")
                return None
        except Exception as e:
            print(f"查询任务数据失败: {e}")
            return None
    
    def get_staff_emails_by_task_id(self, task_id: str) -> List[str]:
        """根据task_id从tasks_staff表获取staff_email列表"""
        if not self.conn:
            print("数据库未连接")
            return []
        
        try:
            cursor = self.conn.cursor()
            query = """
                SELECT staff_email
                FROM tasks_staff
                WHERE task_id = ?
            """
            cursor.execute(query, (task_id,))
            rows = cursor.fetchall()
            
            return [row['staff_email'] for row in rows if row['staff_email']]
        except Exception as e:
            print(f"查询员工邮箱失败: {e}")
            return []
    
    def generate_serial_number(self, task_data: Dict) -> str:
        """生成流水号
        格式: WST_{car_model}_{carid}_{sender_name}_{sent_time}
        """
        try:
            car_model = task_data.get('car_model', '').replace(' ', '_')
            carid = task_data.get('carid', '').replace(' ', '_')
            sender_name = task_data.get('sender_name', '').replace(' ', '_')
            sent_time = task_data.get('sent_time', '').replace(' ', '_').replace(':', '-')
            
            serial_number = f"WST_{car_model}_{carid}_{sender_name}_{sent_time}"
            return serial_number
        except Exception as e:
            print(f"生成流水号失败: {e}")
            return "WST_UNKNOWN"
    
    def prepare_data_for_excel(self, task_id: str) -> Optional[Dict]:
        """准备要写入Excel的数据，包括流水号"""
        task_data = self.get_task_data(task_id)
        if not task_data:
            return None
        
        # 生成流水号
        serial_number = self.generate_serial_number(task_data)
        task_data['serial_number'] = serial_number
        
        return task_data
    
    def __enter__(self):
        """支持with语句"""
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """支持with语句"""
        self.disconnect()
