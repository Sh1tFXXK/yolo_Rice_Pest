# -*- coding: utf-8 -*-
"""
该脚本用于清空 'users.db' 数据库中 'users' 表的所有数据。

当需要重置所有用户注册信息，但又想保留数据库表结构时，可以运行此脚本。
它会执行一个 DELETE FROM users 的SQL命令，将表内容清空。
"""
import sqlite3
import os

def get_db_path():
    """
    获取数据库文件的绝对路径。

    Returns:
        str: 'users.db' 文件的完整路径。
    """
    # __file__ 是当前脚本的路径
    # os.path.dirname(__file__) 获取脚本所在的目录
    # os.path.join(...) 将目录和文件名拼接成一个完整的路径
    return os.path.join(os.path.dirname(__file__), 'users.db')

def clear_users_table():
    """
    清空 users 表中的所有数据。
    """
    db_path = get_db_path()

    if not os.path.exists(db_path):
        print(f"错误：数据库文件 '{db_path}' 不存在。")
        return

    try:
        # 1. 连接到 SQLite 数据库
        conn = sqlite3.connect(db_path)
        
        # 2. 创建一个 cursor 对象
        cursor = conn.cursor()
        
        # 3. 执行 DELETE 语句，删除表中的所有行
        # 使用 DELETE FROM users 而不是 DROP TABLE users，因为我们只想清空数据，而不是删除整个表
        cursor.execute("DELETE FROM users")
        
        # 4. 提交事务，应用更改
        conn.commit()
        
        print(f"成功：已清空 '{db_path}' 数据库的 'users' 表中的所有数据。")
        
    except sqlite3.Error as e:
        print(f"数据库操作失败: {e}")
        
    finally:
        # 5. 确保无论成功还是失败，都关闭数据库连接
        if 'conn' in locals() and conn:
            conn.close()

if __name__ == "__main__":
    # 当该脚本被直接执行时，调用清空函数
    clear_users_table()