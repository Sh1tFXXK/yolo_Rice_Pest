# -*- coding: utf-8 -*-
"""
该脚本是数据库辅助模块。

它封装了所有与SQLite数据库 (users.db) 相关的操作，包括：
- 数据库和表的初始化 (init_db)
- 检查用户是否存在 (user_exists)
- 添加新用户 (add_user)
- 验证用户登录 (check_user)
- 更新用户信息，如密码和头像 (update_user)

通过将数据库逻辑集中在此文件中，可以使主程序代码更清晰，并方便数据库的管理和维护。
"""
import sqlite3
import os

def get_db_path():
    """
    获取数据库文件的绝对路径。

    Returns:
        str: 'users.db' 文件的完整路径。
    """
    return os.path.join(os.path.dirname(__file__), 'users.db')

def init_db():
    """
    初始化数据库。
    
    - 检查 'users' 表是否存在，如果不存在则创建它。
    - 检查 'avatar' 列是否存在，如果不存在则向表中添加该列，以实现向后兼容。
    """
    db_path = get_db_path()
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    # 检查avatar列是否存在，如果不存在则添加
    c.execute("PRAGMA table_info(users)")
    columns = [info[1] for info in c.fetchall()]
    if 'avatar' not in columns:
        try:
            c.execute('ALTER TABLE users ADD COLUMN avatar TEXT')
        except sqlite3.OperationalError:
            # 如果users表不存在，则创建它
            c.execute('''CREATE TABLE users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password TEXT NOT NULL,
                avatar TEXT
            )''')
    else:
        # 如果表已存在，确保表结构是最新的
        c.execute('''CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            avatar TEXT
        )''')
    conn.commit()
    conn.close()

def user_exists(username):
    """
    检查指定用户名的用户是否已存在于数据库中。

    Args:
        username (str): 要检查的用户名。

    Returns:
        bool: 如果用户存在则返回 True，否则返回 False。
    """
    db_path = get_db_path()
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    c.execute('SELECT id FROM users WHERE username=?', (username,))
    result = c.fetchone()
    conn.close()
    return result is not None

def add_user(username, password, avatar='avator/default.jpg'):
    """
    向数据库中添加一个新用户。

    Args:
        username (str): 新用户的用户名。
        password (str): 新用户的密码。
        avatar (str, optional): 用户头像的路径。默认为 'avator/default.jpg'。

    Returns:
        bool: 如果添加成功则返回 True，如果用户已存在则返回 False。
    """
    db_path = get_db_path()
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    try:
        c.execute('INSERT INTO users (username, password, avatar) VALUES (?, ?, ?)', (username, password, avatar))
        conn.commit()
        return True
    except sqlite3.IntegrityError:
        return False
    finally:
        conn.close()

def check_user(username, password):
    """
    验证用户的用户名和密码是否正确。

    Args:
        username (str): 要验证的用户名。
        password (str): 要验证的密码。

    Returns:
        dict or None: 如果验证成功，返回包含 'username' 和 'avatar' 的字典。
                      如果验证失败，返回 None。
    """
    db_path = get_db_path()
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    c.execute('SELECT username, avatar FROM users WHERE username=? AND password=?', (username, password))
    result = c.fetchone()
    conn.close()
    if result:
        return {'username': result[0], 'avatar': result[1]}
    return None

def update_user(username, new_password=None, new_avatar=None):
    """
    根据用户名更新用户信息（密码和/或头像）。
    
    可以只更新密码、只更新头像，或同时更新两者。

    Args:
        username (str): 要更新的用户的用户名。
        new_password (str, optional): 新的密码。如果为 None，则不更新密码。
        new_avatar (str, optional): 新头像的路径。如果为 None，则不更新头像。

    Returns:
        bool: 如果更新成功则返回 True，否则返回 False。
    """
    db_path = get_db_path()
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    
    updates = []
    params = []
    
    if new_password:
        updates.append("password = ?")
        params.append(new_password)
        
    if new_avatar:
        updates.append("avatar = ?")
        params.append(new_avatar)
        
    if not updates:
        return False # 如果没有任何更新，直接返回
        
    params.append(username)
    
    try:
        query = f"UPDATE users SET {', '.join(updates)} WHERE username = ?"
        c.execute(query, tuple(params))
        conn.commit()
        return True
    except sqlite3.Error as e:
        print(f"数据库更新失败: {e}")
        return False
    finally:
        conn.close()

# 启动时自动初始化数据库
init_db()
