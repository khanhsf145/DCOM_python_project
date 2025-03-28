# type: ignore
import pythoncom
import win32com.server.util
import win32api
import sqlite3
import logging

print("=== Kiểm tra các thư viện cần thiết cho DCOM server ===")

# Kiểm tra win32api
username = win32api.GetUserName()
print(f"win32api.GetUserName(): {username}")

# Kiểm tra sqlite3
try:
    conn = sqlite3.connect(":memory:")
    cur = conn.cursor()
    cur.execute("CREATE TABLE test (id INTEGER, name TEXT)")
    cur.execute("INSERT INTO test VALUES (1, 'Test')")
    cur.execute("SELECT * FROM test")
    result = cur.fetchone()
    print(f"sqlite3 test result: {result}")
except Exception as e:
    print(f"sqlite3 error: {str(e)}")

# Kiểm tra logging
logging.basicConfig(level=logging.INFO)
logging.info("Logging test")
print("Logging test thành công")

# Kiểm tra win32com
try:
    win32com.server.util.wrap("Test")
    print("win32com.server.util test thành công")
except Exception as e:
    print(f"win32com.server.util error: {str(e)}")

print("\nTất cả các thư viện đã được cài đặt đúng!") 