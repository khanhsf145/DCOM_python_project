import sqlite3
import logging
import win32api

# Khởi tạo logging
logging.basicConfig(filename="mock_dcom_server.log", level=logging.INFO,
                    format="%(asctime)s - %(levelname)s - %(message)s")
# tên database
sql_db = "customer_service.db"

class MockDCOMServer:
    """
    Phiên bản giả lập của DCOMServer để có thể sử dụng web client
    mà không cần đăng ký COM server.
    """
    
    def __init__(self):
        """ Khởi tạo kết nối database """
        self.conn = None
        self.cur = None
        try:
            self.conn = sqlite3.connect(sql_db, check_same_thread=False)
            self.cur = self.conn.cursor()
            self._initialize_database()
            logging.info("Kết nối database thành công")
        except Exception as e:
            logging.error(f"Lỗi kết nối database: {str(e)}")
            
    def _initialize_database(self):
        """Khởi tạo cơ sở dữ liệu với dữ liệu mẫu nếu chưa tồn tại"""
        try:
            if self.cur is None:
                logging.error("Cursor không khả dụng")
                return
                
            # Tạo bảng Users nếu chưa tồn tại
            self.cur.execute('''
                CREATE TABLE IF NOT EXISTS Users (
                    id INTEGER PRIMARY KEY,
                    name TEXT NOT NULL
                )
            ''')
            
            # Tạo bảng Requests nếu chưa tồn tại
            self.cur.execute('''
                CREATE TABLE IF NOT EXISTS Requests (
                    id INTEGER PRIMARY KEY,
                    type TEXT NOT NULL,
                    detail TEXT,
                    user_id INTEGER,
                    FOREIGN KEY (user_id) REFERENCES Users(id)
                )
            ''')
            
            # Kiểm tra xem đã có dữ liệu chưa
            self.cur.execute("SELECT COUNT(*) FROM Users")
            user_count = self.cur.fetchone()[0]
            
            # Thêm dữ liệu mẫu nếu chưa có
            if user_count == 0:
                self.cur.executemany("INSERT INTO Users (id, name) VALUES (?, ?)", [
                    (1, "Nguyễn Văn A"),
                    (2, "Trần Thị B"),
                    (3, "Lê Văn C")
                ])
                
                self.cur.executemany("INSERT INTO Requests (id, type, detail, user_id) VALUES (?, ?, ?, ?)", [
                    (1, "Hỗ trợ kỹ thuật", "Máy tính không khởi động được", 1),
                    (2, "Yêu cầu tính năng", "Thêm tính năng xuất báo cáo PDF", 2),
                    (3, "Báo lỗi", "Ứng dụng bị crash khi mở file lớn", 3)
                ])
                
                if self.conn:
                    self.conn.commit()
                logging.info("Đã thêm dữ liệu mẫu vào database")
        except Exception as e:
            logging.error(f"Lỗi khi khởi tạo database: {str(e)}")

    def hello(self):
        """ chào người dùng và lưu lại sự kiện gọi đến phương thức chào vào file log """
        user = win32api.GetUserName()
        message = f"Hello, {user}! Welcome to my DCOM Server (Mock Version)."
        logging.info(f"Hello method called by {user}")
        return message
        
    def fetch_all_users(self, _id):
        logging.info("FetchUsers method called")
        if self.cur is None:
            return "Error: Database connection not available"
            
        try:
            sql_command = "SELECT id, name FROM Users"
            self.cur.execute(sql_command)
            users = self.cur.fetchall()
            return str(users)
        except Exception as e:
            logging.error(f"Database error: {str(e)}")
            return f"Error: {str(e)}"
            
    def fetch_user(self, _id):
        logging.info(f"fetch_user method called with id={_id}")
        if self.cur is None:
            return "Error: Database connection not available"
            
        try:
            sql_command = f"SELECT id, name FROM Users where id={_id}"
            self.cur.execute(sql_command)
            user = self.cur.fetchone()
            return str(user)
        except Exception as e:
            logging.error(f"Database error: {str(e)}")
            return f"Error: {str(e)}"
            
    def fetch_request(self, _id):
        logging.info("fetch_requests method called")
        if self.cur is None:
            return "Error: Database connection not available"
            
        try:
            sql_command = f"SELECT id, type, detail FROM Requests WHERE id={_id}" if _id != 0 else "SELECT id, type, detail FROM Requests LIMIT 1"
            self.cur.execute(sql_command)
            request = self.cur.fetchone()
            return str(request)
        except Exception as e:
            logging.error(f"Database error: {str(e)}")
            return f"Error: {str(e)}"
            
    def add_users(self, name):
        logging.info(f"add_users method called with name={name}")
        if self.cur is None:
            return "Error: Database connection not available"
            
        try:
            sql_command = "INSERT INTO Users (name) VALUES (?)"
            self.cur.execute(sql_command, (name,))
            if self.conn:
                self.conn.commit()
            return self.cur.lastrowid
        except Exception as e:
            logging.error(f"Database error: {str(e)}")
            return f"Error: {str(e)}"

# Tạo instance của server để sử dụng cho web client
server_instance = MockDCOMServer()

if __name__ == "__main__":
    print("Mock DCOM Server đang chạy...")
    print("Đây là phiên bản giả lập để thử nghiệm web client mà không cần đăng ký COM.")
    
    # Thử gọi một số phương thức để kiểm tra
    server = server_instance
    print("\nKết quả gọi hello():", server.hello())
    print("\nKết quả gọi fetch_all_users():", server.fetch_all_users(0))
    print("\nKết quả gọi fetch_user(1):", server.fetch_user(1))
    print("\nKết quả gọi fetch_request(1):", server.fetch_request(1)) 