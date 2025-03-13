import pythoncom
import win32com.server.util
import sqlite3
import win32api
import logging

# Khởi tạo logging
logging.basicConfig(filename="dcom_server.log", level=logging.INFO,
                    format="%(asctime)s - %(levelname)s - %(message)s")
# tên database
sql_db="customer_service.db"

class DCOMServer:
    _public_methods_ = ['Hello','authenticate','FetchUsers','AddUsers']
    _reg_progid_ = "DCOM.Server"
    _reg_clsid_ = "{F9DABC68-3C3A-4F4D-8B25-65853ABEE832}"

    def authenticate(self):
        user = win32api.GetUserName()
        logging.info(f"Request received from user: {user}")
        if user.lower() not in ["user"]:
            logging.warning(f"Unauthorized access attempt by {user}")
            # raise Exception("Access Denied: Unauthorized user.")
        return user
    def __init__(self):
        """ Khởi tạo kết nối database """
        self.conn = sqlite3.connect(sql_db, check_same_thread=False)
        self.cur = self.conn.cursor()
    def FetchUsers(self):
        logging.info("FetchUsers method called")
        cur=self.cur
        try:
            sql_command="SELECT id, name FROM Users"
            cur.execute(sql_command)
            users = cur.fetchall()
            print(str(users))
            return str(users)
        except Exception as e:
            logging.error(f"Database error: {str(e)}")
            return f"Error: {str(e)}"
    def AddUsers(self):
        return 0

    def Hello(self):
        """ Return a greeting with logging """
        # user = win32api.GetUserName()
        user = self.authenticate()
        message = f"Hello, {user}! Welcome to the my DCOM Server."
        # logging.info(f"Hello method called by {user}")
        # message = f"Hello, abc! Welcome to the my DCOM Server. "
        return message



if __name__ == "__main__":
    from win32com.server.register import UseCommandLine
    UseCommandLine(DCOMServer)
