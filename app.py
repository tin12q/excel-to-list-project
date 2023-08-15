import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QPushButton
import requests
from check import read_excel_to_2d_list, save_2d_list_to_excel, check

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Set window title and size
        self.setWindowTitle('User Checker')
        self.setGeometry(100, 100, 800, 600)

        # Create a table widget
        self.table = QTableWidget(self)
        self.table.setGeometry(50, 50, 700, 400)

        # Create a button to check users
        self.check_button = QPushButton('Check Users', self)
        self.check_button.setGeometry(50, 500, 150, 50)
        self.check_button.clicked.connect(self.check_users)

        # Create a button to save results
        self.save_button = QPushButton('Save Results', self)
        self.save_button.setGeometry(250, 500, 150, 50)
        self.save_button.clicked.connect(self.save_results)

    def check_users(self):
        # Call the check function from check.py
        check()

        # Get the list of users from check.py
        self.list_usr = read_excel_to_2d_list('./usrdata.xlsx')

        # Create a table with the user data
        self.table.setRowCount(len(self.list_usr))
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(['id', 'name', 'gt', 'tag'])

        for i, row in enumerate(self.list_usr):
            for j, val in enumerate(row):
                item = QTableWidgetItem(str(val))
                self.table.setItem(i, j, item)

    def save_results(self):
        # Call the check function from check.py
        check()

        # Get the list of users from check.py
        list_usr = read_excel_to_2d_list('./usrdata.xlsx')

        # Save the results to a new Excel file
        res_usr = []
        for element in list_usr:
            url = f'https://openapi.zalo.me/v2.0/oa/getprofile?data={{"user_id":\'{element[0]}\'}}'
            payload = {}
            headers = {
              'access_token': 'fjM_1t3ENI-jrxu7UCbS8glGlm9XdoPui9l6F6Ry907gaRHvIlDN2-YOq2fWWduEnkIUJM2wVplouU99O9r47zRdzqL_ZNStrlJcOr-S21xzvP9uQwzpKEZXnpPBd5jHwlAx3qUAM6tRw-egJFfoIFo7a3vjopHVtu3G5rZbTclEWUO8Rfi6QUtvlHHddpf0oOEEFcdTQMh4sV00OhWNKkBPzt5mnLCJmTxzTH2zUpAYtzra0fDG1OBgv7CgaNW1cEoUF3kQ14deyQS_U9eqI_BWWZ9Cd3PrpDsq06t6EchcsgehTv0p3jJwYcjXcoimvlIQ66YDEL34y8mjv2dN47RINYC'
            }
            response = requests.request("GET", url, headers=headers, data=payload)
            if not response.json().get('data'):
                res_usr.append([element[0], element[1], element[2], 'Không có'])
                element[3] = 'Không có'
            else:
                res_usr.append([element[0], element[1], element[2], element[3]])

        save_2d_list_to_excel(res_usr, './res2.xlsx')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())  # type: ignore