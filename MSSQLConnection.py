import pyodbc

import mysql.connector

class MSSQLConnection:
    def __init__(self, host='localhost', port=3306, database='qlnv1', user='root', password='minhtu150320'):
        self.host = host
        self.port = port
        self.database = database
        self.user = user
        self.password = password
        self.connection = None

    def connect(self):
        try:
            self.connection = mysql.connector.connect(
                host=self.host,
                port=self.port,
                database=self.database,
                user=self.user,
                password=self.password
            )
            if self.connection.is_connected():
                print("Connection successful")
        except mysql.connector.Error as e:
            print("Error in connection:", e)

    def query(self, sql):
        try:
            cursor = self.connection.cursor()
            cursor.execute(sql)
            return cursor.fetchall()
        except mysql.connector.Error as e:
            print("Error in query:", e)
            return None

    def update(self, sql):
        try:
            with self.connection.cursor() as cursor:
                cursor.execute(sql)
                self.connection.commit()
        except mysql.connector.Error as e:
            print("Error in update:", e)

    def insert(self, sql):
        try:
            with self.connection.cursor() as cursor:
                cursor.execute(sql)
                self.connection.commit()
        except mysql.connector.Error as e:
            print("Error in insert:", e)

    def delete(self, sql):
        try:
            with self.connection.cursor() as cursor:
                cursor.execute(sql)
                self.connection.commit()
        except mysql.connector.Error as e:
            print("Error in delete:", e)

    def close(self):
        if self.connection and self.connection.is_connected():
            self.connection.close()
            print("Connection closed")

# Sử dụng lớp để kết nối MySQL
db = MSSQLConnection(host='localhost', port=3306, database='qlnv', user='root', password='minhtu150320')
db.connect()
