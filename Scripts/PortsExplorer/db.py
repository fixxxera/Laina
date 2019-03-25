import sqlite3


class Database(object):
    def __init__(self):
        self.conn = sqlite3.connect('ports.db')
        self.c = self.conn.cursor()
        self.c.execute(
            "CREATE TABLE IF NOT EXISTS portlist (port TEXT NOT NULL UNIQUE, destination_name TEXT NOT NULL, destination_code TEXT NOT NULL)")
        self.conn.commit()

    def close(self):
        self.conn.close()
        self.c.close()

    def insert(self, port_name, destination_name, destination_code):
        self.c.execute('INSERT INTO portlist (port, destination_name, destination_code) VALUES(?, ?, ?)',
                       (port_name, destination_code, destination_name))
        self.conn.commit()

    def get_all(self):
        self.c.execute("SELECT * FROM portlist")
        data = self.c.fetchall()
        return data

    def find_specific_port(self, portname):
        self.c.execute("SELECT * FROM portlist WHERE port LIKE '%" + portname + "%'")
        data = self.c.fetchall()
        return data

    def find_specific_dest(self, destname):
        self.c.execute("SELECT * FROM portlist WHERE destination_name='" + destname + "'")
        data = self.c.fetchall()
        return data

    def find_specific_code(self, destcode):
        self.c.execute("SELECT * FROM portlist WHERE destination_code='" + destcode + "'")
        data = self.c.fetchall()
        return data

    def find_by_port_and_code(self, portname, destcode):
        self.c.execute(
            "SELECT * FROM portlist WHERE destination_code='" + destcode + "' AND port LIKE '%" + portname + "%'")
        data = self.c.fetchall()
        return data

    def find_by_port_and_dest(self, portname, destname):
        self.c.execute(
            "SELECT * FROM portlist WHERE destination_name='" + destname + "' AND port LIKE '%" + portname + "%'")
        data = self.c.fetchall()
        return data

    def find_by_all(self, portname, destname, destcode):
        self.c.execute(
            "SELECT * FROM portlist WHERE destination_name='" + destname + "' AND port LIKE '%" + portname + "%' AND destination_code='" + destcode + "'")
        data = self.c.fetchall()
        return data

    def find_by_dest_and_code(self, destname, destcode):
        self.c.execute(
            "SELECT * FROM portlist WHERE destination_code LIKE '%" + destcode + "%' AND destination_name LIKE '%" + destname + "%'")
        data = self.c.fetchall()
        return data

    def remove(self, portname):
        self.c.execute("DELETE FROM portlist WHERE port='" + portname + "';")
        self.conn.commit()

    def alter_table(self, column_name):
        self.c.execute("ALTER TABLE portlist ADD COLUMN '"+column_name+"' TEXT NOT NULL;")
        self.conn.commit()
