"""
gfiquery.py

Class to excecute query on GFI database, and prepare result.

This software uses an oracle external library:
    cx_Oracle: http://cx-oracle.sourceforge.net/html/

"""


import cx_Oracle



class GFIquery:
    """
    Manage GFI Oracle queries.

    Input: sql to be executed, and Oracle credentials in the form
    user/pass@db

    Output: class data and header arrays set if query successful.
    """

    sql = None
    credentials = None
    status = False
    headers = None
    data = {}

    def __init__(self,credentials,sql):
        """
        Create new query object given query and connection string,
        execute, and store result in object variables if successful.
        """

        self.credentials = credentials
        self.sql = sql

        # print "cred: %s, sql: %s" % (self.credentials,self.sql)


    def execute(self):
        try:
            connection = cx_Oracle.connect(self.credentials)
        except cx_Oracle.DatabaseError:
            connection.close()
            status = False
            return

        try:
            cursor = connection.cursor()
            cursor.execute(self.sql)
        except cx_Oracle.DatabaseError:
            connection.close()
            status = False
            return

        # get names in position 0 of description array
        self.headers = [i[0].lower() for i in cursor.description]
        for field in self.headers: self.data[field] = []

        for r in cursor:
            for field,value in zip(self.headers,r):
                self.data[field].append(value)
        connection.close()

        self.status = True
        
