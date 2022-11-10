import pandas as pd

path_sys = 'C:/Log/OneDrive - GALAPAGOS CAPITAL/'

#Classe que cria um objeto que conecta ao banco de dados
class connection_db:
    def __init__(self, db_str): # classe construtora
        self.db_str = db_str
        self.db = self.db_con()

    def db_con(self):
        import pymysql
        sql_username, sql_password = self.credenciais()
        host_name = 'log-rds.cmdzknpslnec.us-east-1.rds.amazonaws.com'
        sql_main_database = self.db_str
        db = pymysql.connect(host=host_name, user=sql_username, passwd=sql_password, db=sql_main_database, port=3306, connect_timeout=100, cursorclass=pymysql.cursors.DictCursor)
        return db

    def db_close(self) -> object:
        """

        :rtype: object
        """
        self.db.close()
        #self.server.stop()

    def db_commit(self):
        self.db.commit()

    def query(self, query):
        cursor = self.db.cursor()
        cursor.execute(query)
        result = cursor.fetchall()
        return result
        #self.db_close()

    def credenciais(self):
        df = pd.read_csv(path_sys + 'credenciais.csv', delimiter=';')
        return df['user'][0], df['password'][0]
