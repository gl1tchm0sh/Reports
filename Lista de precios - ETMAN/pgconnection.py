import psycopg2

# Ingrese host sin comillas, usuario y password entre las comillas definidas.

def get_connection(dbname):

    connect_str = f"dbname={dbname} host= user='' password=''"

    return psycopg2.connect(connect_str)
