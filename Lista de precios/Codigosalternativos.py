import pgconnection

def Get_query_list(connection, query_string):

    """
    Levanta los datos de una query y los convierte a un diccionario
    Formato - {Codigo_original:Codigo_alternativo}
    """
    cursor = connection.cursor()
    cursor.execute(query_string)
    data = cursor.fetchall()
    cursor.close()

    data_reformat = {}
    for chunk in data:
        codigo = str(chunk[0])
        alternativo = str(chunk[1])
        data_reformat[codigo] = alternativo

    return data_reformat
