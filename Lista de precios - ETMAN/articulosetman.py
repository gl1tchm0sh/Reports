import pgconnection
from Retrieve_rbrs import retrieve

# Toma una conexion abierta de pgconnection, pasa un query
# y devuelve los datos en forma de Tupla
def fetch_data(connection, query_string):

    cursor = connection.cursor('') #nombre de la BD
    cursor.execute(query_string)
    data = cursor.fetchall()
    cursor.close()

    return data

# Convierte las claves de un diccionario en una lista
def keys_to_list(dict):
    key_list = []
    for key in dict.keys():
        key_list.append(key)
    return key_list

# Utiliza las claves del diccionario (Rubros) para añadir
# filtros a la query
def build_query(rubros):
    rbrs_lst = keys_to_list(rubros)
    main_query = ""
    for rbr in rbrs_lst:
        if rbr == None:
            continue
        if main_query == "":
            main_query += f" ALIAS_6.NOMBRERUBRO LIKE '{rbr}'"
        else:
            main_query += f" OR ALIAS_6.NOMBRERUBRO LIKE '{rbr}'"

    return main_query


def main_process(filepath):

    conn = pgconnection.get_connection("")
    rbr_dsc = retrieve(filepath)

    query_1 = "SELECT DISTINCT ON (ALIAS_0.CODIGO) ALIAS_6.NOMBRERUBRO ALIAS_6_NOMBRERUBRO, ALIAS_0.CODIGO ALIAS_0_CODIGO,   ALIAS_0.DESCRIPCION ALIAS_0_DESCRIPCION, (CASE WHEN (SUM(ALIAS_7.CANTIDAD2_CANTIDAD) OVER(PARTITION BY ALIAS_0.CODIGO)) IS NULL THEN 0 ELSE round((SUM(ALIAS_7.CANTIDAD2_CANTIDAD) OVER(PARTITION BY ALIAS_0.CODIGO))::NUMERIC,2) END) AS ALIAS_7_CANTIDAD, ALIAS_0.VALOR2_NOMBRE ALIAS_0_VALOR2_NOMBRE, ALIAS_0.VALOR2_MOMENTO ALIAS_0_VALOR2_MOMENTO, itemposicionadorimpuestos.coeficiente FROM V_PRECIO ALIAS_0 LEFT OUTER JOIN V_UNIDADFINANCIERA ALIAS_2 ON ALIAS_0.VALOR2_UNIDADVALORIZACION_ID = ALIAS_2.ID LEFT OUTER JOIN V_UNIDADMEDIDA ALIAS_3 ON ALIAS_0.DCANTIDAD2_UNIDADMEDIDA_ID = ALIAS_3.ID LEFT OUTER JOIN V_UNIDADMEDIDA ALIAS_4 ON ALIAS_0.HCANTIDAD2_UNIDADMEDIDA_ID = ALIAS_4.ID LEFT OUTER JOIN v_PRODUCTO ALIAS_5 ON ALIAS_5.ID = REFERENCIA_ID LEFT OUTER JOIN V_RUBRO ALIAS_6 ON ALIAS_5.RUBRO_ID = ALIAS_6.ID LEFT JOIN posicionadorimpuestos ON posicionadorimpuestos.ID = alias_5.POSICIONADORIMPUESTOS_ID LEFT JOIN itemposicionadorimpuestos ON itemposicionadorimpuestos.BO_PLACE_ID = posicionadorimpuestos.itemsposicionadorimpuestos_id LEFT JOIN definicionimpuesto ON definicionimpuesto.ID = itemposicionadorimpuestos.DEFINICIONIMPUESTO_id LEFT JOIN IMPUESTO ON IMPUESTO.ID = definicionimpuesto.IMPUESTO_ID LEFT JOIN V_ITEMINVENTARIO ALIAS_7 ON ALIAS_7.PRODUCTO_ID = ALIAS_5.ID WHERE IMPUESTO.NOMBRE = 'IVA TASA' AND ALIAS_0.BO_PLACE_ID = '{15430301-BD4D-4501-8393-A3E02F149542}' AND ALIAS_0.ACTIVESTATUS <> 2 AND ("
    query_2 = build_query(rbr_dsc)
    query_3 = " ) AND NOT ALIAS_0.CODIGO LIKE 'MO.%' ORDER BY ALIAS_0.CODIGO, ALIAS_6.NOMBRERUBRO"

    query = query_1 + query_2 + query_3
    data = fetch_data(conn, query)

    return data



