import pgconnection
from totales_x_vdr import list_vdrs,fetch_data

def data_reformat(vdr,rbr,data):
    new_data = {}
    new_data[vdr] = {}
    new_data[vdr][rbr] = data
    return new_data

# Conforma por separado la query 4 en base al listado de rubros hardcodeado.
def form_query(rbrs):
    query = ""
    for rubro in rbrs:
        if rubro == rbrs[0]:
            query = f" AND (ALIAS_8.NOMBRERUBRO = '{rubro}'"
        else:
            query += f" OR ALIAS_8.NOMBRERUBRO = '{rubro}'"
    else:
        query += ")"
    return query


# Conforma dinamicamente las querys en base al listado de vendedores activos, y rubros hardcodeados
def main_process():
    vdrs = list_vdrs()
    conn = pgconnection.get_connection("BD")
    query_1 = "SELECT  SUM(ALIAS_0.TOTAL2_IMPORTE) FROM V_ITEMFACTURAVENTA ALIAS_0 LEFT OUTER JOIN V_PRODUCTO ALIAS_7 ON ALIAS_7.CODIGO = ALIAS_0.NOMBREREFERENCIA LEFT OUTER JOIN V_RUBRO ALIAS_8 ON ALIAS_8.ID = ALIAS_7.RUBRO_ID WHERE ALIAS_0.TIPOTRANSACCION_ID = '{89C235A5-3F01-11D5-86AD-0080AD403F5F}' AND ( ALIAS_0.UNIDADOPERATIVA_ID = '{89C2350B-3F01-11D5-86AD-0080AD403F5F}'  OR   ALIAS_0.UNIDADOPERATIVA_ID IS NULL  )"
    rbrs = [99,120,123,124,125,127,140,150,151,152,153,161,162,164,165,166,167,168,170,171,175,181,183,185,189,192,198,240,245,247,249,250,252,254,256,260,261,265,266,277,278,283,284,285,286,287,288,289,293,296,297,298,303,304,306,309,311,313,314,324,330,331,340,355,356,365,368,370,379,380,381,382,390,395,397,601,602,603,604,605,606,607,610,611,612,613,620,625,626,650,666]
    results = {}

    for vdr in vdrs:
        # Formato de fecha AAAAMMDD sin espacios ni caracteres intermedios
        date_from = str(date.today().replace(day=1)).replace("-","")
        date_until = str(date.today()).replace("-","")
        # date_from = '20210901'
        # date_until = '20210930'
        query_2 = f" AND ALIAS_0.NOMBREORIGINANTETR = '{vdr}'"
        query_3 = f" AND ALIAS_0.FECHADOCUMENTO <= '{date_until}000000000' AND ALIAS_0.FECHADOCUMENTO >= '{date_from}000000000'"
        query_4 = form_query(rbrs)

        query = query_1 + query_2 + query_3 + query_4
        formatted = fetch_data(conn, query)
        if formatted[0][0] == None: #Saltea el vendedor si no tiene ventas facturadas.
            continue

        if vdr in results:
            results[vdr]['Desarrollo'] = data_reformat(vdr,rubro, str(formatted[0][0]))[vdr][rubro]
        else:
            results = results | data_reformat(vdr, 'Desarrollo', str(formatted[0][0]))



    print("Desarrollo calculado")
    return results

# Siguiente porcion diseñada para probar cambios en el modulo.
# Si se ejecuta este modulo como principal, prueba sobre el vendedor declarado
if __name__ == '__main__':
    print(main_process()['JORGE ESTEBAN AYERDI'])