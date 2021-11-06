import pgconnection
from totales_x_vdr import list_vdrs,fetch_data
from datetime import date


# Da formato de Diccionario con clave {Vendedor : {Rubro : Importe} }
def data_reformat(vdr,rbr,data):
    new_data = {}
    new_data[vdr] = {}
    new_data[vdr][rbr] = data
    return new_data

# Conforma dinamicamente las querys en base al listado de vendedores activos, y rubros hardcodeados
def main_process():
    vdrs = list_vdrs()
    conn = pgconnection.get_connection("BD")
    query_1 = "SELECT  SUM(ALIAS_0.TOTAL2_IMPORTE) FROM V_ITEMFACTURAVENTA ALIAS_0 LEFT OUTER JOIN V_PRODUCTO ALIAS_7 ON ALIAS_7.CODIGO = ALIAS_0.NOMBREREFERENCIA LEFT OUTER JOIN V_RUBRO ALIAS_8 ON ALIAS_8.ID = ALIAS_7.RUBRO_ID WHERE ALIAS_0.TIPOTRANSACCION_ID = '{89C235A5-3F01-11D5-86AD-0080AD403F5F}' AND ( ALIAS_0.UNIDADOPERATIVA_ID = '{89C2350B-3F01-11D5-86AD-0080AD403F5F}'  OR   ALIAS_0.UNIDADOPERATIVA_ID IS NULL  )"
    rbrs = [380, 395, 297, 275, 280, 164, 165, 166, 167]
    results = {}

    for vdr in vdrs:
        # Formato de fecha AAAAMMDD sin espacios ni caracteres intermedios
        date_from = str(date.today().replace(day=1)).replace("-","")
        date_until = str(date.today()).replace("-","")
        # date_from = '20210901'
        # date_until = '20210930'
        query_2 = f" AND ALIAS_0.NOMBREORIGINANTETR = '{vdr}'"
        query_3 = f" AND ALIAS_0.FECHADOCUMENTO <= '{date_until}000000000' AND ALIAS_0.FECHADOCUMENTO >= '{date_from}000000000'"

        for rubro in rbrs:
            query_4 = f" AND ALIAS_8.NOMBRERUBRO = '{rubro}'"
            if rubro == 275 or rubro == 280:
                query_5 = query_1.replace('ALIAS_0.TOTAL2_IMPORTE','ALIAS_0.CANTIDAD2_CANTIDAD')
                query = query_5 + query_2 + query_3 + query_4
            else:
                query = query_1 + query_2 + query_3 + query_4
            formatted = fetch_data(conn, query)
            if formatted[0][0] == None: #Saltea el vendedor si no tiene ventas facturadas.
                continue

            if vdr in results:
                results[vdr][rubro] = data_reformat(vdr,rubro, str(formatted[0][0]))[vdr][rubro]
            else:
                results = results | data_reformat(vdr, rubro, str(formatted[0][0]))

    print("oxr calculado")
    return results

# Siguiente porcion diseñada para probar cambios en el modulo.
# Si se ejecuta este modulo como principal, prueba sobre el vendedor declarado
if __name__ == '__main__':
    print(main_process()['JORGE ESTEBAN AYERDI'])

