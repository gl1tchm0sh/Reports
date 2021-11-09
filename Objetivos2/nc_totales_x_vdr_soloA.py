import pgconnection
from totales_x_vdr import fetch_data, data_reformat, list_vdrs
from datetime import date


def main_process():
    """
    Ejecuta la query conformada para traer el dato de la facturación
    solamente en A
    """
    vdrs = list_vdrs()
    query_1 = "SELECT SUM(ALIAS_0.SUBTOTALIMPORTE) FROM V_TRCREDITOVENTA ALIAS_0 LEFT OUTER JOIN V_TREXTENSION ALIAS_2 ON ALIAS_0.TREXTENSION_ID = ALIAS_2.ID LEFT OUTER JOIN V_ESQUEMAOPERATIVO ALIAS_8 ON ALIAS_2.ESQUEMAOPERATIVO_ID = ALIAS_8.ID LEFT OUTER JOIN V_FLAG ALIAS_4 ON ALIAS_0.FLAG_ID = ALIAS_4.ID WHERE ALIAS_0.BO_PLACE_ID = '{89C23513-3F01-11D5-86AD-0080AD403F5F}'   AND   0=0  AND  ALIAS_0.TIPOTRANSACCION_ID = '{9B9915C3-4FA6-11D5-B060-004854841C8A}' AND ALIAS_0.ESTADO = 'C' AND ALIAS_8.CODIGO LIKE '01'"
    totales_por_vendedor = {}
    conn = pgconnection.get_connection("DB")

    # Cicla el listado de vendedores, agregando su nombre al query, y ajusta automaticamente la fecha
    # para reconocer el mes en curso, hasta el dia de ejecucion.
    for vdr in vdrs:
        # Formato de fecha AAAAMMDD sin espacios ni caracteres intermedios
        date_from = str(date.today().replace(day=1)).replace("-","")
        date_until = str(date.today()).replace("-","")
        query_2 = f" AND ALIAS_0.NOMBREORIGINANTE = '{vdr}'"
        query_3 = f"AND ALIAS_0.FECHAACTUAL <= '{date_until}000000000' AND ALIAS_0.FECHAACTUAL >= '{date_from}000000000'"

        query = query_1 + query_2 + query_3
        formatted = fetch_data(conn, query)
        if formatted[0][0] == None: #Saltea el vendedor si no tiene ventas facturadas.
            continue
        totales_por_vendedor = totales_por_vendedor | data_reformat(vdr, formatted)
    print('NC_SoloA calculado')
    return totales_por_vendedor

# Siguiente porcion diseñada para probar cambios en el modulo.
# Si se ejecuta este modulo como principal, prueba sobre el vendedor declarado
if __name__ == '__main__':
    print(main_process()['VDR'])
