from totales_x_vdr import main_process as totales
from objetivos_x_rubro import main_process as objetivos
from Desarrollo import main_process as desarrollo
from totales_x_vdr_soloA import main_process as solo_a


def merge():
    """Funcion utilizada para unificar todos los datos de las queries en un solo diccionario"""
    txv = totales()
    oxr = objetivos()
    des = desarrollo()
    txva = solo_a()
    merge = {} # Los resultados se almacenan en esta variable

    # Listas para definir los rubros que entran en sus respectivos objetivos
    bosch_lst = [164, 165, 166, 167]
    tyrolit_list = [275, 280]

    for vdr in txv:
        merge[vdr] = {'Total':txv[vdr]}
        merge[vdr]['TotalA'] = txva[vdr]
        if vdr in oxr:
            merge[vdr] = merge[vdr] | oxr[vdr]
        if vdr in des:
            merge[vdr] = merge[vdr] | des[vdr]

        # Recorre los valores según los rubros definidos y los va sumando en la variable
        # definida (acumulado_bosch / acumulado_tyrolit), luego borra el contenido
        acumulado_bosch = 0
        for rubro in bosch_lst:
            if rubro in merge[vdr]:
                acumulado_bosch += float(merge[vdr][rubro])
                del(merge[vdr][rubro])
        merge[vdr]['Bosch'] = acumulado_bosch

        acumulado_tyrolit = 0
        for rubro in tyrolit_list:
            if rubro in merge[vdr]:
                acumulado_tyrolit += float(merge[vdr][rubro])
                del (merge[vdr][rubro])
        merge[vdr]['Tyrolit'] = acumulado_tyrolit

        # Reemplaza los nombres almacenados en el diccionario por su expresión en texto
        # para aumentar su legibilidad.
        replace = {380:'Gedore', 297:'Atlas', 395:'Bovenau'}
        keys = list(replace.keys())

        for rubro in keys:
            if rubro in merge[vdr]:
                merge[vdr][replace[rubro]] = merge[vdr][rubro]
                del(merge[vdr][rubro])



    print('merged')
    return merge


# Siguiente porcion diseñada para probar cambios en el modulo.
# Si se ejecuta este modulo como principal, prueba sobre el vendedor declarado
if __name__ == '__main__':
    print(merge()['JORGE ESTEBAN AYERDI'])