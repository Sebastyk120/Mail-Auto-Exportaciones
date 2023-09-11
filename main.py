import pathlib
import os

import win32com.client as win32

# ---------------------------------------------------------------------AGENCIAS DE ADUANA
Alisas = "ALISAS"
Alisas_Co = "reservas@alisas.com.co; jefeaeropuerto@alisas.com.co; aduanacierres@alisas.com.co; " \
            "aeropuerto@alisas.com.co; aduana@alisas.com.co; "

Mag_Custom = "MAG CUSTOM"
Mag_Custom_Co = "luis.cardona@cargologisticsystem.com; leanis.narvaez@cargologisticsystem.com; " \
                "lorena.tovar@cargologisticsystem.com; tatiana.mendoza@cargologisticsystem.com; " \
                "david.gantiva@cargologisticsystem.com; juliana.aguilar@cargologisticsystem.com; " \
                "jeanet.parada@cargologisticsystem.com; maria.vargas@cargologisticsystem.com; " \
                "adriana.avila@cargologisticsystem.com "
# ---------------------------------------------------------------------AGENCIAS DE CARGA
Carguex_Sa = "CARGUEX S.A"
Carguex_Sa_Co = "wilson.eslava@cargex.com.co; aeropuerto.documentos@cargex.com.co; operaciones1@cargex.com.co; " \
                "aeropuerto.cgx@cargex.com.co; john.martinez@cargex.com.co; Giovani.ramirez@cargex.com.co "

Cargo_logis = "CARGO LOGISTIC SYSTEM S.A.S"
Cargo_logis_Co = ";luis.cardona@cargologisticsystem.com; leanis.narvaez@cargologisticsystem.com; " \
                 "tatiana.mendoza@cargologisticsystem.com; "

Kuehne = "KUEHNE + NAGEL S.A.S"
Kuehne_Co = "knbog.invoicing.perishables@kuehne-nagel.com; knbog.documents.perishables@kuehne-nagel.com; " \
            "juan.Linares@Kuehne-Nagel.com; "

Logiztik_a = "LOGIZTIK ALLIANCE GROUP S.A.S"
Logiztik_a_CO = "col-facturas@logiztikalliance.com; bog-ops@logiztikalliance.com;"
# ----------------------------------------------------------------------------------------------------------------------
SOLICITUD_MUISCA_OA = 1
SOLICITUD_CO = 2
FITOS = 3
ENVIO_DOCUMENTOS = 4
AGREGAR_CLIENTES = 5
VER_CLIENTES = 6
SALIR = 7


def mostrar_menu():
    os.system('cls')
    print(f'''                  DOCUMENTOS EXPORTACIÓN HEAVENS.
    {SOLICITUD_MUISCA_OA}) SOLICITUD MUISCA Y OA.
    {SOLICITUD_CO}) SOLICITUD CERTIFICADO DE ORIGEN.
    {FITOS}) SOLICITUD DE FITOS.
    {ENVIO_DOCUMENTOS}) ENVIO DE DOCUMENTOS CLIENTE.
    {AGREGAR_CLIENTES}) AGREGAR UN NUEVO CLIENTE.
    {VER_CLIENTES}) VER LISTADO DE CLIENTES ACTUAL.
    {SALIR}) SALIR...''')


# -------------------------------------------------CARGAR CLIENTES EXPORTACIÓN------------------------------------------
def cargar_clientes(clientes, nombre_archivo):
    if pathlib.Path(nombre_archivo).exists():
        with open(nombre_archivo, 'r') as archivo:
            for linea in archivo:
                co_cliente, nombre, emails, documentos = linea.strip().split(',')
                clientes.setdefault(co_cliente, (nombre, emails, documentos))
    else:
        with open(nombre_archivo, 'w') as archivo:
            pass


# -------------------------------------------------AGREGAR CLIENTES EXPORTACIÓN-----------------------------------------
def agregar_clientes(clientes, nombre_archivo):
    os.system('cls')
    print("                 AGREGAR UN CLIENTE NUEVO.")
    co_cliente = int(input("Escriba el código del cliente: "))
    if clientes.get(co_cliente):
        print("¡El cliente ya existe!")
    else:
        nombre = input("Nombre del cliente: ")
        emails = input("Emails: ")
        documentos = input("Documentos obligatorios: ")
        clientes.setdefault(co_cliente, (nombre, emails, documentos))
        with open(nombre_archivo, 'a') as archivo:
            archivo.write(f'{co_cliente},{nombre},{emails},{documentos}\n')
        print("El cliente se ha agregado correctamente.")


# -------------------------------------------------VER CLIENTES EXPORTACIÓN---------------------------------------------

def ver_clientes(clientes):
    os.system('cls')
    print("                 VER CLIENTES EXPORTACIÓN")
    if len(clientes) > 0:
        cantidad_clientes = 0
        for co_cliente, datos in clientes.items():
            print(f'Codigo: {co_cliente}')
            print(f'Nombre: {datos[0]}')
            print(f'Emails:{datos[1]}')
            print(f'Documentos: {datos[2]}')
            print("---------------------------------------------------------------------------------------------------")
            cantidad_clientes = cantidad_clientes + 1
        if cantidad_clientes == 0:
            print("No existe ningún cliente.")

        else:
            print(f'Se encontraron: {cantidad_clientes} clientes.')


# -----------------------------------------------------------------------------------------------------------------------
def solicitud_muisca(clientes):
    eleccion = 0
    os.system('cls')
    print(f"                    SOLICITUD MUISCA Y OA.")
    seleccion = 0
    if len(clientes) > 0:
        codigo = input("Escriba el codigo del cliente: ")
        for co_cliente, datos in clientes.items():
            if codigo in co_cliente:
                print(f'Usted a seleccionado a: {datos[0]}')
                while seleccion != 3:
                    print("Agencias de aduana: ")
                    print("1. ", Alisas)
                    print("2. ", Mag_Custom)
                    print("3. Volver...")
                    seleccion = int(input("Escriba una de las opciones:"))
                    if seleccion == 1:
                        print("Ha seleccionado: ", Alisas)
                        while eleccion != 5:
                            eleccion = 0
                            print("Elija una de las agencias de carga: ")
                            print("1.", Carguex_Sa)
                            print("2. ", Cargo_logis)
                            print("3. ", Kuehne)
                            print("4. ", Logiztik_a)
                            print("5. Volver...")
                            eleccion = int(input("Escriba una de las opciones:"))
                            if eleccion == 1:
                                print(f'Solicitud Muisca y OA con: {Alisas} y {Carguex_Sa} ')
                                print(f'Los documentos a adjuntar son: FACTURA - CARTAS DE RESPONSABILIDAD')
                                olApp = win32.Dispatch('Outlook.Application')
                                olNS = olApp.GetnameSpace('MAPI')
                                mailItem = olApp.CreateItem(0)
                                mailItem.Subject = 'SOLICITUD MUISCA Y OA ' + datos[0]
                                mailItem.BodyFormat = 2
                                mailItem.Body = 'Cordial saludo\nAdjunto cartas de responsabilidad y factura para generación de Muisca y OA \nAgradezco su colaboración.\nCordialmente.'
                                mailItem.To = '' + Alisas_Co + Carguex_Sa_Co
                                mailItem.Cc = 'gerencia@heavensfruit.com; mabdime@heavensfruit.com; valentinagaray@heavensfruit.com; mainassistant@heavensfruit.com'
                                mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('subgerencia@heavensfruit.com'),))
                                mailItem.Display()
                            elif eleccion == 2:
                                print(f'Solicitud Muisca y OA con: {Alisas} y {Cargo_logis} ')
                                print(f'Los documentos a adjuntar son: FACTURA - CARTAS DE RESPONSABILIDAD.')
                                olApp = win32.Dispatch('Outlook.Application')
                                olNS = olApp.GetnameSpace('MAPI')
                                mailItem = olApp.CreateItem(0)
                                mailItem.Subject = 'SOLICITUD MUISCA Y OA ' + datos[0]
                                mailItem.BodyFormat = 2
                                mailItem.Body = 'Cordial saludo,\nAdjunto cartas de responsabilidad y factura para generación de Muisca y OA \nAgradezco su colaboración.\nCordialmente.'
                                mailItem.To = '' + Alisas_Co + Cargo_logis_Co
                                mailItem.Cc = 'gerencia@heavensfruit.com; mabdime@heavensfruit.com; valentinagaray@heavensfruit.com; mainassistant@heavensfruit.com'
                                mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('subgerencia@heavensfruit.com')))
                                mailItem.Display()
                            elif eleccion == 3:
                                print(f'Solicitud Muisca y OA con: {Alisas} y {Kuehne} ')
                                print(f'Los documentos a adjuntar son: FACTURA - CARTAS DE RESPONSABILIDAD.')
                                olApp = win32.Dispatch('Outlook.Application')
                                olNS = olApp.GetnameSpace('MAPI')
                                mailItem = olApp.CreateItem(0)
                                mailItem.Subject = 'SOLICITUD MUISCA Y OA ' + datos[0]
                                mailItem.BodyFormat = 2
                                mailItem.Body = 'Cordial saludo,\nAdjunto cartas de responsabilidad y factura para generación de Muisca y OA \nAgradezco su colaboración.\nCordialmente.'
                                mailItem.To = '' + Alisas_Co + Kuehne_Co
                                mailItem.Cc = 'gerencia@heavensfruit.com; mabdime@heavensfruit.com; valentinagaray@heavensfruit.com; mainassistant@heavensfruit.com'
                                mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('subgerencia@heavensfruit.com')))
                                mailItem.Display()
                            elif eleccion == 4:
                                print(f'Solicitud Muisca y OA con: {Alisas} y {Logiztik_a} ')
                                print(f'Los documentos a adjuntar son: FACTURA - CARTAS DE RESPONSABILIDAD.')
                                olApp = win32.Dispatch('Outlook.Application')
                                olNS = olApp.GetnameSpace('MAPI')
                                mailItem = olApp.CreateItem(0)
                                mailItem.Subject = 'SOLICITUD MUISCA Y OA ' + datos[0]
                                mailItem.BodyFormat = 2
                                mailItem.Body = 'Cordial saludo,\nAdjunto cartas de responsabilidad y factura para generación de Muisca y OA \nAgradezco su colaboración.\nCordialmente.'
                                mailItem.To = '' + Alisas_Co + Logiztik_a_CO
                                mailItem.Cc = 'gerencia@heavensfruit.com; mabdime@heavensfruit.com; valentinagaray@heavensfruit.com; mainassistant@heavensfruit.com'
                                mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('subgerencia@heavensfruit.com')))
                                mailItem.Display()
                            else:
                                print("Opción no valida.")

                    elif seleccion == 2:
                        print("Ha seleccionado: ", Mag_Custom)
                        while eleccion != 5:
                            eleccion = 0
                            print("Elija una de las agencias de carga: ")
                            print("1.", Carguex_Sa)
                            print("2. ", Cargo_logis)
                            print("3. ", Kuehne)
                            print("4. ", Logiztik_a)
                            print("5. Volver...")
                            eleccion = int(input("Escriba una de las opciones:"))
                            if eleccion == 1:
                                print(f'Solicitud Muisca y OA con: {Mag_Custom} y {Carguex_Sa} ')
                                print(f'Los documentos a adjuntar son: FACTURA - CARTAS DE RESPONSABILIDAD')
                                olApp = win32.Dispatch('Outlook.Application')
                                olNS = olApp.GetnameSpace('MAPI')
                                mailItem = olApp.CreateItem(0)
                                mailItem.Subject = 'SOLICITUD MUISCA Y OA ' + datos[0]
                                mailItem.BodyFormat = 2
                                mailItem.Body = 'Cordial saludo\nAdjunto cartas de responsabilidad y factura para generación de Muisca y OA \nAgradezco su colaboración.\nCordialmente.'
                                mailItem.To = '' + Mag_Custom_Co + Carguex_Sa_Co
                                mailItem.Cc = 'gerencia@heavensfruit.com; mabdime@heavensfruit.com; valentinagaray@heavensfruit.com; mainassistant@heavensfruit.com'
                                mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('subgerencia@heavensfruit.com')))
                                mailItem.Display()
                            elif eleccion == 2:
                                print(f'Solicitud Muisca y OA con: {Mag_Custom} y ;{Cargo_logis} ')
                                print(f'Los documentos a adjuntar son: FACTURA - CARTAS DE RESPONSABILIDAD.')
                                olApp = win32.Dispatch('Outlook.Application')
                                olNS = olApp.GetnameSpace('MAPI')
                                mailItem = olApp.CreateItem(0)
                                mailItem.Subject = 'SOLICITUD MUISCA Y OA ' + datos[0]
                                mailItem.BodyFormat = 2
                                mailItem.Body = 'Cordial saludo,\nAdjunto cartas de responsabilidad y factura para generación de Muisca y OA \nAgradezco su colaboración.\nCordialmente.'
                                mailItem.To = '' + Mag_Custom_Co + Cargo_logis_Co
                                mailItem.Cc = 'gerencia@heavensfruit.com; mabdime@heavensfruit.com; valentinagaray@heavensfruit.com; mainassistant@heavensfruit.com'
                                mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('subgerencia@heavensfruit.com')))
                                mailItem.Display()
                            elif eleccion == 3:
                                print(f'Solicitud Muisca y OA con: {Mag_Custom} y {Kuehne} ')
                                print(f'Los documentos a adjuntar son: FACTURA - CARTAS DE RESPONSABILIDAD.')
                                olApp = win32.Dispatch('Outlook.Application')
                                olNS = olApp.GetnameSpace('MAPI')
                                mailItem = olApp.CreateItem(0)
                                mailItem.Subject = 'SOLICITUD MUISCA Y OA ' + datos[0]
                                mailItem.BodyFormat = 2
                                mailItem.Body = 'Cordial saludo,\nAdjunto cartas de responsabilidad y factura para generación de Muisca y OA \nAgradezco su colaboración.\nCordialmente.'
                                mailItem.To = '' + Mag_Custom_Co + Kuehne_Co
                                mailItem.Cc = 'gerencia@heavensfruit.com; mabdime@heavensfruit.com; valentinagaray@heavensfruit.com; mainassistant@heavensfruit.com'
                                mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('subgerencia@heavensfruit.com')))
                                mailItem.Display()
                            elif eleccion == 4:
                                print(f'Solicitud Muisca y OA con: {Mag_Custom} y {Logiztik_a} ')
                                print(f'Los documentos a adjuntar son: FACTURA - CARTAS DE RESPONSABILIDAD.')
                                olApp = win32.Dispatch('Outlook.Application')
                                olNS = olApp.GetnameSpace('MAPI')
                                mailItem = olApp.CreateItem(0)
                                mailItem.Subject = 'SOLICITUD MUISCA Y OA ' + datos[0]
                                mailItem.BodyFormat = 2
                                mailItem.Body = 'Cordial saludo,\nAdjunto cartas de responsabilidad y factura para generación de Muisca y OA \nAgradezco su colaboración.\nCordialmente.'
                                mailItem.To = '' + Mag_Custom_Co + Logiztik_a_CO
                                mailItem.Cc = 'gerencia@heavensfruit.com; mabdime@heavensfruit.com; valentinagaray@heavensfruit.com; mainassistant@heavensfruit.com'
                                mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('subgerencia@heavensfruit.com')))
                                mailItem.Display()
                            else:
                                print("Opción no valida.")
    else:
        print("No hay contactos registrados")


# -----------------------------------------------------------------------------------------------------------------------
def solicitud_co(clientes):
    seleccion = 0
    print(f"                    SOLICITUD CERTIFICADO DE ORIGEN.")
    while seleccion !=4:
        if seleccion == 1:
            clienteuno(clientes)
            break
        elif seleccion == 2:
            clientedos(clientes)
            break
        elif seleccion == 3:
            clientetres(clientes)
            break
        seleccion = int(input("Escriba el número de clientes que desea ingresar, maximo 3: "))
    else:
        print("Opción no valida.")



def clienteuno(clientes):
    eleccion = 0
    os.system('cls')
    print(f"                    SOLICITUD CERTIFICADO DE ORIGEN Para 1 cliente...")
    seleccion = 0
    if len(clientes) > 0:
        codigo = input("Escriba el codigo del cliente: ")
        for co_cliente, datos in clientes.items():
            if codigo in co_cliente:
                print(f'Usted a seleccionado a: {datos[0]}')
                while seleccion != 5:
                    print(f'''Seleccione la agencia de carga:
                    1) {Carguex_Sa}
                    2) {Cargo_logis}
                    3) {Kuehne}
                    4) {Logiztik_a}
                    5) Volver ...''')
                    seleccion = int(input("Seleccione una opción: "))
                    if seleccion == 1:
                        print(f'SOLICITUD CERTIFICADO DE ORIGEN: {Carguex_Sa}')
                        print(f'Los documentos a adjuntar son: FACTURA QR.')
                        olApp = win32.Dispatch('Outlook.Application')
                        olNS = olApp.GetnameSpace('MAPI')
                        mailItem = olApp.CreateItem(0)
                        mailItem.Subject = 'SOLICITUD CO ' + datos[0]
                        mailItem.BodyFormat = 2
                        mailItem.Body = 'Cordial saludo,\nAdjunto Facturas QR para generación de Certificados de Origen, \nAgradezco su colaboración.\nCordialmente.'
                        mailItem.To = '' + Carguex_Sa_Co
                        mailItem.Cc = 'gerencia@heavensfruit.com; mabdime@heavensfruit.com; valentinagaray@heavensfruit.com; mainassistant@heavensfruit.com'
                        mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('subgerencia@heavensfruit.com')))
                        mailItem.Display()
                    elif seleccion == 2:
                        print(f'SOLICITUD CERTIFICADO DE ORIGEN: {Cargo_logis}')
                        print(f'Los documentos a adjuntar son: FACTURA QR.')
                        olApp = win32.Dispatch('Outlook.Application')
                        olNS = olApp.GetnameSpace('MAPI')
                        mailItem = olApp.CreateItem(0)
                        mailItem.Subject = 'SOLICITUD CO ' + datos[0]
                        mailItem.BodyFormat = 2
                        mailItem.Body = 'Cordial saludo,\nAdjunto Facturas QR para generación de Certificados de Origen, \nAgradezco su colaboración.\nCordialmente.'
                        mailItem.To = '' + Cargo_logis_Co
                        mailItem.Cc = 'gerencia@heavensfruit.com; mabdime@heavensfruit.com; valentinagaray@heavensfruit.com; mainassistant@heavensfruit.com'
                        mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('subgerencia@heavensfruit.com')))
                        mailItem.Display()
                    elif seleccion == 3:
                        print(f'SOLICITUD CERTIFICADO DE ORIGEN: {Kuehne}')
                        print(f'Los documentos a adjuntar son: FACTURA QR.')
                        olApp = win32.Dispatch('Outlook.Application')
                        olNS = olApp.GetnameSpace('MAPI')
                        mailItem = olApp.CreateItem(0)
                        mailItem.Subject = 'SOLICITUD CO ' + datos[0]
                        mailItem.BodyFormat = 2
                        mailItem.Body = 'Cordial saludo,\nAdjunto Facturas QR para generación de Certificados de Origen, \nAgradezco su colaboración.\nCordialmente.'
                        mailItem.To = '' + Kuehne_Co
                        mailItem.Cc = 'gerencia@heavensfruit.com; mabdime@heavensfruit.com; valentinagaray@heavensfruit.com; mainassistant@heavensfruit.com'
                        mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('subgerencia@heavensfruit.com')))
                        mailItem.Display()
                    elif seleccion == 4:
                        print(f'SOLICITUD CERTIFICADO DE ORIGEN: {Logiztik_a}')
                        print(f'Los documentos a adjuntar son: FACTURA QR.')
                        olApp = win32.Dispatch('Outlook.Application')
                        olNS = olApp.GetnameSpace('MAPI')
                        mailItem = olApp.CreateItem(0)
                        mailItem.Subject = 'SOLICITUD CO ' + datos[0]
                        mailItem.BodyFormat = 2
                        mailItem.Body = 'Cordial saludo,\nAdjunto Facturas QR para generación de Certificados de Origen, \nAgradezco su colaboración.\nCordialmente.'
                        mailItem.To = '' + Logiztik_a_CO
                        mailItem.Cc = 'gerencia@heavensfruit.com; mabdime@heavensfruit.com; valentinagaray@heavensfruit.com; mainassistant@heavensfruit.com'
                        mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('subgerencia@heavensfruit.com')))
                        mailItem.Display()
                    else:
                        print("Opción no valida")
            else:
                print(f"No se encontro ningun cliente con el codigo: {codigo}")
                break
    else:
        print("No existen datos.")

def clientedos(clientes):
    global clientenombre
    os.system('cls')
    seleccion = 0
    codigo1= input("Escriba el codigo del cliente 1: ")
    for co_cliente, datos in clientes.items():
        if codigo1 in co_cliente:
            clientenombre = datos[0]
            print(f'Usted a seleccionado a: {datos[0]}')
            codigo = input("Escriba el codigo del cliente 2: ")
            for co_cliente, datos in clientes.items():
                if codigo in co_cliente:
                    print(f'Usted a seleccionado a: {datos[0]}')
                    while seleccion != 5:
                        print(f'''Seleccione la agencia de carga:
                            1) {Carguex_Sa}
                            2) {Cargo_logis}
                            3) {Kuehne}
                            4) {Logiztik_a}
                            5) Volver ...''')
                        seleccion = int(input("Seleccione una opción: "))
                        if seleccion == 1:
                            print(f'SOLICITUD CERTIFICADO DE ORIGEN: {Carguex_Sa}')
                            print(f'Los documentos a adjuntar son: FACTURA QR.')
                            olApp = win32.Dispatch('Outlook.Application')
                            olNS = olApp.GetnameSpace('MAPI')
                            mailItem = olApp.CreateItem(0)
                            mailItem.Subject = (f'SOLICITUD CO {datos[0]} - {clientenombre}')
                            mailItem.BodyFormat = 2
                            mailItem.Body = 'Cordial saludo,\nAdjunto Facturas QR para generación de Certificados de Origen, \nAgradezco su colaboración.\nCordialmente.'
                            mailItem.To = '' + Carguex_Sa_Co
                            mailItem.Cc = 'gerencia@heavensfruit.com; mabdime@heavensfruit.com; valentinagaray@heavensfruit.com; mainassistant@heavensfruit.com'
                            mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('subgerencia@heavensfruit.com')))
                            mailItem.Display()
                        elif seleccion == 2:
                            print(f'SOLICITUD CERTIFICADO DE ORIGEN: {Cargo_logis}')
                            print(f'Los documentos a adjuntar son: FACTURA QR.')
                            olApp = win32.Dispatch('Outlook.Application')
                            olNS = olApp.GetnameSpace('MAPI')
                            mailItem = olApp.CreateItem(0)
                            mailItem.Subject = (f'SOLICITUD CO {datos[0]} - {clientenombre}')
                            mailItem.BodyFormat = 2
                            mailItem.Body = 'Cordial saludo,\nAdjunto Facturas QR para generación de Certificados de Origen, \nAgradezco su colaboración.\nCordialmente.'
                            mailItem.To = '' + Cargo_logis_Co
                            mailItem.Cc = 'gerencia@heavensfruit.com; mabdime@heavensfruit.com; valentinagaray@heavensfruit.com; mainassistant@heavensfruit.com'
                            mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('subgerencia@heavensfruit.com')))
                            mailItem.Display()
                        elif seleccion == 3:
                            print(f'SOLICITUD CERTIFICADO DE ORIGEN: {Kuehne}')
                            print(f'Los documentos a adjuntar son: FACTURA QR.')
                            olApp = win32.Dispatch('Outlook.Application')
                            olNS = olApp.GetnameSpace('MAPI')
                            mailItem = olApp.CreateItem(0)
                            mailItem.Subject = (f'SOLICITUD CO {datos[0]} - {clientenombre}')
                            mailItem.BodyFormat = 2
                            mailItem.Body = 'Cordial saludo,\nAdjunto Facturas QR para generación de Certificados de Origen, \nAgradezco su colaboración.\nCordialmente.'
                            mailItem.To = '' + Kuehne_Co
                            mailItem.Cc = 'gerencia@heavensfruit.com; mabdime@heavensfruit.com; valentinagaray@heavensfruit.com; mainassistant@heavensfruit.com'
                            mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('subgerencia@heavensfruit.com')))
                            mailItem.Display()
                        elif seleccion == 4:
                            print(f'SOLICITUD CERTIFICADO DE ORIGEN: {Logiztik_a}')
                            print(f'Los documentos a adjuntar son: FACTURA QR.')
                            olApp = win32.Dispatch('Outlook.Application')
                            olNS = olApp.GetnameSpace('MAPI')
                            mailItem = olApp.CreateItem(0)
                            mailItem.Subject = (f'SOLICITUD CO {datos[0]} - {clientenombre}')
                            mailItem.BodyFormat = 2
                            mailItem.Body = 'Cordial saludo,\nAdjunto Facturas QR para generación de Certificados de Origen, \nAgradezco su colaboración.\nCordialmente.'
                            mailItem.To = '' + Logiztik_a_CO
                            mailItem.Cc = 'gerencia@heavensfruit.com; mabdime@heavensfruit.com; valentinagaray@heavensfruit.com; mainassistant@heavensfruit.com'
                            mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('subgerencia@heavensfruit.com')))
                            mailItem.Display()
                        else:
                            print("Opción no valida")
            else:
                print(f"No se encontro ningun cliente con el codigo: {codigo}")
                break
    else:
        print("No existen datos.")

def clientetres(clientes):
    global clientenombre, clientenombre2
    os.system('cls')
    seleccion = 0
    codigo2 = input("Escriba el codigo del cliente 1: ")
    for co_cliente, datos in clientes.items():
        if codigo2 in co_cliente:
            clientenombre2 = datos[0]
            print(f'Usted a seleccionado a: {datos[0]}')
    codigo1 = input("Escriba el codigo del cliente 1: ")
    for co_cliente, datos in clientes.items():
        if codigo1 in co_cliente:
            clientenombre = datos[0]
            print(f'Usted a seleccionado a: {datos[0]}')
            codigo = input("Escriba el codigo del cliente 2: ")
            for co_cliente, datos in clientes.items():
                if codigo in co_cliente:
                    print(f'Usted a seleccionado a: {datos[0]}')
                    while seleccion != 5:
                        print(f'''Seleccione la agencia de carga:
                                1) {Carguex_Sa}
                                2) {Cargo_logis}
                                3) {Kuehne}
                                4) {Logiztik_a}
                                5) Volver ...''')
                        seleccion = int(input("Seleccione una opción: "))
                        if seleccion == 1:
                            print(f'SOLICITUD CERTIFICADO DE ORIGEN: {Carguex_Sa}')
                            print(f'Los documentos a adjuntar son: FACTURA QR.')
                            olApp = win32.Dispatch('Outlook.Application')
                            olNS = olApp.GetnameSpace('MAPI')
                            mailItem = olApp.CreateItem(0)
                            mailItem.Subject = (f'SOLICITUD CO {datos[0]} - {clientenombre} - {clientenombre2}')
                            mailItem.BodyFormat = 2
                            mailItem.Body = 'Cordial saludo,\nAdjunto Facturas QR para generación de Certificados de Origen, \nAgradezco su colaboración.\nCordialmente.'
                            mailItem.To = '' + Carguex_Sa_Co
                            mailItem.Cc = 'gerencia@heavensfruit.com; mabdime@heavensfruit.com; valentinagaray@heavensfruit.com; mainassistant@heavensfruit.com'
                            mailItem._oleobj_.Invoke(
                                *(64209, 0, 8, 0, olNS.Accounts.Item('subgerencia@heavensfruit.com')))
                            mailItem.Display()
                        elif seleccion == 2:
                            print(f'SOLICITUD CERTIFICADO DE ORIGEN: {Cargo_logis}')
                            print(f'Los documentos a adjuntar son: FACTURA QR.')
                            olApp = win32.Dispatch('Outlook.Application')
                            olNS = olApp.GetnameSpace('MAPI')
                            mailItem = olApp.CreateItem(0)
                            mailItem.Subject = (f'SOLICITUD CO {datos[0]} - {clientenombre} - {clientenombre2}')
                            mailItem.BodyFormat = 2
                            mailItem.Body = 'Cordial saludo,\nAdjunto Facturas QR para generación de Certificados de Origen, \nAgradezco su colaboración.\nCordialmente.'
                            mailItem.To = '' + Cargo_logis_Co
                            mailItem.Cc = 'gerencia@heavensfruit.com; mabdime@heavensfruit.com; valentinagaray@heavensfruit.com; mainassistant@heavensfruit.com'
                            mailItem._oleobj_.Invoke(
                                *(64209, 0, 8, 0, olNS.Accounts.Item('subgerencia@heavensfruit.com')))
                            mailItem.Display()
                        elif seleccion == 3:
                            print(f'SOLICITUD CERTIFICADO DE ORIGEN: {Kuehne}')
                            print(f'Los documentos a adjuntar son: FACTURA QR.')
                            olApp = win32.Dispatch('Outlook.Application')
                            olNS = olApp.GetnameSpace('MAPI')
                            mailItem = olApp.CreateItem(0)
                            mailItem.Subject = (f'SOLICITUD CO {datos[0]} - {clientenombre} - {clientenombre2}')
                            mailItem.BodyFormat = 2
                            mailItem.Body = 'Cordial saludo,\nAdjunto Facturas QR para generación de Certificados de Origen, \nAgradezco su colaboración.\nCordialmente.'
                            mailItem.To = '' + Kuehne_Co
                            mailItem.Cc = 'gerencia@heavensfruit.com; mabdime@heavensfruit.com; valentinagaray@heavensfruit.com; mainassistant@heavensfruit.com'
                            mailItem._oleobj_.Invoke(
                                *(64209, 0, 8, 0, olNS.Accounts.Item('subgerencia@heavensfruit.com')))
                            mailItem.Display()
                        elif seleccion == 4:
                            print(f'SOLICITUD CERTIFICADO DE ORIGEN: {Logiztik_a}')
                            print(f'Los documentos a adjuntar son: FACTURA QR.')
                            olApp = win32.Dispatch('Outlook.Application')
                            olNS = olApp.GetnameSpace('MAPI')
                            mailItem = olApp.CreateItem(0)
                            mailItem.Subject = (f'SOLICITUD CO {datos[0]} - {clientenombre} - {clientenombre2}')
                            mailItem.BodyFormat = 2
                            mailItem.Body = 'Cordial saludo,\nAdjunto Facturas QR para generación de Certificados de Origen, \nAgradezco su colaboración.\nCordialmente.'
                            mailItem.To = '' + Logiztik_a_CO
                            mailItem.Cc = 'gerencia@heavensfruit.com; mabdime@heavensfruit.com; valentinagaray@heavensfruit.com; mainassistant@heavensfruit.com'
                            mailItem._oleobj_.Invoke(
                                *(64209, 0, 8, 0, olNS.Accounts.Item('subgerencia@heavensfruit.com')))
                            mailItem.Display()
                        else:
                            print("Opción no valida")
            else:
                print(f"No se encontro ningun cliente con el codigo: {codigo}")
                break
    else:
        print("No existen datos.")

# ---------------------------------------------------------------------------------------------------------------------
def fitos():
    os.system('cls')
    seleccion = 0
    while seleccion != 5:
        print(f'''HA SELECCIONADO FITOS.
        Seleccione la agencia de carga:
        1) {Carguex_Sa}
        2) {Cargo_logis}
        3) {Kuehne}
        4) {Logiztik_a}
        5) Volver ...''')
        seleccion = int(input("Seleccione una opción: "))
        if seleccion == 1:
            print(f'SOLICITUD FITOS: {Carguex_Sa}')
            print(f'Los documentos a adjuntar son: FACTURA QR.')
            olApp = win32.Dispatch('Outlook.Application')
            olNS = olApp.GetnameSpace('MAPI')
            mailItem = olApp.CreateItem(0)
            mailItem.Subject = 'FITOS'
            mailItem.BodyFormat = 2
            mailItem.Body = 'Cordial saludo,\nAdjunto Fito para correspondiente pago, \nAgradezco su colaboración.\nCordialmente.'
            mailItem.To = '' + Carguex_Sa_Co
            mailItem.Cc = 'gerencia@heavensfruit.com; mabdime@heavensfruit.com; valentinagaray@heavensfruit.com; mainassistant@heavensfruit.com'
            mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('subgerencia@heavensfruit.com')))
            mailItem.Display()
        elif seleccion == 2:
            print(f'SOLICITUD FITOS: {Cargo_logis}')
            print(f'Los documentos a adjuntar son: FACTURA QR.')
            olApp = win32.Dispatch('Outlook.Application')
            olNS = olApp.GetnameSpace('MAPI')
            mailItem = olApp.CreateItem(0)
            mailItem.Subject = 'FITOS'
            mailItem.BodyFormat = 2
            mailItem.Body = 'Cordial saludo,\nAdjunto Fito para correspondiente pago, \nAgradezco su colaboración.\nCordialmente.'
            mailItem.To = '' + Cargo_logis_Co
            mailItem.Cc = 'gerencia@heavensfruit.com; mabdime@heavensfruit.com; valentinagaray@heavensfruit.com; mainassistant@heavensfruit.com'
            mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('subgerencia@heavensfruit.com')))
            mailItem.Display()
        elif seleccion == 3:
            print(f'SOLICITUD FITOS: {Kuehne}')
            print(f'Los documentos a adjuntar son: FACTURA QR.')
            olApp = win32.Dispatch('Outlook.Application')
            olNS = olApp.GetnameSpace('MAPI')
            mailItem = olApp.CreateItem(0)
            mailItem.Subject = 'FITOS'
            mailItem.BodyFormat = 2
            mailItem.Body = 'Cordial saludo,\nAdjunto Fito para correspondiente pago, \nAgradezco su colaboración.\nCordialmente.'
            mailItem.To = '' + Kuehne_Co
            mailItem.Cc = 'gerencia@heavensfruit.com; mabdime@heavensfruit.com; valentinagaray@heavensfruit.com; mainassistant@heavensfruit.com'
            mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('subgerencia@heavensfruit.com')))
            mailItem.Display()
        elif seleccion == 4:
            print(f'SOLICITUD CERTIFICADO DE ORIGEN: {Logiztik_a}')
            print(f'Los documentos a adjuntar son: FACTURA QR.')
            olApp = win32.Dispatch('Outlook.Application')
            olNS = olApp.GetnameSpace('MAPI')
            mailItem = olApp.CreateItem(0)
            mailItem.Subject = 'FITOS'
            mailItem.BodyFormat = 2
            mailItem.Body = 'Cordial saludo,\nAdjunto Fito para correspondiente pago, \nAgradezco su colaboración.\nCordialmente.'
            mailItem.To = '' + Logiztik_a_CO
            mailItem.Cc = 'gerencia@heavensfruit.com; mabdime@heavensfruit.com; valentinagaray@heavensfruit.com; mainassistant@heavensfruit.com'
            mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('subgerencia@heavensfruit.com')))
            mailItem.Display()
        else:
            print("Opción no valida")


# ---------------------------------------------------------------------------------------------------------------------
def envio_documentos(clientes):
    os.system('cls')
    print(f"                    ENVÍO DE DOCUMENTOS CLIENTE.")
    seleccion = 0
    if len(clientes) > 0:
        codigo = input("Escriba el codigo del cliente: ")
        for co_cliente, datos in clientes.items():
            if codigo in co_cliente:
                print(f'Usted a seleccionado a: {datos[0]}')
                awb = input("Ingrese el AWB: ")
                print(f'Los documentos que debe adjuntar son: {datos[2]}')
                olApp = win32.Dispatch('Outlook.Application')
                olNS = olApp.GetnameSpace('MAPI')
                mailItem = olApp.CreateItem(0)
                mailItem.Subject = f'{datos[0]} AWB: {awb}'
                mailItem.BodyFormat = 2
                mailItem.Body = 'Estimado cliente,\nEn el adjunto encontrará copia de los documentos. Cualquier pregunta no dude en contactarnos.\nCordialmente.\n\n' \
                                'Dear Customer, \n In the attachment you will find copy of the documents. Any questions feel free to contact us.\n ¡Best Regards!'
                mailItem.To = '' + datos[1]
                mailItem.Cc = 'gerencia@heavensfruit.com; mabdime@heavensfruit.com; valentinagaray@heavensfruit.com; mainassistant@heavensfruit.com'
                mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('subgerencia@heavensfruit.com')))
                mailItem.Display()



# ----------------------------------------------------------------------------------------------------------------------
def main():
    continuar = True
    clientes = dict()
    nombre_archivo = 'Clientes_heavens.txt'
    cargar_clientes(clientes, nombre_archivo)
    while continuar:
        mostrar_menu()
        opc = int(input("Escriba una de las opciones:"))

        if opc == SOLICITUD_MUISCA_OA:
            solicitud_muisca(clientes)
        elif opc == AGREGAR_CLIENTES:
            agregar_clientes(clientes, nombre_archivo)
        elif opc == VER_CLIENTES:
            ver_clientes(clientes)
        elif opc == SOLICITUD_CO:
            solicitud_co(clientes)
        elif opc == FITOS:
            fitos()
        elif opc == ENVIO_DOCUMENTOS:
            envio_documentos(clientes)
        elif opc == SALIR:
            continuar = False
        else:
            print("Opción no valida. ")

        input('Presiona enter para continuar....')
    print('Nos vemos.')


if __name__ == '__main__':
    main()
