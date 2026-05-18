from datetime import date, datetime
from abc import ABC, abstractmethod
from openpyxl import Workbook

### CLASE ABSTRACTA REPORTEGASTOS

class ReporteGastos(ABC):
    def __init__(self, gastos):
        self._gastos = gastos

    @property
    def gastos(self):
        return self._gastos

    @abstractmethod
    def generar_reporte(self):
        pass


# Area


class Area:
    def __init__(self, nombre):
        self._nombre = nombre

    @property
    def nombre(self):
        return self._nombre


# Categoria


class Categoria:
    def __init__(self, nombre):
        self._nombre = nombre

    @property
    def nombre(self):
        return self._nombre

# Proveedor


class Proveedor:
    def __init__(self, nombre, contacto, RUC=str):
        self._nombre = nombre
        self._contacto = contacto
        self._RUC = RUC

    @property
    def nombre(self):
        return self._nombre

    @nombre.setter
    def nombre(self, value):
        self._nombre = value

    @property
    def contacto(self):
        return self._contacto

    @contacto.setter
    def contacto(self, value):
        self._contacto = value

    @property
    def RUC(self):
        return self._RUC

    @RUC.setter
    def RUC(self, value):
        self._RUC = value


# Usuario


class Usuario:
    def __init__(self, nombre, email):
        self._nombre = nombre
        self._email = email

    @property
    def nombre(self):
        return self._nombre

    @property
    def email(self):
        return self._email

### CLASES HIJAS DE USUSRARIO

# CLASE EMPLEADO

class Empleado(Usuario):
    def __init__(self, nombre, email, area):
        super().__init__(nombre, email)
        self._area = area

    @property
    def area(self):
        return self._area

# CLASE ADMINISTRADOR

class Administrador(Usuario):
    def __init__(self, nombre, email, privilegios):
        super().__init__(nombre, email)
        self._privilegios = privilegios

    @property
    def privilegios(self):
        return self._privilegios

# CLASE GERENTE
class Gerente(Usuario):
    def __init__(self, nombre, email, departamento):
        super().__init__(nombre, email)
        self._departamento = departamento

    @property
    def departamento(self):
        return self._departamento

### CLASE GASTO Y SUS HIJAS (POR DEPARTAMENTO)

# CLASE GASTO

class Gasto:
    def __init__(self, monto, fecha, categoria, proveedor, usuario, area, concepto, comprobante):
        if monto <= 0:
            raise ValueError("El monto debe ser mayor a cero.")

        self._monto = monto
        self._fecha = fecha
        self._categoria = categoria
        self._proveedor = proveedor
        self._usuario = usuario
        self._area = area
        self._concepto = concepto
        self._comprobante = comprobante

    @property
    def monto(self):
        return self._monto

    @property
    def fecha(self):
        return self._fecha

    @property
    def categoria(self):
        return self._categoria

    @property
    def proveedor(self):
        return self._proveedor

    @property
    def usuario(self):
        return self._usuario

    @property
    def area(self):
        return self._area

    @property
    def concepto(self):
        return self._concepto

    @property
    def comprobante(self):
        return self._comprobante

    def calcular_total(self):
        return self.monto

    def obtener_tipo(self) -> str:
        return "Gasto general"

    def __str__(self):
        return (
            f"{self.fecha} | {self.obtener_tipo()} | "
            f"Área: {self.area.nombre} | "
            f"Categoría: {self.categoria.nombre} | "
            f"Proveedor: {self.proveedor.nombre} | "
            f"Usuario: {self.usuario.nombre} | "
            f"Total: S/ {self.calcular_total():.2f}"
        )


class GastoMarketing(Gasto):
    def __init__(self, monto, fecha, categoria, proveedor, usuario, area, concepto,  comprobante, campania):
        super().__init__(monto, fecha, categoria, proveedor, usuario, area, concepto, comprobante)
        self._campania = campania

    @property
    def campania(self):
        return self._campania

    def calcular_total(self):
        recargo = self.monto * 0.05
        return self.monto + recargo

    def obtener_tipo(self):
        return "Gasto de marketing"


class GastoAdministracion(Gasto):
    def __init__(self, monto, fecha, categoria, proveedor, usuario, area, concepto, comprobante, proyecto):
        super().__init__(monto, fecha, categoria, proveedor, usuario, area, concepto, comprobante)
        self._proyecto = proyecto

    @property
    def proyecto(self):
        return self._proyecto

    def calcular_total(self):
        return self.monto

    def obtener_tipo(self):
        return "Gasto de administración"


class GastoLogistica(Gasto):
    def __init__(self, monto, fecha, categoria, proveedor, usuario, area, concepto, comprobante, ruta, almacen, costo_transporte):
        super().__init__(monto, fecha, categoria, proveedor, usuario, area, concepto, comprobante)

        if costo_transporte < 0:
            raise ValueError("El costo de transporte no puede ser negativo.")

        self._ruta = ruta
        self._almacen = almacen
        self._costo_transporte = costo_transporte

    @property
    def ruta(self):
        return self._ruta

    @property
    def almacen(self):
        return self._almacen

    @property
    def costo_transporte(self):
        return self._costo_transporte

    def calcular_total(self):
        return self.monto + self.costo_transporte

    def obtener_tipo(self) -> str:
        return "Gasto de logística"
### CLASE PRESUPUESTO

class Presupuesto:
    def __init__(self, departamento, monto_asignado, fecha_inicio, fecha_fin):
        self._departamento = departamento
        self._monto_asignado = monto_asignado
        self._fecha_inicio = fecha_inicio
        self._fecha_fin = fecha_fin

    @property
    def departamento(self):
        return self._departamento

    @property
    def monto_asignado(self):
        return self._monto_asignado

    @property
    def fecha_inicio(self):
        return self._fecha_inicio

    @property
    def fecha_fin(self):
        return self._fecha_fin


# 6. CLASE SISTEMA_GASTOS

class SistemaGastos:
    def __init__(self):
        self._areas = []
        self._categorias = []
        self._proveedores = []
        self._usuarios = []
        self._gastos = []
        self._presupuestos = []

    @property
    def areas(self):
        return self._areas

    @property
    def categorias(self):
        return self._categorias

    @property
    def proveedores(self):
        return self._proveedores

    @property
    def usuarios(self):
        return self._usuarios

    @property
    def gastos(self):
        return self._gastos

    @property
    def presupuestos(self):
        return self._presupuestos

    def agregar_area(self, area):
        self._areas.append(area)

    def agregar_categoria(self, categoria):
        self._categorias.append(categoria)

    def agregar_proveedor(self, proveedor):
        self._proveedores.append(proveedor)

    def agregar_usuario(self, usuario):
        self._usuarios.append(usuario)

    def agregar_gasto(self, gasto):
        self._gastos.append(gasto)

    def comprobante_existe(self, comprobante):
        for gasto in self._gastos:
            if gasto.comprobante.lower() == comprobante.lower():
                return True

        return False

    def agregar_presupuesto(self, presupuesto):
        self._presupuestos.append(presupuesto)

    def puede_registrar_gasto(self, area, monto_nuevo):
        total_actual = self.calcular_gasto_total_por_area(area.nombre)

        for presupuesto in self._presupuestos:
            if presupuesto.departamento.lower() == area.nombre.lower():

                total_final = total_actual + monto_nuevo

                if total_final > presupuesto.monto_asignado:
                    return False

                return True

        return True

    def listar_gastos(self):
        if len(self._gastos) == 0:
            print("No hay gastos registrados.")
            return

        for gasto in self._gastos:
            print(gasto)

    def calcular_gasto_total_por_area(self, nombre_area):
        total = 0

        for gasto in self._gastos:
            if gasto.area.nombre.lower() == nombre_area.lower():
                total += gasto.calcular_total()

        return total

    def calcular_gasto_total_general(self):
        total = 0

        for gasto in self._gastos:
            total += gasto.calcular_total()

        return total

    def buscar_gastos_por_area(self, nombre_area):
        resultados = []

        for gasto in self._gastos:
            if gasto.area.nombre.lower() == nombre_area.lower():
                resultados.append(gasto)

        return resultados
    
    def buscar_gastos_por_categoria(self, nombre_categoria):
        resultados = []

        for gasto in self._gastos:
            if gasto.categoria.nombre.lower() == nombre_categoria.lower():
                resultados.append(gasto)

        return resultados

    def buscar_gastos_por_proveedor(self, nombre_proveedor):
        resultados = []

        for gasto in self._gastos:
            if gasto.proveedor.nombre.lower() == nombre_proveedor.lower():
                resultados.append(gasto)

        return resultados

    def buscar_gastos_por_fecha(self, fecha_busqueda):
        resultados = []

        for gasto in self._gastos:
            if gasto.fecha == fecha_busqueda:
                resultados.append(gasto)

        return resultados
    
    def verificar_presupuesto(self, nombre_area):
        gasto_total = self.calcular_gasto_total_por_area(nombre_area)

        for presupuesto in self._presupuestos:
            if presupuesto.departamento.lower() == nombre_area.lower():
                return gasto_total <= presupuesto.monto_asignado

        return True
    
    def exportar_gastos_excel(self, nombre_archivo="reporte_gastos.xlsx"):
        workbook = Workbook()
        hoja = workbook.active
        hoja.title = "Reporte Gastos"
        encabezados = [
            "Fecha",
            "Tipo",
            "Área",
            "Categoría",
            "Proveedor",
            "Usuario",
            "Concepto",
            "Total"
        ]

        hoja.append(encabezados)
        for gasto in self._gastos:

            fila = [
                str(gasto.fecha),
                gasto.obtener_tipo(),
                gasto.area.nombre,
                gasto.categoria.nombre,
                gasto.proveedor.nombre,
                gasto.usuario.nombre,
                gasto.concepto,
                gasto.calcular_total()
            ]

            hoja.append(fila)

        workbook.save(nombre_archivo)

## CLASES DE REPORTE DE GASTOS

class reporte_filtrados(ReporteGastos):
    def __init__(
        self,
        gastos,
        area=None,
        categoria=None,
        proveedor=None,
        fecha_inicio=None,
        fecha_fin=None,
        tipo_gasto=None
    ):
        super().__init__(gastos)
        self._area = area
        self._categoria = categoria
        self._proveedor = proveedor
        self._fecha_inicio = fecha_inicio
        self._fecha_fin = fecha_fin
        self._tipo_gasto = tipo_gasto

    def cumple_filtros(self, gasto):
        if self._area and gasto.area.nombre.lower() != self._area.lower():
            return False
        if self._categoria and gasto.categoria.nombre.lower() != self._categoria.lower():
            return False
        if self._proveedor and gasto.proveedor.nombre.lower() != self._proveedor.lower():
            return False
        if self._fecha_inicio and gasto.fecha < self._fecha_inicio:
            return False
        if self._fecha_fin and gasto.fecha > self._fecha_fin:
            return False
        if self._tipo_gasto and gasto.obtener_tipo().lower() != self._tipo_gasto.lower():
            return False
        return True
    
    def generar_reporte(self) -> None:
        total = 0
        cantidad = 0

        print("\n--- REPORTE FILTRADO ---")

        for gasto in self.gastos:
            if self.cumple_filtros(gasto):
                print(gasto)
                total += gasto.calcular_total()
                cantidad += 1

        if cantidad == 0:
            print("No se encontraron gastos con los filtros ingresados.")
        else:
            print("\nResumen:")
            print(f"Cantidad de gastos encontrados: {cantidad}")
            print(f"Total filtrado: S/ {total:.2f}")

class ReportePorArea(ReporteGastos):

    def __init__(self, gastos, area):
        super().__init__(gastos)
        self._area = area

    def generar_reporte(self):

        total = 0
        cantidad = 0

        print(f"\n--- REPORTE DEL ÁREA {self._area.upper()} ---")

        for gasto in self.gastos:

            if gasto.area.nombre.lower() == self._area.lower():

                print(gasto)

                total += gasto.calcular_total()
                cantidad += 1

        if cantidad == 0:
            print("No se encontraron gastos.")
        else:
            print(f"\nCantidad de gastos: {cantidad}")
            print(f"Total del área: S/ {total:.2f}")

class ReportePorCategoria(ReporteGastos):

    def __init__(self, gastos, categoria):
        super().__init__(gastos)
        self._categoria = categoria

    def generar_reporte(self):

        total = 0
        cantidad = 0

        print(f"\n--- REPORTE DE CATEGORÍA {self._categoria.upper()} ---")

        for gasto in self.gastos:

            if gasto.categoria.nombre.lower() == self._categoria.lower():

                print(gasto)

                total += gasto.calcular_total()
                cantidad += 1

        if cantidad == 0:
            print("No se encontraron gastos.")
        else:
            print(f"\nCantidad de gastos: {cantidad}")
            print(f"Total categoría: S/ {total:.2f}")

class ReportePorFecha(ReporteGastos):

    def __init__(self, gastos, fecha):
        super().__init__(gastos)
        self._fecha = fecha

    def generar_reporte(self):

        total = 0
        cantidad = 0

        print(f"\n--- REPORTE DE FECHA {self._fecha} ---")

        for gasto in self.gastos:

            if gasto.fecha == self._fecha:

                print(gasto)

                total += gasto.calcular_total()
                cantidad += 1

        if cantidad == 0:
            print("No se encontraron gastos.")
        else:
            print(f"\nCantidad de gastos: {cantidad}")
            print(f"Total fecha: S/ {total:.2f}")

def mostrar_menu():
    print("\n========== SISTEMA DE REGISTRO DE GASTOS ==========")
    print("1. Registrar gasto")
    print("2. Listar gastos")
    print("3. Buscar gastos por área")
    print("4. Buscar gastos por categoría")
    print("5. Buscar gastos por proveedor")
    print("6. Buscar gastos por fecha")
    print("7. Ver total general")
    print("8. Ver presupuesto por área")
    print("9. Exportar reporte Excel")
    print("10. Salir")
    print("===================================================")


def pedir_opcion():
    try:
        return int(input("Ingrese una opción: "))
    except ValueError:
        print("Debe ingresar un número.")
        return 0


def buscar_area_por_nombre(sistema, nombre_area):
    for area in sistema.areas:
        if area.nombre.lower() == nombre_area.lower():
            return area
    return None


def buscar_categoria_por_nombre(sistema, nombre_categoria):
    for categoria in sistema.categorias:
        if categoria.nombre.lower() == nombre_categoria.lower():
            return categoria
    return None


def buscar_proveedor_por_nombre(sistema, nombre_proveedor):
    for proveedor in sistema.proveedores:
        if proveedor.nombre.lower() == nombre_proveedor.lower():
            return proveedor
    return None

def pedir_texto(mensaje: str):
    while True:
        dato = input(mensaje)

        if dato == "0":
            return None

        if dato.strip() == "":
            print("Error: este campo no puede estar vacío.")
        else:
            return dato


def pedir_numero(mensaje: str):
    while True:
        dato = input(mensaje)

        if dato == "0":
            return None

        try:
            numero = float(dato)

            if numero <= 0:
                print("Error: el número debe ser mayor a cero.")
            else:
                return numero

        except ValueError:
            print("Error: debe ingresar un número válido.")


def pedir_opcion_menu(mensaje: str, minimo: int, maximo: int):
    while True:
        dato = input(mensaje)

        if dato == "0":
            return None

        try:
            opcion = int(dato)

            if opcion < minimo or opcion > maximo:
                print(f"Error: ingrese una opción entre {minimo} y {maximo}.")
            else:
                return opcion

        except ValueError:
            print("Error: debe ingresar un número entero.")


def cancelar_registro():
    print("Registro cancelado. Volviendo al menú principal...")


def registrar_gasto_menu(sistema):
    print("\n--- REGISTRAR GASTO ---")
    print("Escriba 0 en cualquier momento para cancelar y volver al menú principal.")

    if len(sistema.usuarios) == 0:
        print("No hay usuarios registrados.")
        return

    if len(sistema.areas) == 0:
        print("No hay áreas registradas.")
        return

    if len(sistema.categorias) == 0:
        print("No hay categorías registradas.")
        return

    if len(sistema.proveedores) == 0:
        print("No hay proveedores registrados.")
        return

    concepto = pedir_texto("Ingrese concepto del gasto: ")
    if concepto is None:
        cancelar_registro()
        return

    comprobante = pedir_texto("Ingrese número de comprobante: ")
    if comprobante is None:
        cancelar_registro()
        return
    if sistema.comprobante_existe(comprobante):
        print("Error: el comprobante ya existe.")
        return

    monto = pedir_numero("Ingrese monto del gasto: ")
    if monto is None:
        cancelar_registro()
        return

    while True:
        print("\nÁREAS DISPONIBLES:")
        for area in sistema.areas:
            print("-", area.nombre)

        nombre_area = pedir_texto("Ingrese área: ")
        if nombre_area is None:
            cancelar_registro()
            return

        area = buscar_area_por_nombre(sistema, nombre_area)

        if area is None:
            print("Área no encontrada. Intente nuevamente.")
        else:
            break

    while True:
        print("\nCATEGORÍAS DISPONIBLES:")
        for categoria in sistema.categorias:
            print("-", categoria.nombre)

        nombre_categoria = pedir_texto("Ingrese categoría: ")
        if nombre_categoria is None:
            cancelar_registro()
            return

        categoria = buscar_categoria_por_nombre(sistema, nombre_categoria)

        if categoria is None:
            print("Categoría no encontrada. Intente nuevamente.")
        else:
            break

    while True:
        print("\nPROVEEDORES DISPONIBLES:")
        for proveedor in sistema.proveedores:
            print("-", proveedor.nombre, " RUC :", proveedor._RUC)

        nombre_proveedor = pedir_texto("Ingrese proveedor: ")
        if nombre_proveedor is None:
            cancelar_registro()
            return

        proveedor = buscar_proveedor_por_nombre(sistema, nombre_proveedor)

        if proveedor is None:
            print("Proveedor no encontrado. Intente nuevamente.")
        else:
            break

    usuario = sistema.usuarios[0]

    print("\nTIPO DE GASTO:")
    print("1. Marketing")
    print("2. Administración")
    print("3. Logística")

    tipo = pedir_opcion_menu("Seleccione el tipo de gasto: ", 1, 3)

    gasto=None

    if tipo is None:
        cancelar_registro()
        return

    if not sistema.puede_registrar_gasto(area, monto):
        total_actual = sistema.calcular_gasto_total_por_area(area.nombre)
        for presupuesto in sistema.presupuestos:
            if presupuesto.departamento.lower() == area.nombre.lower():
                disponible = presupuesto.monto_asignado - total_actual
                print("Error: el gasto supera el presupuesto disponible.")
                print(f"Disponible actual: S/ {disponible:.2f}")
                return

    try:
        if tipo == 1:
            campania = pedir_texto("Ingrese campaña: ")

            if campania is None:
                cancelar_registro()
                return

            gasto = GastoMarketing(
                monto,
                date.today(),
                categoria,
                proveedor,
                usuario,
                area,
                concepto,
                comprobante,
                campania
            )

        elif tipo == 2:
            proyecto = pedir_texto("Ingrese proyecto: ")

            if proyecto is None:
                cancelar_registro()
                return

            gasto = GastoAdministracion(
                monto,
                date.today(),
                categoria,
                proveedor,
                usuario,
                area,
                concepto,
                comprobante,
                proyecto
            )

        elif tipo == 3:
            ruta = pedir_texto("Ingrese ruta: ")

            if ruta is None:
                cancelar_registro()
                return

            almacen = pedir_texto("Ingrese almacén: ")

            if almacen is None:
                cancelar_registro()
                return

            costo_transporte = pedir_numero("Ingrese costo de transporte: ")

            if costo_transporte is None:
                cancelar_registro()
                return

            gasto = GastoLogistica(
                monto,
                date.today(),
                categoria,
                proveedor,
                usuario,
                area,
                concepto,
                comprobante,
                ruta,
                almacen,
                costo_transporte
            )

        sistema.agregar_gasto(gasto)
        print("Gasto registrado correctamente.")

    except ValueError as error:
        print("Error al registrar gasto:", error)

def listar_gastos_menu(sistema):
    print("\n--- LISTADO DE GASTOS ---")
    sistema.listar_gastos()


def buscar_gastos_por_area_menu(sistema):
    print("\n--- BUSCAR GASTOS POR ÁREA ---")
    nombre_area = input("Ingrese el área a buscar: ")

    reporte = ReportePorArea(
        sistema.gastos,
        nombre_area
    )

    reporte.generar_reporte()

def buscar_gastos_por_categoria_menu(sistema):

    print("\n--- BUSCAR GASTOS POR CATEGORÍA ---")

    nombre_categoria = input("Ingrese la categoría a buscar: ")

    reporte = ReportePorCategoria(
        sistema.gastos,
        nombre_categoria
    )

    reporte.generar_reporte()

def buscar_gastos_por_proveedor_menu(sistema):
    print("\n--- BUSCAR GASTOS POR PROVEEDOR ---")

    nombre_proveedor = input("Ingrese el proveedor a buscar: ")

    resultados = sistema.buscar_gastos_por_proveedor(nombre_proveedor)

    if len(resultados) == 0:
        print("No se encontraron gastos para ese proveedor.")
        return

    total = 0

    for gasto in resultados:
        print(gasto)
        total += gasto.calcular_total()

    print(f"\nTotal encontrado: S/ {total:.2f}")

def buscar_gastos_por_fecha_menu(sistema):

    print("\n--- BUSCAR GASTOS POR FECHA ---")

    fecha_texto = input("Ingrese la fecha (YYYY-MM-DD): ")

    try:

        fecha_busqueda = datetime.strptime(
            fecha_texto,
            "%Y-%m-%d"
        ).date()

    except ValueError:
        print("Error: formato inválido.")
        return

    reporte = ReportePorFecha(
        sistema.gastos,
        fecha_busqueda
    )

    reporte.generar_reporte()

def ver_total_general_menu(sistema):
    total = sistema.calcular_gasto_total_general()
    print(f"\nTotal general de gastos: S/ {total:.2f}")


def ver_presupuesto_menu(sistema):
    print("\n--- VER PRESUPUESTO POR ÁREA ---")
    nombre_area = input("Ingrese área: ")

    total = sistema.calcular_gasto_total_por_area(nombre_area)

    presupuesto_encontrado = None

    for presupuesto in sistema.presupuestos:
        if presupuesto.departamento.lower() == nombre_area.lower():
            presupuesto_encontrado = presupuesto
            break

    if presupuesto_encontrado is None:
        print("No hay presupuesto registrado para esa área.")
        return

    disponible = presupuesto_encontrado.monto_asignado - total
    porcentaje = total / presupuesto_encontrado.monto_asignado * 100

    print(f"Área: {nombre_area}")
    print(f"Presupuesto asignado: S/ {presupuesto_encontrado.monto_asignado:.2f}")
    print(f"Total gastado: S/ {total:.2f}")
    print(f"Disponible: S/ {disponible:.2f}")
    print(f"Porcentaje usado: {porcentaje:.2f}%")

    if porcentaje >= 100:
        print("Estado: presupuesto excedido.")
    elif porcentaje >= 80:
        print("Estado: alerta, presupuesto usado al 80% o más.")
    else:
        print("Estado: presupuesto controlado.")

def exportar_excel_menu(sistema):
    print("\n--- EXPORTAR REPORTE EXCEL ---")

    nombre = input(
        "Ingrese nombre del archivo: "
    )
    if nombre.strip() == "":
        nombre = "reporte_gastos"

    if not nombre.endswith(".xlsx"):
        nombre += ".xlsx"

    try:
        sistema.exportar_gastos_excel(nombre)
        print(
            f"Reporte exportado correctamente: {nombre}"
        )
    except Exception as error:

        print(
            "Error al exportar:",
            error
        )

def ejecutar_menu(sistema):
    while True:
        mostrar_menu()
        opcion = pedir_opcion()

        if opcion == 1:
            registrar_gasto_menu(sistema)
        elif opcion == 2:
            listar_gastos_menu(sistema)
        elif opcion == 3:
            buscar_gastos_por_area_menu(sistema)
        elif opcion == 4:
            buscar_gastos_por_categoria_menu(sistema)
        elif opcion == 5:
            buscar_gastos_por_proveedor_menu(sistema)
        elif opcion == 6:
            buscar_gastos_por_fecha_menu(sistema)
        elif opcion == 7:
            ver_total_general_menu(sistema)
        elif opcion == 8:
            ver_presupuesto_menu(sistema)
        elif opcion == 9:
            exportar_excel_menu(sistema)
        elif opcion == 10:
            print("Saliendo del sistema...")
            break
        else:
            print("Opción no válida.")


sistema = SistemaGastos()

area_marketing = Area("Marketing")
area_administracion = Area("Administración")
area_logistica = Area("Logística")

categoria_publicidad = Categoria("Publicidad")
categoria_servicios = Categoria("Servicios")
categoria_transporte = Categoria("Transporte")

proveedor_1 = Proveedor("Proveedor Dental SAC", "contacto@proveedor.com")

usuario_1 = Empleado("Ángel Navarro", "angel@empresa.com", area_marketing)

presupuesto_marketing = Presupuesto("Marketing", 5000, date(2026, 5, 1), date(2026, 5, 31))
presupuesto_administracion = Presupuesto("Administración", 3000, date(2026, 5, 1), date(2026, 5, 31))
presupuesto_logistica = Presupuesto("Logística", 4000, date(2026, 5, 1), date(2026, 5, 31))

sistema.agregar_area(area_marketing)
sistema.agregar_area(area_administracion)
sistema.agregar_area(area_logistica)

sistema.agregar_categoria(categoria_publicidad)
sistema.agregar_categoria(categoria_servicios)
sistema.agregar_categoria(categoria_transporte)

sistema.agregar_proveedor(proveedor_1)

sistema.agregar_usuario(usuario_1)

sistema.agregar_presupuesto(presupuesto_marketing)
sistema.agregar_presupuesto(presupuesto_administracion)
sistema.agregar_presupuesto(presupuesto_logistica)

ejecutar_menu(sistema)
