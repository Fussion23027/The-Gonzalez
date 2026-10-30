import getpass
import hashlib
import hmac
import os
import re
import secrets
import sqlite3
import sys
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

DB = "serviciosTheGonzalez.db"
CARPETA = "cotizaciones"
RESPALDOS = "respaldos"
MAX_RESPALDOS = 30
LIMITE_MONTO = Decimal("1000000000")
IVA_INICIAL = "0.16"
TASAS_IVA = ("0", "8", "16")
ITERACIONES = 200000
ESTADOS = [
    "REGISTRADO", "COTIZACION_GENERADA", "COTIZACION_ENVIADA", "ACEPTADO",
    "RECHAZADO", "ASIGNADO", "PROGRAMADO", "EN_PROCESO", "PENDIENTE_REVISION",
    "CORRECCION_SOLICITADA", "APROBADO", "FACTURADO", "FINALIZADO", "CANCELADO",
]
ESTADOS_CANCELABLES = [
    "REGISTRADO", "ASIGNADO", "COTIZACION_GENERADA", "ACEPTADO", "EN_PROCESO",
    "PENDIENTE_REVISION", "CORRECCION_SOLICITADA", "APROBADO",
]
ROLES_PUESTO = ["Director General", "Supervisor de Campo", "Tecnico", "Programador", "Soporte", "Compras"]
ROLES_REVISOR = ("Supervisor de Campo", "Director General")
ROLES_APP = ("GERENTE", "SUPERVISOR", "COTIZADOR", "CONTADOR")
NOMBRES_ROL = {"GERENTE": "Gerente", "SUPERVISOR": "Supervisor", "COTIZADOR": "Cotizador", "CONTADOR": "Contador"}
SESION = {"id": None, "nombre": None, "nombreCompleto": None, "rol": None}

SCHEMA = """
CREATE TABLE IF NOT EXISTS estado_servicio (
    idEstado INTEGER PRIMARY KEY AUTOINCREMENT,
    nombreEstado TEXT NOT NULL UNIQUE
);
CREATE TABLE IF NOT EXISTS cliente (
    idCliente INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT NOT NULL,
    telefono TEXT,
    correo TEXT,
    direccion TEXT,
    rfc TEXT,
    activo INTEGER NOT NULL DEFAULT 1
);
CREATE TABLE IF NOT EXISTS empleado (
    idEmpleado INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT NOT NULL,
    aPaterno TEXT NOT NULL,
    aMaterno TEXT,
    telefono TEXT,
    correo TEXT,
    activo INTEGER NOT NULL DEFAULT 1
);
CREATE TABLE IF NOT EXISTS rol (
    idRol INTEGER PRIMARY KEY AUTOINCREMENT,
    nombreRol TEXT NOT NULL UNIQUE
);
CREATE TABLE IF NOT EXISTS puesto (
    idEmpleado INTEGER NOT NULL,
    idRol INTEGER NOT NULL,
    PRIMARY KEY (idEmpleado, idRol),
    FOREIGN KEY (idEmpleado) REFERENCES empleado(idEmpleado),
    FOREIGN KEY (idRol) REFERENCES rol(idRol)
);
CREATE TABLE IF NOT EXISTS servicio (
    idServicio INTEGER PRIMARY KEY AUTOINCREMENT,
    idCliente INTEGER NOT NULL,
    idEstado INTEGER NOT NULL,
    descripcionProblema TEXT NOT NULL,
    descripcionServicio TEXT,
    observaciones TEXT,
    fechaRegistro TEXT NOT NULL DEFAULT (date('now')),
    FOREIGN KEY (idCliente) REFERENCES cliente(idCliente),
    FOREIGN KEY (idEstado) REFERENCES estado_servicio(idEstado)
);
CREATE TABLE IF NOT EXISTS servicioempleado (
    idServicio INTEGER NOT NULL,
    idEmpleado INTEGER NOT NULL,
    tipoParticipacion TEXT NOT NULL,
    PRIMARY KEY (idServicio, idEmpleado),
    FOREIGN KEY (idServicio) REFERENCES servicio(idServicio),
    FOREIGN KEY (idEmpleado) REFERENCES empleado(idEmpleado)
);
CREATE TABLE IF NOT EXISTS revisionservicio (
    idRevision INTEGER PRIMARY KEY AUTOINCREMENT,
    idServicio INTEGER NOT NULL,
    idSupervisor INTEGER NOT NULL,
    comentarios TEXT,
    aprobado INTEGER NOT NULL,
    FOREIGN KEY (idServicio) REFERENCES servicio(idServicio),
    FOREIGN KEY (idSupervisor) REFERENCES empleado(idEmpleado)
);
CREATE TABLE IF NOT EXISTS retroalimentacion (
    idRetroalimentacion INTEGER PRIMARY KEY AUTOINCREMENT,
    idServicio INTEGER NOT NULL,
    idEmpleado INTEGER NOT NULL,
    comentarios TEXT,
    FOREIGN KEY (idServicio) REFERENCES servicio(idServicio),
    FOREIGN KEY (idEmpleado) REFERENCES empleado(idEmpleado)
);
CREATE TABLE IF NOT EXISTS configuracion (
    clave TEXT PRIMARY KEY,
    valor TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS cotizacion (
    idCotizacion INTEGER PRIMARY KEY AUTOINCREMENT,
    idServicio INTEGER NOT NULL,
    idEmpleado INTEGER NOT NULL,
    fechaCotizacion TEXT NOT NULL DEFAULT (date('now')),
    subtotal REAL NOT NULL DEFAULT 0,
    iva REAL NOT NULL DEFAULT 0,
    tasaIva REAL NOT NULL DEFAULT 0.16,
    total REAL NOT NULL DEFAULT 0,
    observaciones TEXT,
    archivoCotizacion TEXT,
    hashArchivo TEXT,
    estado TEXT NOT NULL DEFAULT 'PENDIENTE',
    FOREIGN KEY (idServicio) REFERENCES servicio(idServicio),
    FOREIGN KEY (idEmpleado) REFERENCES empleado(idEmpleado)
);
CREATE TABLE IF NOT EXISTS detallecotizacion (
    idDetalle INTEGER PRIMARY KEY AUTOINCREMENT,
    idCotizacion INTEGER NOT NULL,
    descripcion TEXT NOT NULL,
    cantidad REAL NOT NULL,
    precioUnitario REAL NOT NULL,
    importe REAL NOT NULL,
    FOREIGN KEY (idCotizacion) REFERENCES cotizacion(idCotizacion)
);
CREATE TABLE IF NOT EXISTS factura (
    idFactura INTEGER PRIMARY KEY AUTOINCREMENT,
    idServicio INTEGER NOT NULL,
    idCliente INTEGER NOT NULL,
    numeroFactura TEXT UNIQUE,
    fechaEmision TEXT NOT NULL DEFAULT (date('now')),
    subtotal REAL NOT NULL DEFAULT 0,
    iva REAL NOT NULL DEFAULT 0,
    total REAL NOT NULL DEFAULT 0,
    estado TEXT NOT NULL DEFAULT 'GENERADA',
    fechaPago TEXT,
    archivoFactura TEXT,
    FOREIGN KEY (idServicio) REFERENCES servicio(idServicio),
    FOREIGN KEY (idCliente) REFERENCES cliente(idCliente)
);
CREATE TABLE IF NOT EXISTS usuario (
    idUsuario INTEGER PRIMARY KEY AUTOINCREMENT,
    nombreUsuario TEXT NOT NULL UNIQUE,
    nombreCompleto TEXT,
    claveHash TEXT NOT NULL,
    rol TEXT NOT NULL,
    activo INTEGER NOT NULL DEFAULT 1,
    creado TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))
);
CREATE TABLE IF NOT EXISTS bitacora (
    idBitacora INTEGER PRIMARY KEY AUTOINCREMENT,
    fecha TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    usuario TEXT NOT NULL,
    accion TEXT NOT NULL,
    tabla TEXT,
    idRegistro INTEGER,
    detalle TEXT
);
"""


def conectar():
    c = sqlite3.connect(DB)
    c.execute("PRAGMA foreign_keys = ON;")
    c.row_factory = sqlite3.Row
    c.executescript(SCHEMA)
    with c:
        agregar_columna(c, "cotizacion", "tasaIva", "REAL NOT NULL DEFAULT 0.16")
        agregar_columna(c, "cotizacion", "hashArchivo", "TEXT")
        agregar_columna(c, "servicio", "motivoCancelacion", "TEXT")
        agregar_columna(c, "usuario", "nombreCompleto", "TEXT")
        c.execute("UPDATE usuario SET rol = 'GERENTE' WHERE rol = 'DIRECTOR';")
        c.executemany("INSERT OR IGNORE INTO estado_servicio (nombreEstado) VALUES (?);", [(e,) for e in ESTADOS])
        c.executemany("INSERT OR IGNORE INTO rol (nombreRol) VALUES (?);", [(r,) for r in ROLES_PUESTO])
        c.execute("INSERT OR IGNORE INTO configuracion (clave, valor) VALUES ('IVA', ?);", (IVA_INICIAL,))
    return c


def agregar_columna(c, tabla, columna, definicion):
    existentes = [fila["name"] for fila in c.execute(f"PRAGMA table_info({tabla});")]
    if columna not in existentes:
        c.execute(f"ALTER TABLE {tabla} ADD COLUMN {columna} {definicion};")


class Cancelado(Exception):
    pass


def captura(msg=""):
    v = input(msg).strip()
    if v == "*":
        raise Cancelado()
    return v


def ejecutar(funcion, con):
    try:
        funcion(con)
    except Cancelado:
        print("Operacion cancelada. No se guardo nada.")
    except sqlite3.IntegrityError:
        print("No se pudo completar: un dato relacionado ya no existe. Revise el ID e intente de nuevo.")
    except (sqlite3.Error, OSError) as e:
        print("Error:", e)


def bitacora(con, accion, tabla=None, idRegistro=None, detalle=None):
    con.execute(
        "INSERT INTO bitacora (usuario, accion, tabla, idRegistro, detalle) VALUES (?, ?, ?, ?, ?);",
        (SESION["nombre"] or "sistema", accion, tabla, idRegistro, detalle))


def hashear_clave(clave):
    sal = secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac("sha256", clave.encode(), sal.encode(), ITERACIONES).hex()
    return f"{sal}${digest}"


def verificar_clave(clave, almacenada):
    sal, digest = almacenada.split("$")
    nuevo = hashlib.pbkdf2_hmac("sha256", clave.encode(), sal.encode(), ITERACIONES).hex()
    return hmac.compare_digest(nuevo, digest)


def leer_clave(msg):
    v = getpass.getpass(msg) if sys.stdin.isatty() else input(msg)
    if v == "*":
        raise Cancelado()
    return v


def pedir_clave_nueva():
    print("La clave debe tener al menos 8 caracteres.")
    while True:
        c1 = leer_clave("Clave nueva: ")
        if len(c1) < 8:
            print("Muy corta. Minimo 8 caracteres.")
            continue
        c2 = leer_clave("Repita la clave: ")
        if c1 != c2:
            print("No coinciden. Intente de nuevo.")
            continue
        return c1


def hay_usuarios(con):
    return con.execute("SELECT COUNT(*) FROM usuario;").fetchone()[0] > 0


def pedir_usuario_nuevo(con):
    print("Usuario para entrar: 3 a 20 letras, numeros o _ (ej. jperez).")
    while True:
        nombre = captura("Usuario: ")
        if not re.fullmatch(r"[A-Za-z0-9_]{3,20}", nombre):
            print("Usuario no valido.")
            continue
        if con.execute("SELECT 1 FROM usuario WHERE nombreUsuario = ?;", (nombre,)).fetchone():
            print("Ese usuario ya existe. Elija otro.")
            continue
        return nombre


def pedir_rol():
    for i, r in enumerate(ROLES_APP, 1):
        print(f"  {i}. {NOMBRES_ROL[r]}")
    while True:
        v = captura("Numero del rol: ")
        if v.isascii() and v.isdigit() and 1 <= int(v) <= len(ROLES_APP):
            return ROLES_APP[int(v) - 1]
        print("Escriba un numero de la lista.")


def crear_usuario(con, usuario, nombre_completo, rol, clave):
    with con:
        cur = con.execute(
            "INSERT INTO usuario (nombreUsuario, nombreCompleto, claveHash, rol) VALUES (?, ?, ?, ?);",
            (usuario, nombre_completo, hashear_clave(clave), rol))
        bitacora(con, "CREAR USUARIO", "usuario", cur.lastrowid, f"{usuario} ({NOMBRES_ROL[rol]})")
    return cur.lastrowid


def iniciar_sesion(con, rol_requerido=None):
    try:
        for _ in range(3):
            usuario = captura("Usuario: ")
            clave = leer_clave("Clave: ")
            fila = con.execute(
                """SELECT idUsuario, nombreUsuario, nombreCompleto, claveHash, rol
                   FROM usuario WHERE nombreUsuario = ? AND activo = 1;""", (usuario,)).fetchone()
            if fila and verificar_clave(clave, fila["claveHash"]):
                if rol_requerido and fila["rol"] != rol_requerido:
                    print(f"Este programa es solo para el {NOMBRES_ROL[rol_requerido].lower()}.")
                    return False
                SESION.update({"id": fila["idUsuario"], "nombre": fila["nombreUsuario"],
                               "nombreCompleto": fila["nombreCompleto"] or fila["nombreUsuario"],
                               "rol": fila["rol"]})
                print(f"\nBienvenido, {SESION['nombreCompleto']} ({NOMBRES_ROL[fila['rol']]}).")
                return True
            print("Usuario o clave incorrectos.")
        return False
    except Cancelado:
        return False


def cambiar_mi_clave(con):
    fila = con.execute("SELECT claveHash FROM usuario WHERE idUsuario = ?;", (SESION["id"],)).fetchone()
    if not verificar_clave(leer_clave("Clave actual: "), fila["claveHash"]):
        print("Clave actual incorrecta.")
        return
    nueva = pedir_clave_nueva()
    with con:
        con.execute("UPDATE usuario SET claveHash = ? WHERE idUsuario = ?;", (hashear_clave(nueva), SESION["id"]))
        bitacora(con, "CAMBIAR CLAVE", "usuario", SESION["id"])
    print("Clave actualizada.")


def respaldar(con):
    os.makedirs(RESPALDOS, exist_ok=True)
    ruta = os.path.join(RESPALDOS, f"respaldo_{datetime.now():%Y%m%d_%H%M%S_%f}.db")
    destino = sqlite3.connect(ruta)
    try:
        con.backup(destino)
        estado = destino.execute("PRAGMA integrity_check;").fetchone()[0]
    finally:
        destino.close()
    if estado != "ok":
        raise sqlite3.DatabaseError(f"El respaldo salio con errores: {estado}")
    return ruta


def rotar_respaldos():
    archivos = sorted(f for f in os.listdir(RESPALDOS) if f.startswith("respaldo_") and f.endswith(".db"))
    for viejo in archivos[:-MAX_RESPALDOS]:
        os.remove(os.path.join(RESPALDOS, viejo))


def respaldo_diario(con):
    os.makedirs(RESPALDOS, exist_ok=True)
    prefijo = f"respaldo_{date.today():%Y%m%d}"
    if not any(f.startswith(prefijo) for f in os.listdir(RESPALDOS)):
        respaldar(con)
    rotar_respaldos()


def leer(msg, validar=None, error="Dato no valido.", opcional=False):
    while True:
        v = captura(msg)
        if not v:
            if opcional:
                return None
            print("Este dato es obligatorio.")
            continue
        if validar is None or validar(v):
            return v
        print(error)


def pedir_entero(msg):
    while True:
        v = captura(msg)
        if v.isascii() and v.isdigit():
            return int(v)
        print("Escriba un numero de la lista.")


def numero_desde_texto(v):
    limpio = v.replace("$", "").replace(",", "").replace(" ", "")
    try:
        d = Decimal(limpio)
    except InvalidOperation:
        return None
    return d if d.is_finite() else None


def pedir_monto(msg):
    while True:
        d = numero_desde_texto(captura(msg))
        if d is not None and Decimal("0") < d <= LIMITE_MONTO:
            return d
        print("Escriba un numero mayor a cero (ej. 1500 o 1,500.50).")


def si_no(msg):
    while True:
        v = captura(msg).lower()
        if v in ("s", "n"):
            return v == "s"
        print("Escriba s o n.")


def dinero(valor):
    return Decimal(valor).quantize(Decimal("0.01"), ROUND_HALF_UP)


def nombre_valido(v):
    return re.fullmatch(r"[^\W\d_]+(?:[ '.\-]+[^\W\d_]+)*", v) is not None


def texto_valido(v):
    return len(v) >= 3 and sum(ch.isalpha() for ch in v) >= 2


def telefono_valido(v):
    return re.fullmatch(r"\d{10}", re.sub(r"[\s\-()]", "", v)) is not None


def correo_valido(v):
    return re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", v) is not None


def rfc_valido(v):
    return re.fullmatch(r"[A-ZÑ&]{3,4}\d{6}[A-Z0-9]{3}", v.upper()) is not None


def imprimir_tabla(encabezados, filas):
    if not filas:
        print("No hay registros.")
        return
    matriz = [list(encabezados)] + [["-" if v is None else str(v) for v in f] for f in filas]
    anchos = [max(len(fila[i]) for fila in matriz) for i in range(len(encabezados))]
    for n, fila in enumerate(matriz):
        print("  " + "  ".join(celda.ljust(anchos[i]) for i, celda in enumerate(fila)))
        if n == 0:
            print("  " + "-" * (sum(anchos) + 2 * (len(anchos) - 1)))


def elegir_de(filas, encabezados, pregunta):
    if not filas:
        print("No hay registros disponibles.")
        return None
    imprimir_tabla(encabezados, [tuple(f) for f in filas])
    ids = {f[0] for f in filas}
    while True:
        v = pedir_entero(pregunta)
        if v in ids:
            return v
        print("Ese ID no esta en la lista.")


def imprimir_menu(titulo, opciones, puede):
    print(f"\n=== {titulo} ===")
    print(f"Sesion: {SESION['nombreCompleto']} ({NOMBRES_ROL[SESION['rol']]})")
    ancho = max(len(v[0]) for v in opciones.values())
    for clave, (nombre, _, permisos, ayuda) in opciones.items():
        if puede(permisos):
            print(f"  {clave:>2}. {nombre:<{ancho}}  {ayuda}")
    print("   0. Salir")
    print("   (En cualquier pregunta, escriba * para cancelar)")


def id_estado(con, nombre):
    return con.execute("SELECT idEstado FROM estado_servicio WHERE nombreEstado = ?;", (nombre,)).fetchone()[0]


def obtener_iva(con):
    return float(con.execute("SELECT valor FROM configuracion WHERE clave = 'IVA';").fetchone()[0])


def empleados_activos(con, roles=None):
    if roles is None:
        return con.execute(
            "SELECT idEmpleado, nombre || ' ' || aPaterno FROM empleado WHERE activo = 1 ORDER BY nombre;").fetchall()
    marcas = ",".join("?" * len(roles))
    return con.execute(
        f"""SELECT DISTINCT e.idEmpleado, e.nombre || ' ' || e.aPaterno
            FROM empleado e
            JOIN puesto p ON p.idEmpleado = e.idEmpleado
            JOIN rol r ON r.idRol = p.idRol
            WHERE e.activo = 1 AND r.nombreRol IN ({marcas})
            ORDER BY e.nombre;""", roles).fetchall()


def sha256_archivo(ruta):
    h = hashlib.sha256()
    with open(ruta, "rb") as f:
        for bloque in iter(lambda: f.read(65536), b""):
            h.update(bloque)
    return h.hexdigest()
