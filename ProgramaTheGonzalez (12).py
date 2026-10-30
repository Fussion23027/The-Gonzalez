import functools
import os
import re
import sqlite3
import sys
import hashlib
import shutil
import smtplib
from email.message import EmailMessage
import tempfile
from html import escape
from datetime import date, datetime
from decimal import Decimal, InvalidOperation

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit, QLabel,
    QTableWidget, QTableWidgetItem, QHeaderView, QStackedWidget, QFormLayout, QTextEdit, QMessageBox,
    QListWidget, QListWidgetItem, QAbstractItemView, QFrame, QDialog, QDialogButtonBox,
    QMenu, QComboBox, QRadioButton, QTextBrowser, QSizePolicy, QFileDialog, QCompleter,
)
from PyQt6.QtCore import Qt, QTimer, QUrl, QSize, QSettings, QRect, QRectF
from PyQt6.QtGui import (QPainter, QPen, QColor, QDesktopServices, QFont, QKeySequence,
                         QShortcut, QIcon, QPixmap, QPainterPath)

import nucleo
from nucleo import (
    ESTADOS_CANCELABLES, LIMITE_MONTO, ROLES_REVISOR, SESION, bitacora, conectar, crear_usuario, dinero, empleados_activos,
    numero_desde_texto, obtener_iva, respaldo_diario, sha256_archivo, verificar_clave,
    nombre_valido, texto_valido,
    telefono_valido as telefono_formato, correo_valido as correo_formato,
)

VERSION = "4.0.0"

MSG_OCUPADA = ("La información está siendo usada por otro programa. "
               "Espere unos segundos e intente de nuevo.")
MSG_GENERAL = "No se pudo completar la acción. Intente de nuevo. Si el problema continúa, avise al administrador."
MSG_REGISTRO = "Este registro ya no está disponible. Actualice la lista e intente de nuevo."
MSG_ARCHIVO = ("No se pudo guardar el archivo Excel. Revise que el archivo no esté abierto "
               "en otro programa y vuelva a intentar.")
MSG_BD = ("No se pudo abrir la base de datos. Verifique que el archivo no esté en uso y "
          "vuelva a iniciar el programa.")
ESTADOS_EDITABLES_SERVICIO = ("REGISTRADO", "ASIGNADO", "COTIZACION_GENERADA", "ACEPTADO",
                              "EN_PROCESO", "PENDIENTE_REVISION", "CORRECCION_SOLICITADA", "APROBADO")
ESTADOS_RETROALIMENTACION = ("ASIGNADO", "EN_PROCESO", "PENDIENTE_REVISION", "CORRECCION_SOLICITADA",
                             "APROBADO", "FACTURADO", "FINALIZADO")
ROLES_TEXTO = {"GERENTE": "Gerente", "COTIZADOR": "Cotizador",
               "SUPERVISOR": "Supervisor", "CONTADOR": "Contador"}
ROLES_SISTEMA = ("GERENTE", "COTIZADOR", "SUPERVISOR", "CONTADOR")
ROLES_AUTORREGISTRO = ("COTIZADOR",)
ROLES_LISTA = []


class ReglaError(Exception):
    pass


def telefono_ok(v):
    return not v or telefono_formato(v)


def correo_ok(v):
    return not v or correo_formato(v)


PATRON_RFC = re.compile(r"^[A-ZÑ&]{3,4}\d{6}[A-Z0-9]{3}$")


def rfc_ok(v):
    if not v:
        return True
    texto = str(v).strip().upper()
    if not PATRON_RFC.fullmatch(texto):
        return False
    try:
        mes = int(texto[-7:-5])
        dia = int(texto[-5:-3])
    except ValueError:
        return False
    return 1 <= mes <= 12 and 1 <= dia <= 31


def errores_cliente(d):
    errores = []
    if not nombre_valido(d["nombre"]):
        errores.append(("nombre", "Nombre: solo letras y espacios (ej. Pedro Ramírez)."))
    if not telefono_ok(d["telefono"]):
        errores.append(("telefono", "Teléfono: 10 dígitos sin espacios ni guiones (ej. 8441234567)."))
    if not correo_ok(d["correo"]):
        errores.append(("correo", "Correo: formato nombre@dominio.com (ej. cliente@correo.com)."))
    if not rfc_ok(d["rfc"]):
        errores.append(("rfc", "RFC: 12 o 13 caracteres con letras y fecha (ej. RUIA010101AB1)."))
    if d["direccion"] and not texto_valido(d["direccion"]):
        errores.append(("direccion", "Dirección: incluya al menos 3 letras (ej. Av. Juárez 120)."))
    return errores


def error_cliente(d):
    errores = errores_cliente(d)
    return errores[0] if errores else (None, None)


def validar_cliente(d):
    return error_cliente(d)[1]


def errores_empleado(d):
    errores = []
    if not nombre_valido(d["nombre"]):
        errores.append(("nombre", "Nombre: solo letras y espacios (ej. Juan)."))
    if not nombre_valido(d["ap"]):
        errores.append(("ap", "Apellido paterno: solo letras y espacios (ej. Pérez)."))
    if d["am"] and not nombre_valido(d["am"]):
        errores.append(("am", "Apellido materno: solo letras y espacios (ej. López)."))
    if not telefono_ok(d["telefono"]):
        errores.append(("telefono", "Teléfono: 10 dígitos sin espacios (ej. 8441234567)."))
    if not correo_ok(d["correo"]):
        errores.append(("correo", "Correo: formato nombre@dominio.com (ej. juan@correo.com)."))
    if not d["roles"]:
        errores.append(("roles", "Puestos: seleccione al menos uno."))
    return errores


def error_empleado(d):
    errores = errores_empleado(d)
    return errores[0] if errores else (None, None)


def validar_empleado(d):
    return error_empleado(d)[1]


def errores_servicio(d):
    errores = []
    if not texto_valido(d["problema"]):
        errores.append(("problema", "Problema: al menos 3 letras (ej. Falla la bomba de agua)."))
    return errores


def error_servicio(d):
    errores = errores_servicio(d)
    return errores[0] if errores else (None, None)


def validar_servicio(d):
    return error_servicio(d)[1]


def errores_usuario(d):
    errores = []
    if not nombre_valido(d["nombre"]):
        errores.append(("nombre", "Nombre completo: solo letras y espacios (ej. Juan Pérez)."))
    if not re.fullmatch(r"[A-Za-z0-9_]{3,20}", d["usuario"]):
        errores.append(("usuario", "Usuario: de 3 a 20 caracteres (letras, números o _; ej. jperez)."))
    if d["rol"] not in ROLES_SISTEMA:
        errores.append(("rol", "Puesto: seleccione uno de la lista."))
    if len(d["clave"]) < 8:
        errores.append(("clave", "Contraseña: mínimo 8 caracteres."))
    if d["clave"] != d["clave2"]:
        errores.append(("clave2", "Repetir contraseña: no coincide con la anterior."))
    return errores


def error_usuario(d):
    errores = errores_usuario(d)
    return errores[0] if errores else (None, None)


def validar_usuario(d):
    return error_usuario(d)[1]


def ayuda_campo(texto):
    etiqueta = QLabel(texto)
    etiqueta.setObjectName("ayudaCampo")
    etiqueta.setWordWrap(True)
    return etiqueta


def marcar_error(campo, activo=True, enfocar=True):
    if campo is None:
        return
    campo.setProperty("error", "true" if activo else "false")
    estilo = campo.style()
    if estilo is not None:
        estilo.unpolish(campo)
        estilo.polish(campo)
    if activo and enfocar:
        campo.setFocus()


def marcar_errores(errores, campos):
    for campo, _mensaje in errores:
        marcar_error(campos.get(campo), True, enfocar=False)
    if errores:
        primero = campos.get(errores[0][0])
        if primero is not None:
            primero.setFocus()


def texto_errores(errores):
    return "\n".join(mensaje for _campo, mensaje in errores)


def cambiar_clave_usuario(con, idUsuario, clave):
    """Cambia la contraseña usando la función de nucleo si existe.

    Si nucleo no expone una función para ello, se vuelve a crear la cuenta con
    los mismos datos, de modo que se usa el mismo esquema de cifrado del sistema.
    """
    for nombre in ("cambiar_clave", "actualizar_clave", "cambiar_clave_usuario"):
        funcion = getattr(nucleo, nombre, None)
        if not callable(funcion):
            continue
        try:
            funcion(con, idUsuario, clave)
        except TypeError:
            funcion(con, clave, idUsuario)
        fila = con.execute("SELECT claveHash FROM usuario WHERE idUsuario = ?;",
                           (idUsuario,)).fetchone()
        if fila is not None:
            try:
                if verificar_clave(clave, fila[0]):
                    return
            except (ValueError, TypeError):
                pass
    fila_hash = con.execute("SELECT claveHash FROM usuario WHERE idUsuario = ?;",
                            (idUsuario,)).fetchone()
    if fila_hash and re.fullmatch(r"[0-9a-fA-F]{64}", str(fila_hash[0] or "")):
        con.execute("UPDATE usuario SET claveHash = ? WHERE idUsuario = ?;",
                    (hashlib.sha256(clave.encode()).hexdigest(), idUsuario))
        con.commit()
        return
    columnas = {fila[1] for fila in con.execute("PRAGMA table_info(usuario);").fetchall()}
    campos = ["nombreUsuario", "nombreCompleto", "rol", "activo"]
    for opcional in ("correo", "telefono", "foto"):
        if opcional in columnas:
            campos.append(opcional)
    fila = con.execute(f"SELECT {', '.join(campos)} FROM usuario WHERE idUsuario = ?;",
                       (idUsuario,)).fetchone()
    if fila is None:
        raise ReglaError(MSG_REGISTRO)
    with con:
        con.execute("DELETE FROM usuario WHERE idUsuario = ?;", (idUsuario,))
        crear_usuario(con, fila["nombreUsuario"], fila["nombreCompleto"], fila["rol"], clave)
        nueva = con.execute("SELECT idUsuario FROM usuario WHERE nombreUsuario = ?;",
                            (fila["nombreUsuario"],)).fetchone()
        if nueva is not None and any(c in columnas for c in ("correo", "telefono", "foto")):
            con.execute("UPDATE usuario SET correo = ?, telefono = ?, foto = ? WHERE idUsuario = ?;",
                        (fila["correo"] if "correo" in columnas else None,
                         fila["telefono"] if "telefono" in columnas else None,
                         fila["foto"] if "foto" in columnas else None, nueva[0]))
        if not fila["activo"]:
            con.execute("UPDATE usuario SET activo = 0 WHERE nombreUsuario = ?;",
                        (fila["nombreUsuario"],))


def preguntar(parent, titulo, texto, texto_si="Sí, continuar", texto_no="Cancelar",
              por_defecto_no=True):
    caja = QMessageBox(parent)
    caja.setWindowTitle(titulo)
    caja.setIcon(QMessageBox.Icon.Question)
    caja.setText(texto)
    boton_si = caja.addButton(texto_si, QMessageBox.ButtonRole.AcceptRole)
    boton_no = caja.addButton(texto_no, QMessageBox.ButtonRole.RejectRole)
    caja.setDefaultButton(boton_no if por_defecto_no else boton_si)
    caja.exec()
    return caja.clickedButton() is boton_si


def confirmar(parent, titulo, resumen):
    return preguntar(parent, titulo, resumen + "\n\n¿Desea continuar?")


def aviso(parent, texto):
    caja = QMessageBox(parent)
    caja.setWindowTitle("Atención")
    caja.setIcon(QMessageBox.Icon.Warning)
    caja.setText(texto)
    caja.addButton("Entendido", QMessageBox.ButtonRole.AcceptRole)
    caja.exec()


def listo(parent, texto):
    caja = QMessageBox(parent)
    caja.setWindowTitle("Operación completada")
    caja.setIcon(QMessageBox.Icon.Information)
    caja.setText(texto)
    caja.addButton("Aceptar", QMessageBox.ButtonRole.AcceptRole)
    caja.exec()


SUGERENCIA_DB = (
    "El programa no pudo guardar los cambios. Esto puede ocurrir por:\n"
    "1. La base de datos está abierta en otra ventana del programa.\n"
    "2. La carpeta del programa no permite guardar cambios.\n"
    "3. Algún dato no cumple una regla interna de la base de datos.\n\n"
    "Qué puede hacer:\n"
    "1. Cierre otras ventanas del programa y vuelva a intentar.\n"
    "2. Revise los datos del formulario y corrija lo que esté marcado.\n"
    "3. Si el problema continúa, avise al administrador.\n\n"
    "El detalle técnico quedó guardado en el registro de errores. "
    "La ruta exacta aparece en Ayuda > Diagnóstico."
)
SUGERENCIA_ARCHIVO = (
    "No se pudo escribir el archivo de Excel. Esto puede ocurrir por:\n"
    "1. El archivo está abierto en Excel u otro programa.\n"
    "2. La carpeta no permite guardar cambios.\n\n"
    "Qué puede hacer:\n"
    "1. Cierre el archivo en Excel y vuelva a intentar.\n"
    "2. Guarde el programa en una carpeta donde tenga permiso de escritura.\n"
    "3. Si el problema continúa, avise al administrador."
)
SUGERENCIA_INESPERADO = (
    "Ocurrió una situación no prevista al procesar la operación.\n\n"
    "Qué puede hacer:\n"
    "1. Vuelva a intentar la operación.\n"
    "2. Cierre y abra de nuevo el programa.\n"
    "3. Si el problema continúa, avise al administrador."
)


COLUMNAS_LEGIBLES = {
    "nombre": "nombre", "telefono": "teléfono", "correo": "correo",
    "direccion": "dirección", "rfc": "RFC", "nombreusuario": "usuario",
    "clavehash": "contraseña", "rol": "puesto", "idcliente": "cliente",
    "idempleado": "empleado", "idservicio": "servicio",
}


def describir_error_bd(error):
    """Traduce los errores más comunes de la base de datos a lenguaje del usuario."""
    texto = str(error)
    bajo = texto.lower()
    if "readonly" in bajo or "read-only" in bajo:
        return ("La base de datos es de solo lectura. Mueva el programa y la base de datos "
                "a una carpeta donde tenga permiso de escritura, por ejemplo Documentos.")
    if "unable to open database file" in bajo:
        return ("No se encontró el archivo de la base de datos o la carpeta no permite abrirlo. "
                "Verifique que el archivo esté junto al programa.")
    if "disk is full" in bajo:
        return "No hay espacio suficiente en el disco. Libere espacio y vuelva a intentar."
    if "i/o error" in bajo:
        return ("Se produjo un error de lectura o escritura en el disco. Revise la unidad "
                "y vuelva a intentar.")
    if "no such table" in bajo:
        return ("La base de datos no tiene la estructura esperada. Puede ser de otra versión "
                "del programa; avise al administrador.")
    if "no such column" in bajo:
        return ("La base de datos no coincide con esta versión del programa. "
                "Avise al administrador.")
    columna = texto.rsplit(":", 1)[-1].strip().split(".")[-1]
    nombre = COLUMNAS_LEGIBLES.get(columna.lower(), columna)
    if "not null constraint failed" in bajo:
        return f"Falta un dato obligatorio: {nombre}. Complételo y vuelva a intentar."
    if "unique constraint failed" in bajo:
        return f"Ya existe un registro con ese {nombre}. Use otro valor."
    if "check constraint failed" in bajo:
        return ("Algún dato no cumple el formato permitido por la base de datos. "
                "Revise los campos y vuelva a intentar.")
    if "foreign key constraint failed" in bajo:
        return ("El registro está relacionado con otros datos y no se puede modificar así. "
                "Avise al administrador.")
    if "datatype mismatch" in bajo:
        return "Un dato tiene un tipo incorrecto. Revise los campos y vuelva a intentar."
    return None


def carpeta_programa():
    return os.path.dirname(os.path.abspath(__file__))


def carpeta_escribible(ruta):
    if not ruta:
        return False
    try:
        with tempfile.NamedTemporaryFile(dir=ruta, prefix=".tg_prueba_", delete=True):
            return True
    except OSError:
        return False


def ruta_log():
    carpeta = carpeta_programa()
    if carpeta_escribible(carpeta):
        return os.path.join(carpeta, "errores.log")
    return os.path.join(tempfile.gettempdir(), "TheGonzalez_errores.log")


def configurar_ruta_base():
    """Asegura que la base se abra junto al programa y no en la carpeta de donde se lanzó."""
    nombre = "serviciosTheGonzalez.db"
    carpeta = carpeta_programa()
    candidatos = [
        os.path.join(carpeta, nombre),
        os.path.join(os.path.abspath(os.getcwd()), nombre),
    ]
    for candidato in candidatos:
        if os.path.exists(candidato):
            nucleo.DB = candidato
            nucleo.RESPALDOS = os.path.join(carpeta, "respaldos")
            return candidato
    nucleo.DB = candidatos[0]
    nucleo.RESPALDOS = os.path.join(carpeta, "respaldos")
    return nucleo.DB


def asegurar_esquema(con):
    """Agrega columnas y tablas nuevas sin destruir la información existente."""
    try:
        if con.execute("SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'usuario';").fetchone() is None:
            return
        columnas = {fila[1] for fila in con.execute("PRAGMA table_info(usuario);").fetchall()}
        for columna, tipo in (("correo", "TEXT"), ("telefono", "TEXT"),
                              ("foto", "TEXT"), ("conectado", "INTEGER DEFAULT 0"),
                              ("ultimoAcceso", "TEXT")):
            if columna not in columnas:
                con.execute(f"ALTER TABLE usuario ADD COLUMN {columna} {tipo};")
        con.execute("""CREATE TABLE IF NOT EXISTS movimiento (
            idMovimiento INTEGER PRIMARY KEY AUTOINCREMENT,
            idUsuario INTEGER, usuario TEXT, rol TEXT, accion TEXT,
            tabla TEXT, idRegistro INTEGER, detalle TEXT, fecha TEXT);""")
        con.execute("""CREATE TABLE IF NOT EXISTS solicitud_rol (
            idSolicitud INTEGER PRIMARY KEY AUTOINCREMENT,
            idUsuario INTEGER, usuario TEXT, nombre TEXT,
            rolActual TEXT, rolSolicitado TEXT, motivo TEXT,
            estado TEXT DEFAULT 'PENDIENTE', fecha TEXT,
            resueltoPor TEXT, fechaResolucion TEXT);""")
        con.execute("CREATE INDEX IF NOT EXISTS idx_mov_usuario ON movimiento (idUsuario);")
        con.execute("CREATE INDEX IF NOT EXISTS idx_mov_fecha ON movimiento (fecha);")
        con.execute("CREATE INDEX IF NOT EXISTS idx_sol_estado ON solicitud_rol (estado);")
        con.execute("""CREATE TABLE IF NOT EXISTS notificacion (
            idNotificacion INTEGER PRIMARY KEY AUTOINCREMENT,
            idUsuario INTEGER, titulo TEXT, mensaje TEXT,
            leida INTEGER DEFAULT 0, fecha TEXT, origen TEXT);""")
        con.execute("CREATE INDEX IF NOT EXISTS idx_noti_usuario ON notificacion (idUsuario, leida);")
        con.execute("""CREATE TABLE IF NOT EXISTS tg_detalle_cotizacion (
            idDetalle INTEGER PRIMARY KEY AUTOINCREMENT,
            idCotizacion INTEGER, descripcion TEXT, unidad TEXT,
            cantidad REAL, precio REAL, importe REAL);""")
        con.execute("CREATE INDEX IF NOT EXISTS idx_detalle_cot ON tg_detalle_cotizacion (idCotizacion);")
        # La unidad de medida debe vivir tambien en la tabla historica de detalles.
        if con.execute("SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'detallecotizacion';").fetchone():
            columnas_detalle = {fila[1] for fila in con.execute("PRAGMA table_info(detallecotizacion);").fetchall()}
            if "unidad" not in columnas_detalle:
                con.execute("ALTER TABLE detallecotizacion ADD COLUMN unidad TEXT;")
        # Columnas de apoyo para la cotizacion formal (vigencia, condiciones y estado del documento).
        if con.execute("SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'cotizacion';").fetchone():
            columnas_cot = {fila[1] for fila in con.execute("PRAGMA table_info(cotizacion);").fetchall()}
            for columna, tipo in (("vigenciaDias", "INTEGER NOT NULL DEFAULT 15"),
                                  ("formaPago", "TEXT"),
                                  ("tiempoEntrega", "TEXT"),
                                  ("garantia", "TEXT"),
                                  ("estadoDocumento", "TEXT")):
                if columna not in columnas_cot:
                    con.execute(f"ALTER TABLE cotizacion ADD COLUMN {columna} {tipo};")
            con.execute("CREATE INDEX IF NOT EXISTS idx_cot_servicio ON cotizacion (idServicio);")
        if con.execute("SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'factura';").fetchone():
            con.execute("CREATE INDEX IF NOT EXISTS idx_fac_servicio ON factura (idServicio);")
        con.execute("INSERT OR REPLACE INTO configuracion (clave, valor) VALUES ('ESQUEMA', '4');")
        if con.execute("SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'movimiento';").fetchone():
            columnas_mov = {fila[1] for fila in con.execute("PRAGMA table_info(movimiento);").fetchall()}
            for columna, tipo in (("equipo", "TEXT"), ("importe", "REAL")):
                if columna not in columnas_mov:
                    con.execute(f"ALTER TABLE movimiento ADD COLUMN {columna} {tipo};")
            con.execute("CREATE INDEX IF NOT EXISTS idx_mov_accion ON movimiento (accion);")
        if con.execute("SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'cotizacion';").fetchone():
            try:
                con.execute("""UPDATE cotizacion SET estado = 'REGISTRADA'
                               WHERE UPPER(COALESCE(estado, '')) IN ('PENDIENTE', 'GENERADA', '')
                                  OR estado IS NULL;""")
            except sqlite3.Error:
                pass
        con.commit()
    except sqlite3.Error as e:
        registrar_error("ESQUEMA", str(e))


def registrar_movimiento(con, accion, tabla, idRegistro, detalle=None, importe=None):
    """Registra la bitácora existente y el movimiento con el usuario que lo hizo."""
    bitacora(con, accion, tabla, idRegistro, detalle)
    texto = None if detalle is None else str(detalle)
    if texto and len(texto) > 400:
        texto = texto[:397] + "..."
    try:
        con.execute(
            "INSERT INTO movimiento (idUsuario, usuario, rol, accion, tabla, idRegistro, detalle, fecha, equipo) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?);",
            (SESION.get("id"), SESION.get("nombreCompleto") or SESION.get("nombre") or "sistema",
             str(SESION.get("rol") or "").strip().upper() or "SISTEMA", accion, tabla, idRegistro,
             texto, datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
             os.environ.get("COMPUTERNAME") or os.environ.get("HOSTNAME") or ""))
        if importe is not None:
            con.execute("UPDATE movimiento SET importe = ? WHERE idMovimiento = last_insert_rowid();",
                        (float(importe),))
    except sqlite3.Error:
        pass


def marcar_conectado(valor):
    id_usuario = SESION.get("id")
    if not id_usuario:
        return
    try:
        con = conectar()
        try:
            con.execute("UPDATE usuario SET conectado = ? WHERE idUsuario = ?;",
                        (1 if valor else 0, id_usuario))
            con.execute("UPDATE usuario SET ultimoAcceso = ? WHERE idUsuario = ?;",
                        (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), id_usuario))
            con.commit()
        finally:
            con.close()
    except sqlite3.Error:
        pass
    try:
        con = conectar()
        try:
            with con:
                registrar_movimiento(con, "INICIAR SESION" if valor else "CERRAR SESION",
                                     "usuario", id_usuario,
                                     f"{SESION.get('nombre')} ({SESION.get('rol')})")
        finally:
            con.close()
    except sqlite3.Error:
        pass


def copiar_foto(origen, usuario):
    carpeta = os.path.join(carpeta_programa(), "perfiles")
    os.makedirs(carpeta, exist_ok=True)
    extension = os.path.splitext(origen)[1].lower() or ".png"
    destino = os.path.join(carpeta, f"{usuario}_{datetime.now():%Y%m%d%H%M%S}{extension}")
    shutil.copyfile(origen, destino)
    return destino


def config_correo():
    ajustes = QSettings("TheGonzalez", "Servicios")
    usuario = str(os.environ.get("TG_GMAIL_USUARIO") or ajustes.value("correo/usuario", "") or "")
    clave = str(os.environ.get("TG_GMAIL_CLAVE") or ajustes.value("correo/clave", "") or "")
    admin = str(ajustes.value("correo/admin", "") or "")
    return usuario, clave, admin


def correo_administrador(con):
    """Correo del administrador del panel; si no existe, usa el del gerente."""
    try:
        fila = con.execute(
            """SELECT correo FROM administrador
               WHERE activo = 1 AND correo IS NOT NULL AND TRIM(correo) <> ''
               ORDER BY idAdmin LIMIT 1;""").fetchone()
        if fila and fila[0]:
            return str(fila[0]).strip()
    except sqlite3.Error:
        pass
    fila = con.execute(
        """SELECT correo FROM usuario
           WHERE UPPER(TRIM(rol)) = 'GERENTE' AND activo = 1
             AND correo IS NOT NULL AND TRIM(correo) <> ''
           ORDER BY idUsuario LIMIT 1;""").fetchone()
    if fila and fila[0]:
        return str(fila[0]).strip()
    return config_correo()[2]


def enviar_correo(destino, asunto, cuerpo):
    usuario, clave, _admin = config_correo()
    if not usuario or not clave:
        raise RuntimeError("No hay una cuenta de Gmail configurada. Use “Configurar Gmail”.")
    mensaje = EmailMessage()
    mensaje["From"] = usuario
    mensaje["To"] = destino
    mensaje["Subject"] = asunto
    mensaje.set_content(cuerpo)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=15) as servidor:
        servidor.login(usuario, clave)
        servidor.send_message(mensaje)


def registrar_error(titulo, detalle):
    """Guarda el detalle técnico en el registro para el administrador."""
    try:
        with open(ruta_log(), "a", encoding="utf-8") as archivo:
            archivo.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {titulo}\n{detalle}\n\n")
    except Exception:
        pass


def ruta_base_datos():
    try:
        con = conectar()
        try:
            filas = con.execute("PRAGMA database_list;").fetchall()
        finally:
            con.close()
    except Exception:
        return ""
    for fila in filas:
        try:
            if fila[1] == "main":
                return fila[2] or ""
        except (IndexError, TypeError):
            continue
    return ""


def base_escribible(ruta):
    if not ruta or not os.path.exists(ruta):
        return False
    try:
        con = conectar()
        try:
            version = con.execute("PRAGMA user_version;").fetchone()[0]
            con.execute(f"PRAGMA user_version = {int(version)};")
            return True
        finally:
            con.close()
    except Exception:
        return False


def revisar_integridad():
    """Comprueba la estructura y las referencias más importantes de la base de datos."""
    problemas = 0
    try:
        con = conectar()
        try:
            chequeo = con.execute("PRAGMA quick_check;").fetchone()
            estado = str(chequeo[0]) if chequeo else "sin respuesta"
            consultas = [
                ("Servicios sin cliente",
                 "SELECT COUNT(*) FROM servicio s LEFT JOIN cliente c ON c.idCliente = s.idCliente "
                 "WHERE c.idCliente IS NULL;"),
                ("Asignaciones sin empleado",
                 "SELECT COUNT(*) FROM servicioempleado se LEFT JOIN empleado e ON e.idEmpleado = se.idEmpleado "
                 "WHERE e.idEmpleado IS NULL;"),
                ("Cotizaciones sin servicio",
                 "SELECT COUNT(*) FROM cotizacion co LEFT JOIN servicio s ON s.idServicio = co.idServicio "
                 "WHERE s.idServicio IS NULL;"),
                ("Facturas sin servicio",
                 "SELECT COUNT(*) FROM factura f LEFT JOIN servicio s ON s.idServicio = f.idServicio "
                 "WHERE s.idServicio IS NULL;"),
                ("Detalles de cotización sin cotización",
                 "SELECT COUNT(*) FROM tg_detalle_cotizacion d "
                 "LEFT JOIN cotizacion c ON c.idCotizacion = d.idCotizacion "
                 "WHERE c.idCotizacion IS NULL;"),
                ("Servicios en cotización sin cotización registrada",
                 "SELECT COUNT(*) FROM servicio s "
                 "JOIN estado_servicio es ON es.idEstado = s.idEstado "
                 "WHERE es.nombreEstado = 'COTIZACION_GENERADA' AND NOT EXISTS ("
                 "  SELECT 1 FROM cotizacion c WHERE c.idServicio = s.idServicio);"),
                ("Facturas sin cotización aceptada",
                 "SELECT COUNT(*) FROM factura f WHERE NOT EXISTS ("
                 "  SELECT 1 FROM cotizacion c WHERE c.idServicio = f.idServicio "
                 "  AND UPPER(COALESCE(c.estado, '')) = 'ACEPTADA');"),
                ("Empleados activos sin puesto",
                 "SELECT COUNT(*) FROM empleado e WHERE e.activo = 1 AND NOT EXISTS ("
                 "  SELECT 1 FROM puesto p WHERE p.idEmpleado = e.idEmpleado);"),
                ("Cotizaciones con importe inválido",
                 "SELECT COUNT(*) FROM cotizacion WHERE total IS NULL OR total <= 0;"),
                ("Facturas con importe inválido",
                 "SELECT COUNT(*) FROM factura WHERE total IS NULL OR total <= 0;"),
                ("Servicios facturados sin factura",
                 "SELECT COUNT(*) FROM servicio s "
                 "JOIN estado_servicio es ON es.idEstado = s.idEstado "
                 "WHERE es.nombreEstado = 'FACTURADO' AND NOT EXISTS ("
                 "  SELECT 1 FROM factura f WHERE f.idServicio = s.idServicio);"),
                ("Servicios con más de una factura",
                 "SELECT COUNT(*) FROM (SELECT idServicio FROM factura "
                 "GROUP BY idServicio HAVING COUNT(*) > 1);"),
                ("Detalles sin unidad de medida",
                 "SELECT COUNT(*) FROM tg_detalle_cotizacion "
                 "WHERE unidad IS NULL OR TRIM(unidad) = '';"),
                ("Cotizaciones con total que no cuadra",
                 "SELECT COUNT(*) FROM cotizacion "
                 "WHERE ABS(ROUND(COALESCE(subtotal, 0) + COALESCE(iva, 0), 2) - ROUND(COALESCE(total, 0), 2)) > 0.01;"),
                ("Usuarios con puesto desconocido",
                 "SELECT COUNT(*) FROM usuario WHERE UPPER(TRIM(rol)) NOT IN "
                 "('GERENTE', 'COTIZADOR', 'SUPERVISOR', 'CONTADOR');"),
                ("Movimientos sin usuario registrado",
                 "SELECT COUNT(*) FROM movimiento WHERE idUsuario IS NULL OR idUsuario = 0;"),
                ("Facturas con cliente distinto al del servicio",
                 "SELECT COUNT(*) FROM factura f JOIN servicio s ON s.idServicio = f.idServicio "
                 "WHERE f.idCliente <> s.idCliente;"),
            ]
            filas = []
            for etiqueta, sql in consultas:
                try:
                    total = con.execute(sql).fetchone()[0]
                except sqlite3.Error:
                    total = None
                if total is None:
                    problemas += 1
                elif int(total) > 0:
                    problemas += int(total)
                valor = "No se pudo comprobar" if total is None else str(total)
                color = "#5B6B80" if total in (0, None) else "#B42318"
                filas.append(f'<tr><td style="color:#5B6B80;">{etiqueta}</td>'
                             f'<td style="color:{color};"><b>{valor}</b></td></tr>')
        finally:
            con.close()
    except sqlite3.Error as e:
        return ('<table width="100%" cellspacing="0" cellpadding="10" style="background-color:#FEF3F2;">'
                f'<tr><td style="color:#B42318;">No se pudo revisar la integridad: {escape(str(e))}</td></tr></table>')

    correcta = estado.lower() == "ok" and problemas == 0
    if correcta:
        resumen = ('<table width="100%" cellspacing="0" cellpadding="10" style="background-color:#ECFDF3;">'
                   '<tr><td style="color:#1B7F4B;"><b>Integridad correcta.</b> No hay registros '
                   'huérfanos ni errores internos.</td></tr></table>')
    else:
        resumen = ('<table width="100%" cellspacing="0" cellpadding="10" style="background-color:#FEF3F2;">'
                   '<tr><td style="color:#B42318;"><b>Revise la integridad.</b> Hay registros sin '
                   'su referencia o un problema interno. Guarde una copia y avise al administrador.</td></tr></table>')
    detalle = ""
    if estado.lower() != "ok":
        detalle = f'<p style="color:#B42318;">Resultado interno: {escape(estado)}</p>'
    return (f'<h3 style="color:#111C2E; font-size:12pt;">Integridad de la base de datos</h3>'
            f'<table width="100%" cellspacing="0" cellpadding="7" style="border:1px solid #E3E8F0;">'
            f'{"".join(filas)}</table>{resumen}{detalle}')


def diagnostico_html():
    archivo = os.path.abspath(__file__)
    carpeta = carpeta_programa()
    ruta_bd = ruta_base_datos()
    log = ruta_log()
    pruebas = [
        ("Programa", f"Versión {VERSION}"),
        ("Archivo en ejecución", archivo),
        ("Carpeta del programa", carpeta),
        ("Base de datos", ruta_bd or "No se pudo determinar"),
        ("Registro de errores", log),
        ("Python", sys.version.split()[0]),
    ]
    ok_carpeta = carpeta_escribible(carpeta)
    ok_bd = base_escribible(ruta_bd)
    ok_log = carpeta_escribible(os.path.dirname(log))

    def marca(valor):
        if valor:
            return '<span style="color:#1B7F4B;"><b>Sí</b></span>'
        return '<span style="color:#B42318;"><b>No</b></span>'

    filas = []
    for etiqueta, valor in pruebas:
        extra = ""
        if etiqueta == "Carpeta del programa":
            extra = f"<br>Escritura: {marca(ok_carpeta)}"
        elif etiqueta == "Base de datos":
            extra = f"<br>Escritura: {marca(ok_bd)}"
        elif etiqueta == "Registro de errores":
            extra = f"<br>Escritura: {marca(ok_log)}"
        filas.append(
            f'<tr><td width="30%" style="color:#5B6B80;"><b>{etiqueta}</b></td>'
            f'<td style="color:#1B2432;">{escape(str(valor)).replace(chr(10), "<br>")}{extra}</td></tr>')

    consejos = []
    if not ok_bd:
        consejos.append("La base de datos no permite escribir. Cierre otras ventanas del programa, "
                        "quite la marca de solo lectura del archivo y evite carpetas sincronizadas "
                        "en la nube o unidades de solo lectura.")
    if not ok_carpeta:
        consejos.append("La carpeta del programa no permite escribir. Mueva el programa y la base "
                        "de datos a Documentos o al Escritorio.")
    if not ok_log:
        consejos.append("No se puede escribir el registro de errores. El detalle se guardará en la "
                        "carpeta temporal del sistema.")
    aviso = ""
    if consejos:
        aviso = ('<table width="100%" cellspacing="0" cellpadding="10" style="background-color:#FEF3F2;">'
                 '<tr><td style="color:#B42318;">' + "<br>".join(consejos) + "</td></tr></table>")
    else:
        aviso = ('<table width="100%" cellspacing="0" cellpadding="10" style="background-color:#ECFDF3;">'
                 '<tr><td style="color:#1B7F4B;"><b>Todo listo.</b> El programa puede leer y escribir '
                 'la base de datos correctamente.</td></tr></table>')

    return f"""
<p style="color:#475569; font-size:11pt;">Compruebe que está ejecutando el archivo actualizado y que
la carpeta permite guardar cambios.</p>
<table width="100%" cellspacing="0" cellpadding="7" style="border:1px solid #E3E8F0;">
{''.join(filas)}
</table>
{aviso}
{revisar_integridad()}
<table width="100%" cellspacing="0" cellpadding="10" style="background-color:#EEF4FF;">
<tr><td><b>Importante:</b> si la ruta de <b>Archivo en ejecución</b> no es el archivo que actualizó,
el programa está abriendo otra copia. Cierre el programa, abra el archivo correcto y vuelva a intentar.</td></tr>
</table>
"""


def fallo(parent, texto, sugerencia=None, tecnico=None):
    if tecnico:
        registrar_error(texto, tecnico)
    caja = QMessageBox(parent)
    caja.setWindowTitle("No se pudo completar la operación")
    caja.setIcon(QMessageBox.Icon.Critical)
    caja.setText(texto)
    aceptar = caja.addButton("Aceptar", QMessageBox.ButtonRole.AcceptRole)
    ver_causa = None
    if sugerencia:
        ver_causa = caja.addButton("Posible causa", QMessageBox.ButtonRole.ActionRole)
    caja.setDefaultButton(aceptar)
    caja.exec()
    if ver_causa is not None and caja.clickedButton() is ver_causa:
        detalle_soporte = None
        if tecnico:
            detalle_soporte = (
                f"Versión: {VERSION}\n"
                f"Archivo: {os.path.abspath(__file__)}\n"
                f"Base de datos: {ruta_base_datos() or 'no determinada'}\n"
                f"Fecha: {datetime.now():%Y-%m-%d %H:%M:%S}\n"
                f"Mensaje: {texto}\n"
                f"Detalle técnico: {tecnico}"
            )
        DialogoMensaje(parent, "Posible causa", sugerencia, detalle_soporte).exec()


def plural(cantidad, singular, plural_):
    return f"{cantidad} {singular if cantidad == 1 else plural_}"


def estado_de(con, idServicio):
    fila = con.execute(
        """SELECT es.nombreEstado FROM servicio s
           JOIN estado_servicio es ON es.idEstado = s.idEstado
           WHERE s.idServicio = ?;""", (idServicio,)).fetchone()
    if fila is None:
        raise ReglaError("Este servicio ya no existe. Actualice la lista.")
    return fila[0]


def id_estado(con, nombre):
    fila = con.execute("SELECT idEstado FROM estado_servicio WHERE nombreEstado = ?;", (nombre,)).fetchone()
    if fila is None:
        raise ReglaError("La configuración de estados del servicio está incompleta. "
                         "Avise al administrador.")
    return fila[0]


REFERENCIAS_POR_TABLA = {
    "cliente": ("servicio", "factura"),
    "empleado": ("servicioempleado", "puesto", "revisionservicio", "retroalimentacion"),
    "servicio": ("servicioempleado", "cotizacion", "factura", "revisionservicio", "retroalimentacion"),
    "usuario": (),
}


def reiniciar_secuencia(con, tabla):
    """Reinicia el contador solo si no hay registros ni referencias que puedan cruzarse."""
    try:
        if con.execute(f"SELECT COUNT(*) FROM {tabla};").fetchone()[0]:
            return
        for referencia in REFERENCIAS_POR_TABLA.get(tabla, ()):
            if con.execute(f"SELECT COUNT(*) FROM {referencia};").fetchone()[0]:
                return
        con.execute("DELETE FROM sqlite_sequence WHERE name = ?;", (tabla,))
    except sqlite3.Error:
        pass


def buscar_clientes(con, texto):
    """Busca clientes por nombre, número de cliente, teléfono o RFC (sin distinguir mayúsculas)."""
    texto = (texto or "").strip()
    if not texto:
        return con.execute("SELECT idCliente, nombre FROM cliente WHERE activo = 1 ORDER BY nombre;").fetchall()
    if texto.isdigit():
        fila = con.execute("SELECT idCliente, nombre FROM cliente WHERE idCliente = ? AND activo = 1;",
                           (int(texto),)).fetchall()
        if fila:
            return fila
    exactos = con.execute(
        "SELECT idCliente, nombre FROM cliente WHERE activo = 1 AND lower(trim(nombre)) = lower(?);",
        (texto,)).fetchall()
    if exactos:
        return exactos
    patron = "%" + texto.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_") + "%"
    return con.execute(
        """SELECT idCliente, nombre FROM cliente
           WHERE activo = 1 AND (nombre LIKE ? ESCAPE '\\' OR COALESCE(telefono, '') LIKE ? ESCAPE '\\'
                                 OR COALESCE(rfc, '') LIKE ? ESCAPE '\\')
           ORDER BY nombre;""", (patron, patron, patron.upper())).fetchall()


def registrar_cliente_db(con, d):
    with con:
        reiniciar_secuencia(con, "cliente")
        cur = con.execute("INSERT INTO cliente (nombre, telefono, correo, direccion, rfc) VALUES (?, ?, ?, ?, ?);",
                          (d["nombre"], d["telefono"], d["correo"], d["direccion"], d["rfc"]))
        registrar_movimiento(con, "REGISTRAR CLIENTE", "cliente", cur.lastrowid, d["nombre"])
    return cur.lastrowid


def modificar_cliente_db(con, idCliente, d):
    with con:
        cur = con.execute(
            "UPDATE cliente SET nombre=?, telefono=?, correo=?, direccion=?, rfc=? WHERE idCliente=? AND activo = 1;",
            (d["nombre"], d["telefono"], d["correo"], d["direccion"], d["rfc"], idCliente))
        if cur.rowcount == 0:
            raise ReglaError(MSG_REGISTRO)
        registrar_movimiento(con, "MODIFICAR CLIENTE", "cliente", idCliente, d["nombre"])


def eliminar_cliente_db(con, idCliente):
    abiertos = con.execute(
        """SELECT COUNT(*) FROM servicio s
           JOIN estado_servicio es ON es.idEstado = s.idEstado
           WHERE s.idCliente = ? AND es.nombreEstado NOT IN ('FINALIZADO', 'CANCELADO');""",
        (idCliente,)).fetchone()[0]
    if abiertos:
        raise ReglaError(f"No se puede dar de baja: el cliente tiene "
                         f"{plural(abiertos, 'servicio', 'servicios')} en proceso. "
                         "Finalice o cancele esos servicios primero.")
    with con:
        cur = con.execute("UPDATE cliente SET activo = 0 WHERE idCliente = ? AND activo = 1;", (idCliente,))
        if cur.rowcount == 0:
            raise ReglaError(MSG_REGISTRO)
        registrar_movimiento(con, "ELIMINAR CLIENTE", "cliente", idCliente)


def registrar_empleado_db(con, d):
    with con:
        reiniciar_secuencia(con, "empleado")
        cur = con.execute("INSERT INTO empleado (nombre, aPaterno, aMaterno, telefono, correo) VALUES (?, ?, ?, ?, ?);",
                          (d["nombre"], d["ap"], d["am"], d["telefono"], d["correo"]))
        idEmp = cur.lastrowid
        for idRol in d["roles"]:
            con.execute("INSERT INTO puesto (idEmpleado, idRol) VALUES (?, ?);", (idEmp, idRol))
        registrar_movimiento(con, "REGISTRAR EMPLEADO", "empleado", idEmp, f"{d['nombre']} {d['ap']}")
    return idEmp


def modificar_empleado_db(con, idEmpleado, d):
    with con:
        cur = con.execute(
            "UPDATE empleado SET nombre=?, aPaterno=?, aMaterno=?, telefono=?, correo=? WHERE idEmpleado=? AND activo = 1;",
            (d["nombre"], d["ap"], d["am"], d["telefono"], d["correo"], idEmpleado))
        if cur.rowcount == 0:
            raise ReglaError(MSG_REGISTRO)
        con.execute("DELETE FROM puesto WHERE idEmpleado = ?;", (idEmpleado,))
        for idRol in d["roles"]:
            con.execute("INSERT INTO puesto (idEmpleado, idRol) VALUES (?, ?);", (idEmpleado, idRol))
        registrar_movimiento(con, "MODIFICAR EMPLEADO", "empleado", idEmpleado, f"{d['nombre']} {d['ap']}")


def eliminar_empleado_db(con, idEmpleado):
    abiertos = con.execute(
        """SELECT COUNT(*) FROM servicioempleado se
           JOIN servicio s ON s.idServicio = se.idServicio
           JOIN estado_servicio es ON es.idEstado = s.idEstado
           WHERE se.idEmpleado = ? AND es.nombreEstado NOT IN ('FINALIZADO', 'CANCELADO');""",
        (idEmpleado,)).fetchone()[0]
    if abiertos:
        raise ReglaError(f"No se puede dar de baja: el empleado está asignado a "
                         f"{plural(abiertos, 'servicio', 'servicios')} en proceso.")
    with con:
        cur = con.execute("UPDATE empleado SET activo = 0 WHERE idEmpleado = ? AND activo = 1;", (idEmpleado,))
        if cur.rowcount == 0:
            raise ReglaError(MSG_REGISTRO)
        registrar_movimiento(con, "ELIMINAR EMPLEADO", "empleado", idEmpleado)


def registrar_servicio_db(con, idCliente, problema, descripcion, observaciones):
    with con:
        reiniciar_secuencia(con, "servicio")
        cur = con.execute(
            """INSERT INTO servicio (idCliente, idEstado, descripcionProblema, descripcionServicio,
                                     observaciones, fechaRegistro)
               VALUES (?, ?, ?, ?, ?, ?);""",
            (idCliente, id_estado(con, "REGISTRADO"), problema, descripcion, observaciones,
             date.today().isoformat()))
        registrar_movimiento(con, "REGISTRAR SERVICIO", "servicio", cur.lastrowid, problema)
    return cur.lastrowid


def modificar_servicio_db(con, idServicio, problema, descripcion, observaciones):
    estado = estado_de(con, idServicio)
    if estado not in ESTADOS_EDITABLES_SERVICIO:
        raise ReglaError(f"Este servicio ya no se puede modificar porque está {estado.lower().replace('_', ' ')}.")
    with con:
        con.execute("UPDATE servicio SET descripcionProblema=?, descripcionServicio=?, observaciones=? WHERE idServicio=?;",
                    (problema, descripcion, observaciones, idServicio))
        registrar_movimiento(con, "MODIFICAR SERVICIO", "servicio", idServicio, problema)


def cancelar_servicio_db(con, idServicio, motivo):
    estado = estado_de(con, idServicio)
    if estado not in ESTADOS_CANCELABLES:
        raise ReglaError(f"No se puede cancelar un servicio que está {estado.lower().replace('_', ' ')}.")
    with con:
        con.execute("UPDATE servicio SET idEstado=?, motivoCancelacion=? WHERE idServicio=?;",
                    (id_estado(con, "CANCELADO"), motivo, idServicio))
        registrar_movimiento(con, "CANCELAR SERVICIO", "servicio", idServicio, motivo)


def finalizar_servicio_db(con, idServicio):
    estado = estado_de(con, idServicio)
    if estado != "FACTURADO":
        raise ReglaError("Solo se puede finalizar un servicio que ya fue facturado. "
                         f"Estado actual: {estado.lower().replace('_', ' ')}.")
    with con:
        con.execute("UPDATE servicio SET idEstado=? WHERE idServicio=?;",
                    (id_estado(con, "FINALIZADO"), idServicio))
        registrar_movimiento(con, "ESTADO FINALIZADO", "servicio", idServicio)


def participantes_servicio(con, idServicio):
    return con.execute(
        """SELECT e.idEmpleado, e.nombre || ' ' || e.aPaterno, se.tipoParticipacion
           FROM servicioempleado se JOIN empleado e ON e.idEmpleado = se.idEmpleado
           WHERE se.idServicio = ? ORDER BY e.nombre;""", (idServicio,)).fetchall()


def retroalimentar_db(con, idServicio, idEmpleado, texto):
    estado = estado_de(con, idServicio)
    if estado not in ESTADOS_RETROALIMENTACION:
        raise ReglaError("Todavía no se puede dar retroalimentación en este servicio.")
    if not any(p[0] == idEmpleado for p in participantes_servicio(con, idServicio)):
        raise ReglaError("Ese trabajador no está asignado a este servicio.")
    with con:
        con.execute("INSERT INTO retroalimentacion (idServicio, idEmpleado, comentarios) VALUES (?, ?, ?);",
                    (idServicio, idEmpleado, texto))
        registrar_movimiento(con, "RETROALIMENTACION", "servicio", idServicio, f"empleado {idEmpleado}")


def registrar_usuario_db(con, d):
    with con:
        reiniciar_secuencia(con, "usuario")
        crear_usuario(con, d["usuario"], d["nombre"], d["rol"], d["clave"])
        fila = con.execute("SELECT idUsuario FROM usuario WHERE nombreUsuario = ?;",
                           (d["usuario"],)).fetchone()
        idNuevo = fila[0] if fila else None
        registrar_movimiento(con, "CREAR USUARIO", "usuario", idNuevo, f"{d['usuario']} ({d['rol']})")
    return idNuevo


def eliminar_usuario_db(con, idUsuario):
    fila = con.execute("SELECT nombreUsuario, rol, activo FROM usuario WHERE idUsuario = ?;",
                       (idUsuario,)).fetchone()
    if fila is None:
        raise ReglaError(MSG_REGISTRO)
    if fila["activo"]:
        raise ReglaError("Primero desactive la cuenta y después elimínela definitivamente.")
    with con:
        con.execute("DELETE FROM usuario WHERE idUsuario = ?;", (idUsuario,))
        reiniciar_secuencia(con, "usuario")
        registrar_movimiento(con, "ELIMINAR USUARIO", "usuario", idUsuario, fila["nombreUsuario"])


def cambiar_estado_usuario_db(con, idUsuario, activo):
    fila = con.execute("SELECT nombreUsuario, rol FROM usuario WHERE idUsuario = ?;",
                       (idUsuario,)).fetchone()
    if fila is None:
        raise ReglaError(MSG_REGISTRO)
    if not activo and str(fila["rol"] or "").strip().upper() == "GERENTE":
        activos = con.execute(
            "SELECT COUNT(*) FROM usuario WHERE UPPER(TRIM(rol)) = 'GERENTE' AND activo = 1;").fetchone()[0]
        if activos <= 1:
            raise ReglaError("Debe quedar al menos un gerente activo en el sistema.")
    with con:
        con.execute("UPDATE usuario SET activo = ? WHERE idUsuario = ?;", (1 if activo else 0, idUsuario))
        registrar_movimiento(con, "ACTIVAR USUARIO" if activo else "DESACTIVAR USUARIO",
                 "usuario", idUsuario, fila["nombreUsuario"])
    return fila["nombreUsuario"]


def cambiar_clave_db(con, idUsuario, clave):
    fila = con.execute("SELECT nombreUsuario FROM usuario WHERE idUsuario = ?;",
                       (idUsuario,)).fetchone()
    if fila is None:
        raise ReglaError(MSG_REGISTRO)
    cambiar_clave_usuario(con, idUsuario, clave)
    with con:
        registrar_movimiento(con, "CAMBIAR CLAVE", "usuario", idUsuario, fila["nombreUsuario"])
    return fila["nombreUsuario"]


PERMISO_GERENTE = ("GERENTE",)
PERMISO_VENTA = ("GERENTE", "COTIZADOR")
PERMISO_SUPERVISION = ("GERENTE", "SUPERVISOR")
PERMISO_FACTURACION = ("GERENTE", "CONTADOR")


def puede(permiso):
    rol = str(SESION.get("rol") or "").strip().upper()
    return rol in {str(p).strip().upper() for p in permiso}


def requiere(permiso):
    def decorador(metodo):
        @functools.wraps(metodo)
        def envoltura(self, *args, **kwargs):
            if not puede(permiso):
                actual = str(SESION.get("rol") or "").strip().upper()
                actual = ROLES_TEXTO.get(actual, actual.capitalize() or "sin puesto")
                requeridos = ", ".join(ROLES_TEXTO.get(str(p).strip().upper(), str(p).capitalize())
                                       for p in permiso)
                aviso(self, f"Su cuenta tiene el puesto {actual}.\n\n"
                            f"Esta acción requiere: {requeridos}.\n\n"
                            "Pida al gerente que ajuste su puesto en la sección Usuarios.")
                return None
            return metodo(self, *_sin_marca_de_clic(args), **kwargs)
        return envoltura
    return decorador


def abrir_archivo(ruta):
    QDesktopServices.openUrl(QUrl.fromLocalFile(os.path.abspath(ruta)))


def nombre_servicio_db(con, idServicio):
    fila = con.execute("SELECT c.nombre FROM servicio s JOIN cliente c ON c.idCliente = s.idCliente "
                       "WHERE s.idServicio = ?;", (idServicio,)).fetchone()
    return fila[0] if fila else "-"


def asignar_personal_db(con, idServicio, asignaciones):
    estado = estado_de(con, idServicio)
    if estado not in ("REGISTRADO", "ASIGNADO", "COTIZACION_GENERADA", "ACEPTADO",
                      "EN_PROCESO", "CORRECCION_SOLICITADA"):
        raise ReglaError(f"No se puede asignar personal a un servicio que está "
                         f"{estado.lower().replace('_', ' ')}. "
                         "El personal se asigna antes de enviar el trabajo a revisión.")
    if not asignaciones:
        raise ReglaError("Seleccione al menos una persona para asignar.")
    agregados = 0
    with con:
        for idEmpleado, participacion in asignaciones:
            participacion = (participacion or "").strip() or "Apoyo"
            cur = con.execute(
                "INSERT OR IGNORE INTO servicioempleado (idServicio, idEmpleado, tipoParticipacion) VALUES (?, ?, ?);",
                (idServicio, idEmpleado, participacion))
            agregados += cur.rowcount
        if agregados and estado == "REGISTRADO":
            con.execute("UPDATE servicio SET idEstado = ? WHERE idServicio = ?;",
                        (id_estado(con, "ASIGNADO"), idServicio))
        registrar_movimiento(con, "ASIGNAR PERSONAL", "servicio", idServicio,
                             f"{plural(agregados, 'persona nueva', 'personas nuevas')}")
    return agregados


_ESPECIALES = [
    "CERO", "UN", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE",
    "DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISÉIS",
    "DIECISIETE", "DIECIOCHO", "DIECINUEVE", "VEINTE", "VEINTIUNO",
    "VEINTIDÓS", "VEINTITRÉS", "VEINTICUATRO", "VEINTICINCO", "VEINTISÉIS",
    "VEINTISIETE", "VEINTIOCHO", "VEINTINUEVE",
]
_DECENAS = ["", "", "VEINTE", "TREINTA", "CUARENTA", "CINCUENTA",
            "SESENTA", "SETENTA", "OCHENTA", "NOVENTA"]
_CENTENAS = ["", "CIENTO", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS",
             "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS"]


def _letras_centenas(numero):
    if numero < 30:
        return _ESPECIALES[numero]
    if numero < 100:
        decena, unidad = divmod(numero, 10)
        return _DECENAS[decena] + ((" Y " + _ESPECIALES[unidad]) if unidad else "")
    if numero == 100:
        return "CIEN"
    centena, resto = divmod(numero, 100)
    return _CENTENAS[centena] + ((" " + _letras_centenas(resto)) if resto else "")


def numero_a_letras(numero):
    numero = int(abs(numero))
    if numero == 0:
        return "CERO"
    millones, resto = divmod(numero, 1000000)
    miles, unidades = divmod(resto, 1000)
    partes = []
    if millones:
        partes.append("UN MILLÓN" if millones == 1 else numero_a_letras(millones) + " MILLONES")
    if miles:
        partes.append("MIL" if miles == 1 else _letras_centenas(miles) + " MIL")
    if unidades:
        partes.append(_letras_centenas(unidades))
    return " ".join(partes)


def importe_con_letra(total):
    centavos_total = int(round(float(total) * 100))
    entero, centavos = divmod(centavos_total, 100)
    if entero == 1:
        moneda = "UN PESO"
    elif entero == 0:
        moneda = "CERO PESOS"
    else:
        moneda = f"{numero_a_letras(entero)} PESOS"
    return f"{moneda} {centavos:02d}/100 M.N."


def calcular_totales(conceptos, tasa):
    """Calcula subtotal, IVA y total con dos decimales, sin errores de punto flotante.

    Acepta conceptos de 5 elementos (descripcion, unidad, cantidad, precio, importe)
    o de 4 (descripcion, cantidad, precio, importe)."""
    subtotal = dinero(0)
    for c in conceptos:
        try:
            if len(c) >= 5:
                cantidad, precio = Decimal(str(c[2])), Decimal(str(c[3]))
            else:
                cantidad, precio = Decimal(str(c[1])), Decimal(str(c[2]))
            subtotal += dinero(cantidad * precio)
        except (InvalidOperation, ValueError, TypeError, IndexError):
            continue
    iva = dinero(subtotal * Decimal(str(tasa or 0)))
    return subtotal, iva, subtotal + iva


def normalizar_conceptos(conceptos):
    """Valida y normaliza los conceptos: descripcion, unidad, cantidad, precio e importe."""
    salida = []
    for concepto in conceptos:
        if len(concepto) >= 5:
            descripcion, unidad, cantidad, precio, importe = concepto[:5]
        elif len(concepto) == 4:
            descripcion, cantidad, precio, importe = concepto
            unidad = "SERVICIO"
        else:
            continue
        try:
            cantidad = Decimal(str(cantidad))
            precio = Decimal(str(precio))
        except (InvalidOperation, ValueError, TypeError):
            raise ReglaError("Uno de los conceptos tiene una cantidad o un precio que no es un número válido.")
        if not cantidad.is_finite() or not precio.is_finite() or cantidad <= 0 or precio <= 0:
            raise ReglaError("La cantidad y el precio de cada concepto deben ser mayores a cero.")
        if cantidad > LIMITE_MONTO or precio > LIMITE_MONTO:
            raise ReglaError("Un concepto supera el límite permitido.")
        descripcion = str(descripcion).strip()
        unidad = str(unidad or "SERVICIO").strip().upper()[:12] or "SERVICIO"
        if not texto_valido(descripcion):
            raise ReglaError("Escriba una descripción válida en cada concepto.")
        importe = dinero(cantidad * precio)
        if importe <= 0:
            raise ReglaError("El importe de un concepto debe ser mayor a cero.")
        salida.append((descripcion, unidad, cantidad, precio, importe))
    if not salida:
        raise ReglaError("Agregue al menos un concepto a la cotización.")
    return salida


def generar_excel_formal(ruta, idCot, cliente, idServicio, problema, conceptos,
                         subtotal, iva, total, tasa, datos=None):
    """Genera la cotización en Excel con formato formal: cuadros, totales e importe con letra."""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
        from datetime import date as _fecha, timedelta as _dias
    except ImportError:
        return ""

    datos = datos or {}
    if not ruta:
        carpeta = os.path.join(carpeta_programa(), "cotizaciones")
        os.makedirs(carpeta, exist_ok=True)
        ruta = os.path.join(carpeta, f"COT-{int(idCot):05d}.xlsx")

    azul = "1F3A5F"
    azul_medio = "2F5C8F"
    gris_fila = "F5F7FA"
    gris_texto = "5B6B80"
    linea = "C9D1DC"
    moneda = '"$"#,##0.00'

    libro = Workbook()
    hoja = libro.active
    hoja.title = "Cotización"
    hoja.sheet_view.showGridLines = False
    for columna, ancho in (("A", 2.5), ("B", 10), ("C", 11), ("D", 42), ("E", 17), ("F", 17), ("G", 2.5)):
        hoja.column_dimensions[columna].width = ancho

    def celda(ref, valor=None, negrita=False, tam=10, color="1B2432",
              alineacion="left", relleno=None, formato=None, envolver=False):
        c = hoja[ref]
        if valor is not None:
            c.value = valor
        c.font = Font(name="Calibri", bold=negrita, size=tam, color=color)
        c.alignment = Alignment(horizontal=alineacion, vertical="center", wrap_text=envolver)
        if relleno:
            c.fill = PatternFill("solid", start_color=relleno, end_color=relleno)
        if formato:
            c.number_format = formato
        return c

    def caja(rango, color=linea):
        lado = Side(style="thin", color=color)
        filas = hoja[rango]
        for i, fila in enumerate(filas):
            for j, c in enumerate(fila):
                c.border = Border(
                    left=lado if j == 0 else c.border.left,
                    right=lado if j == len(fila) - 1 else c.border.right,
                    top=lado if i == 0 else c.border.top,
                    bottom=lado if i == len(filas) - 1 else c.border.bottom)

    def rejilla(rango):
        lado = Side(style="thin", color=linea)
        for fila in hoja[rango]:
            for c in fila:
                c.border = Border(left=lado, right=lado, top=lado, bottom=lado)

    # --- Encabezado de la empresa ---
    for fila in (2, 3):
        hoja.merge_cells(f"B{fila}:F{fila}")
        for col in "BCDEF":
            hoja[f"{col}{fila}"].fill = PatternFill("solid", start_color=azul, end_color=azul)
    hoja.row_dimensions[2].height = 34
    hoja.row_dimensions[3].height = 18
    celda("B2", "THE GONZALEZ", negrita=True, tam=20, color="FFFFFF", relleno=azul)
    celda("B3", "Gestión de servicios profesionales", tam=9, color="D6E0EE", relleno=azul)

    try:
        vigencia_dias = int(datos.get("vigencia_dias") or 15)
    except (TypeError, ValueError):
        vigencia_dias = 15
    try:
        emision = _fecha.fromisoformat(str(datos.get("fecha") or ""))
    except (TypeError, ValueError):
        emision = _fecha.today()
    vence = emision + _dias(days=vigencia_dias)
    folio = f"COT-{int(idCot):05d}"

    # --- Título y folio ---
    hoja.merge_cells("B5:C5")
    celda("B5", "COTIZACIÓN", negrita=True, tam=16, color=azul)
    hoja.merge_cells("D5:F5")
    celda("D5", f"Folio {folio}", negrita=True, tam=11, color=azul, alineacion="right")
    hoja.row_dimensions[5].height = 26
    for col in "BCDEF":
        hoja[f"{col}6"].border = Border(bottom=Side(style="medium", color=azul))
        hoja[f"{col}6"].fill = PatternFill("solid", start_color=azul, end_color=azul)

    estado = str(datos.get("estado") or "REGISTRADA").replace("_", " ").title()
    bloque_izq = [
        ("Cliente", cliente or "-"),
        ("RFC", datos.get("rfc") or "No registrado"),
        ("Teléfono", datos.get("telefono") or "No registrado"),
        ("Correo", datos.get("correo") or "No registrado"),
        ("Dirección", datos.get("direccion") or "No registrada"),
        ("Servicio", f"#{idServicio}  {problema or ''}".strip()),
    ]
    bloque_der = [
        ("Folio", folio),
        ("Fecha de emisión", emision.strftime("%d/%m/%Y")),
        ("Vigencia", vence.strftime("%d/%m/%Y")),
        ("Estado", estado),
        ("Elaboró", datos.get("elaboro") or "-"),
        ("Moneda", "MXN - Peso mexicano"),
    ]
    for i, (etiqueta, valor) in enumerate(bloque_izq):
        fila = 7 + i
        celda(f"B{fila}", etiqueta, negrita=True, tam=9, color=gris_texto)
        celda(f"D{fila}", valor, tam=10, envolver=True)
        hoja.row_dimensions[fila].height = 19
    for i, (etiqueta, valor) in enumerate(bloque_der):
        fila = 7 + i
        celda(f"E{fila}", etiqueta, negrita=True, tam=9, color=gris_texto, alineacion="right")
        celda(f"F{fila}", valor, tam=10, alineacion="right", envolver=True)
    caja("B7:F12")

    # --- Tabla de conceptos ---
    fila = 14
    hoja.merge_cells(f"B{fila}:F{fila}")
    celda(f"B{fila}", "CONCEPTOS", negrita=True, tam=9, color="FFFFFF", relleno=azul_medio)
    hoja.row_dimensions[fila].height = 18

    fila += 1
    encabezados = ["Cant.", "Unidad", "Descripción", "Precio unitario", "Importe"]
    for col, texto in zip("BCDEF", encabezados):
        alineacion = "left" if col == "D" else ("center" if col in "BC" else "right")
        celda(f"{col}{fila}", texto, negrita=True, tam=9, color="FFFFFF",
              alineacion=alineacion, relleno=azul)
    hoja.row_dimensions[fila].height = 22
    fila_enc = fila

    for indice, (desc, unidad, cantidad, precio, importe) in enumerate(conceptos):
        fila += 1
        relleno = gris_fila if indice % 2 else None
        cantidad_f = float(cantidad)
        celda(f"B{fila}", cantidad_f, tam=10, alineacion="center", relleno=relleno,
              formato="#,##0" if cantidad_f.is_integer() else "#,##0.00")
        celda(f"C{fila}", unidad or "SERVICIO", tam=9, alineacion="center", relleno=relleno)
        celda(f"D{fila}", desc, tam=10, relleno=relleno, envolver=True)
        celda(f"E{fila}", float(precio), tam=10, alineacion="right", relleno=relleno, formato=moneda)
        celda(f"F{fila}", float(importe), tam=10, alineacion="right", relleno=relleno, formato=moneda)
        hoja.row_dimensions[fila].height = 21
    rejilla(f"B{fila_enc}:F{fila}")

    # --- Totales ---
    fila += 2
    for etiqueta, valor, destacado in (("Subtotal", subtotal, False),
                                       (f"IVA {float(tasa) * 100:g}%", iva, False),
                                       ("TOTAL", total, True)):
        if destacado:
            hoja.merge_cells(f"B{fila}:E{fila}")
            celda(f"B{fila}", etiqueta, negrita=True, tam=11, color="FFFFFF",
                  alineacion="right", relleno=azul)
            celda(f"F{fila}", float(valor), negrita=True, tam=11, color="FFFFFF",
                  alineacion="right", relleno=azul, formato=moneda)
            hoja.row_dimensions[fila].height = 24
        else:
            hoja.merge_cells(f"B{fila}:E{fila}")
            celda(f"B{fila}", etiqueta, tam=10, color=gris_texto, alineacion="right")
            celda(f"F{fila}", float(valor), tam=10, alineacion="right", formato=moneda)
            hoja.row_dimensions[fila].height = 18
        caja(f"B{fila}:F{fila}")
        fila += 1

    # --- Importe con letra ---
    fila += 1
    hoja.merge_cells(f"B{fila}:F{fila}")
    celda(f"B{fila}", f"Importe con letra: {importe_con_letra(total)}",
          negrita=True, tam=9, envolver=True)
    hoja.row_dimensions[fila].height = 26
    caja(f"B{fila}:F{fila}", color=azul)

    # --- Detalle del servicio ---
    fila += 2
    hoja.merge_cells(f"B{fila}:F{fila}")
    celda(f"B{fila}", "DETALLE DEL SERVICIO", negrita=True, tam=9, color="FFFFFF", relleno=azul_medio)
    hoja.row_dimensions[fila].height = 18
    inicio_detalle = fila
    for etiqueta, valor in (("Problema reportado", datos.get("problema") or problema or "-"),
                            ("Descripción del servicio", datos.get("descripcion_servicio") or "-"),
                            ("Observaciones", datos.get("observaciones") or "-")):
        fila += 1
        celda(f"B{fila}", etiqueta, negrita=True, tam=9, color=gris_texto, envolver=True)
        hoja.merge_cells(f"C{fila}:F{fila}")
        celda(f"C{fila}", valor, tam=9, envolver=True)
        hoja.row_dimensions[fila].height = 30
    caja(f"B{inicio_detalle}:F{fila}")

    # --- Condiciones comerciales ---
    fila += 2
    hoja.merge_cells(f"B{fila}:F{fila}")
    celda(f"B{fila}", "CONDICIONES COMERCIALES", negrita=True, tam=9, color="FFFFFF", relleno=azul_medio)
    hoja.row_dimensions[fila].height = 18
    inicio_cond = fila
    condiciones = [
        "Precios expresados en pesos mexicanos (MXN) y sujetos a cambio sin previo aviso.",
        f"Vigencia de la cotización: {vigencia_dias} días naturales a partir de la fecha de emisión.",
        f"Forma de pago: {datos.get('forma_pago') or 'A convenir con el cliente'}.",
        f"Tiempo de entrega: {datos.get('tiempo_entrega') or 'a convenir con el cliente'}.",
        f"Garantía: {datos.get('garantia') or 'según las condiciones del servicio'}.",
        "El IVA se calcula a la tasa vigente al momento de la cotización.",
    ]
    for texto in condiciones:
        fila += 1
        hoja.merge_cells(f"B{fila}:F{fila}")
        celda(f"B{fila}", "•  " + texto, tam=9, color=gris_texto, envolver=True)
        hoja.row_dimensions[fila].height = 18
    caja(f"B{inicio_cond}:F{fila}")

    # --- Firmas ---
    fila += 3
    hoja.merge_cells(f"B{fila}:C{fila + 1}")
    hoja.merge_cells(f"D{fila}:F{fila + 1}")
    caja(f"B{fila}:C{fila + 1}", color="1B2432")
    caja(f"D{fila}:F{fila + 1}", color="1B2432")
    celda(f"B{fila}", "Aceptada por el cliente", tam=9, color=gris_texto)
    hoja[f"B{fila}"].alignment = Alignment(horizontal="center", vertical="bottom")
    celda(f"D{fila}", "Elaboró", tam=9, color=gris_texto)
    hoja[f"D{fila}"].alignment = Alignment(horizontal="center", vertical="bottom")
    hoja.row_dimensions[fila].height = 24
    hoja.row_dimensions[fila + 1].height = 18

    fila += 3
    hoja.merge_cells(f"B{fila}:F{fila}")
    celda(f"B{fila}", f"Documento generado por el sistema The Gonzalez · Folio {folio}",
          tam=8, color=gris_texto, alineacion="center")

    hoja.print_area = f"A1:G{fila}"
    hoja.page_setup.orientation = "portrait"
    hoja.page_setup.fitToWidth = 1
    hoja.page_setup.fitToHeight = 1
    hoja.sheet_properties.pageSetUpPr.fitToPage = True
    hoja.print_options.horizontalCentered = True
    hoja.page_margins.left = hoja.page_margins.right = 0.5
    hoja.page_margins.top = hoja.page_margins.bottom = 0.6
    hoja.oddFooter.center.text = "Página &P de &N"
    hoja.oddFooter.center.size = 8

    if not str(ruta).lower().endswith(".xlsx"):
        ruta = os.path.splitext(str(ruta))[0] + ".xlsx"
    carpeta_cotizaciones = os.path.join(carpeta_programa(), "cotizaciones")
    candidatos = [
        os.path.abspath(ruta),
        os.path.join(carpeta_cotizaciones, f"COT-{int(idCot):05d}_{datetime.now():%Y%m%d%H%M%S}.xlsx"),
        os.path.join(os.path.abspath(os.path.dirname(ruta) or "."), f"COT-{int(idCot):05d}.xlsx"),
    ]
    for candidato in candidatos:
        try:
            os.makedirs(os.path.dirname(candidato), exist_ok=True)
            libro.save(candidato)
            return candidato
        except OSError:
            continue
    return ""


def cotizar_db(con, idServicio, idEmpleado, conceptos):
    """Registra la cotización, guarda sus conceptos con unidad y genera un solo Excel formal."""
    estado = estado_de(con, idServicio)
    if estado not in ("REGISTRADO", "ASIGNADO", "COTIZACION_GENERADA"):
        raise ReglaError(f"Este servicio ya no admite una nueva cotización (estado: "
                         f"{estado.lower().replace('_', ' ')}).")
    conceptos = normalizar_conceptos(conceptos)
    tasa = obtener_iva(con)
    subtotal, iva, total = calcular_totales(conceptos, tasa)
    if total <= 0:
        raise ReglaError("El total de la cotización debe ser mayor a cero.")
    if not con.execute("SELECT 1 FROM empleado WHERE idEmpleado = ? AND activo = 1;",
                       (idEmpleado,)).fetchone():
        raise ReglaError("La persona que elabora la cotización no existe o está inactiva.")
    info = con.execute(
        """SELECT c.nombre AS cliente, c.rfc, c.telefono, c.correo, c.direccion,
                  s.descripcionProblema AS problema, s.descripcionServicio, s.observaciones
           FROM servicio s JOIN cliente c ON c.idCliente = s.idCliente
           WHERE s.idServicio = ?;""", (idServicio,)).fetchone()
    if info is None:
        raise ReglaError("Este servicio ya no existe. Actualice la lista e intente de nuevo.")
    elaboro = con.execute(
        """SELECT nombre || ' ' || aPaterno FROM empleado WHERE idEmpleado = ?;""",
        (idEmpleado,)).fetchone()
    hoy = date.today().isoformat()
    try:
        with con:
            cur = con.execute(
                """INSERT INTO cotizacion (idServicio, idEmpleado, fechaCotizacion, subtotal, iva,
                                           tasaIva, total, estado)
                   VALUES (?, ?, ?, ?, ?, ?, ?, 'REGISTRADA');""",
                (idServicio, idEmpleado, hoy, float(subtotal), float(iva), float(tasa), float(total)))
            idCot = cur.lastrowid
            if not idCot:
                raise ReglaError("No se pudo registrar la cotización.")
            columnas_detalle = {f[1] for f in con.execute("PRAGMA table_info(detallecotizacion);").fetchall()}
            for descripcion, unidad, cantidad, precio, importe in conceptos:
                if "unidad" in columnas_detalle:
                    con.execute(
                        """INSERT INTO detallecotizacion (idCotizacion, descripcion, unidad, cantidad,
                                                         precioUnitario, importe)
                           VALUES (?, ?, ?, ?, ?, ?);""",
                        (idCot, descripcion, unidad, float(cantidad), float(precio), float(importe)))
                else:
                    con.execute(
                        """INSERT INTO detallecotizacion (idCotizacion, descripcion, cantidad,
                                                         precioUnitario, importe)
                           VALUES (?, ?, ?, ?, ?);""",
                        (idCot, descripcion, float(cantidad), float(precio), float(importe)))
                con.execute(
                    """INSERT INTO tg_detalle_cotizacion (idCotizacion, descripcion, unidad,
                                                          cantidad, precio, importe)
                       VALUES (?, ?, ?, ?, ?, ?);""",
                    (idCot, descripcion, unidad, float(cantidad), float(precio), float(importe)))
            con.execute("UPDATE servicio SET idEstado = ? WHERE idServicio = ?;",
                        (id_estado(con, "COTIZACION_GENERADA"), idServicio))
            registrar_movimiento(con, "COTIZAR", "cotizacion", idCot,
                                 f"COT-{int(idCot):05d} servicio {idServicio}; "
                                 f"{len(conceptos)} concepto(s); total {total}",
                                 importe=total)
    except ReglaError:
        raise
    except sqlite3.Error as e:
        registrar_error("COTIZAR", str(e))
        raise ReglaError("No se pudo guardar la cotización en la base de datos. "
                         "No se guardó nada; intente de nuevo.")

    ruta = os.path.join(carpeta_programa(), "cotizaciones", f"COT-{int(idCot):05d}.xlsx")
    datos = {
        "fecha": hoy, "estado": "REGISTRADA",
        "rfc": info["rfc"], "telefono": info["telefono"], "correo": info["correo"],
        "direccion": info["direccion"], "problema": info["problema"],
        "descripcion_servicio": info["descripcionServicio"], "observaciones": info["observaciones"],
        "elaboro": (elaboro[0] if elaboro else ""),
    }
    ruta_formal = generar_excel_formal(ruta, idCot, info["cliente"], idServicio, info["problema"],
                                       conceptos, subtotal, iva, total, tasa, datos)
    if ruta_formal:
        try:
            with con:
                con.execute(
                    "UPDATE cotizacion SET archivoCotizacion = ?, hashArchivo = ?, estadoDocumento = ? "
                    "WHERE idCotizacion = ?;",
                    (ruta_formal, sha256_archivo(ruta_formal), "GENERADO", idCot))
            ruta = ruta_formal
        except (OSError, sqlite3.Error) as e:
            registrar_error("EXCEL FORMAL", str(e))
            ruta = ""
    else:
        registrar_error("EXCEL FORMAL",
                        "No se pudo generar el Excel formal. Verifique que openpyxl esté instalado "
                        "y que la carpeta cotizaciones permita escribir.")
        ruta = ""
    return idCot, ruta, total


def aceptar_cotizacion_db(con, idServicio):
    estado = estado_de(con, idServicio)
    if estado != "COTIZACION_GENERADA":
        raise ReglaError("Solo se puede aceptar una cotización que ya fue generada.")
    with con:
        fila = con.execute(
            """SELECT idCotizacion FROM cotizacion
               WHERE idServicio = ?
                 AND UPPER(COALESCE(estado, '')) NOT IN ('ACEPTADA', 'ANULADA', 'RECHAZADA')
               ORDER BY idCotizacion DESC LIMIT 1;""", (idServicio,)).fetchone()
        if fila is None:
            raise ReglaError("No hay una cotización registrada para aceptar. Genere la cotización primero.")
        con.execute("UPDATE cotizacion SET estado = 'ACEPTADA' WHERE idCotizacion = ?;", (fila[0],))
        con.execute("UPDATE servicio SET idEstado = ? WHERE idServicio = ?;",
                    (id_estado(con, "ACEPTADO"), idServicio))
        registrar_movimiento(con, "ACEPTAR COTIZACION", "servicio", idServicio,
                             f"COT-{int(fila[0]):05d}")


def iniciar_trabajo_db(con, idServicio):
    estado = estado_de(con, idServicio)
    if estado != "ACEPTADO":
        raise ReglaError("Solo se puede iniciar un servicio cuya cotización fue aceptada.")
    if not con.execute("SELECT 1 FROM servicioempleado WHERE idServicio = ?;", (idServicio,)).fetchone():
        raise ReglaError("Antes de iniciar el trabajo, asigne al personal que lo realizará.")
    with con:
        con.execute("UPDATE servicio SET idEstado = ? WHERE idServicio = ?;", (id_estado(con, "EN_PROCESO"), idServicio))
        registrar_movimiento(con, "INICIAR TRABAJO", "servicio", idServicio)


def enviar_revision_db(con, idServicio):
    estado = estado_de(con, idServicio)
    if estado not in ("EN_PROCESO", "CORRECCION_SOLICITADA"):
        raise ReglaError("Solo se puede enviar a revisión un trabajo en proceso o con corrección pendiente.")
    with con:
        con.execute("UPDATE servicio SET idEstado = ? WHERE idServicio = ?;",
                    (id_estado(con, "PENDIENTE_REVISION"), idServicio))
        registrar_movimiento(con, "ENVIAR A REVISION", "servicio", idServicio)


def revisar_servicio_db(con, idServicio, idSupervisor, aprobado, comentarios):
    estado = estado_de(con, idServicio)
    if estado != "PENDIENTE_REVISION":
        raise ReglaError("Este servicio no está esperando revisión.")
    revisor = con.execute(
        """SELECT e.idEmpleado, e.activo, COALESCE(GROUP_CONCAT(r.nombreRol, ', '), '') AS puestos
           FROM empleado e LEFT JOIN puesto p ON p.idEmpleado = e.idEmpleado
           LEFT JOIN rol r ON r.idRol = p.idRol
           WHERE e.idEmpleado = ? GROUP BY e.idEmpleado;""", (idSupervisor,)).fetchone()
    if revisor is None or not revisor["activo"]:
        raise ReglaError("El supervisor seleccionado no existe o está inactivo.")
    puestos = {p.strip() for p in (revisor["puestos"] or "").split(",") if p.strip()}
    if not puestos & set(ROLES_REVISOR):
        raise ReglaError("La persona seleccionada no tiene puesto de supervisor. "
                         "Pida al gerente que revise sus puestos.")
    comentarios = (comentarios or "").strip() or None
    if not aprobado and not comentarios:
        raise ReglaError("Para solicitar una corrección escriba el motivo en los comentarios.")
    nuevo = "APROBADO" if aprobado else "CORRECCION_SOLICITADA"
    with con:
        con.execute("INSERT INTO revisionservicio (idServicio, idSupervisor, comentarios, aprobado) VALUES (?, ?, ?, ?);",
                    (idServicio, idSupervisor, comentarios, 1 if aprobado else 0))
        con.execute("UPDATE servicio SET idEstado = ? WHERE idServicio = ?;", (id_estado(con, nuevo), idServicio))
        registrar_movimiento(con, f"REVISION {nuevo}", "servicio", idServicio, comentarios)
    return nuevo


def facturar_db(con, idServicio):
    """Valida la cotización del servicio y factura exactamente el importe cotizado."""
    estado = estado_de(con, idServicio)
    if estado == "FACTURADO":
        raise ReglaError("Este servicio ya fue facturado. Revise la factura en la sección Facturas.")
    if estado == "FINALIZADO":
        raise ReglaError("Este servicio ya está finalizado y no se puede volver a facturar.")
    if estado != "APROBADO":
        raise ReglaError("Solo se factura un servicio aprobado en revisión. "
                         f"Este servicio está en estado: {estado.lower().replace('_', ' ')}.")
    if con.execute("SELECT 1 FROM factura WHERE idServicio = ?;", (idServicio,)).fetchone():
        raise ReglaError("Este servicio ya tiene una factura registrada. No se puede facturar dos veces.")
    cot = con.execute(
        """SELECT idCotizacion, subtotal, iva, tasaIva, total, estado
           FROM cotizacion WHERE idServicio = ? ORDER BY idCotizacion DESC LIMIT 1;""",
        (idServicio,)).fetchone()
    if cot is None:
        raise ReglaError("Este servicio no tiene cotización. Genere y acepte la cotización antes de facturar.")
    estado_cot = str(cot["estado"] or "").strip().upper()
    if estado_cot in ("ANULADA", "CANCELADA", "RECHAZADA"):
        raise ReglaError("La última cotización está cancelada. Genere una nueva cotización antes de facturar.")
    if estado_cot not in ("ACEPTADA", "REGISTRADA", ""):
        raise ReglaError(f"La cotización está en un estado que no permite facturar ({estado_cot.lower()}).")
    subtotal = dinero(cot["subtotal"] or 0)
    iva = dinero(cot["iva"] or 0)
    total = dinero(cot["total"] or 0)
    if total <= 0:
        raise ReglaError("La cotización no tiene un importe válido. Corrija la cotización antes de facturar.")
    if abs(dinero(subtotal + iva) - total) > Decimal("0.01"):
        raise ReglaError("El importe de la cotización no cuadra: subtotal más IVA no coincide con el total. "
                         "Genere de nuevo la cotización.")
    detalle = con.execute(
        "SELECT COALESCE(SUM(importe), 0) FROM tg_detalle_cotizacion WHERE idCotizacion = ?;",
        (cot["idCotizacion"],)).fetchone()[0]
    if not detalle:
        detalle = con.execute(
            "SELECT COALESCE(SUM(importe), 0) FROM detallecotizacion WHERE idCotizacion = ?;",
            (cot["idCotizacion"],)).fetchone()[0]
    if detalle and abs(dinero(detalle) - subtotal) > Decimal("0.01"):
        raise ReglaError("Los conceptos de la cotización no suman el subtotal registrado. "
                         "Genere de nuevo la cotización antes de facturar.")
    idCliente = con.execute("SELECT idCliente FROM servicio WHERE idServicio = ?;",
                            (idServicio,)).fetchone()[0]
    with con:
        if estado_cot != "ACEPTADA":
            con.execute("UPDATE cotizacion SET estado = 'ACEPTADA' WHERE idCotizacion = ?;",
                        (cot["idCotizacion"],))
            registrar_movimiento(con, "VALIDAR COTIZACION", "cotizacion", cot["idCotizacion"],
                                 f"COT-{int(cot['idCotizacion']):05d} validada al facturar",
                                 importe=total)
        cur = con.execute(
            """INSERT INTO factura (idServicio, idCliente, subtotal, iva, total, estado, fechaEmision)
               VALUES (?, ?, ?, ?, ?, 'GENERADA', ?);""",
            (idServicio, idCliente, float(subtotal), float(iva), float(total), date.today().isoformat()))
        numero = f"FAC-{cur.lastrowid:05d}"
        con.execute("UPDATE factura SET numeroFactura = ? WHERE idFactura = ?;", (numero, cur.lastrowid))
        con.execute("UPDATE servicio SET idEstado = ? WHERE idServicio = ?;",
                    (id_estado(con, "FACTURADO"), idServicio))
        registrar_movimiento(con, "FACTURAR", "factura", cur.lastrowid,
                             f"{numero} servicio {idServicio} cotizacion "
                             f"COT-{int(cot['idCotizacion']):05d}", importe=total)
    return numero, total, f"COT-{int(cot['idCotizacion']):05d}"


def cotizaciones_servicio_db(con, idServicio):
    return con.execute(
        """SELECT idCotizacion, fechaCotizacion, total, estado, archivoCotizacion, hashArchivo
           FROM cotizacion WHERE idServicio = ? ORDER BY idCotizacion DESC;""", (idServicio,)).fetchall()


class DialogoAsignar(QDialog):
    def __init__(self, parent, empleados):
        super().__init__(parent)
        self.setWindowTitle("Asignar personal")
        self.setMinimumSize(520, 380)
        v = QVBoxLayout(self)
        v.setContentsMargins(24, 22, 24, 20)
        v.setSpacing(12)
        lbl_titulo = QLabel("Asignar personal")
        lbl_titulo.setObjectName("dialogoTitulo")
        v.addWidget(lbl_titulo)
        ayuda = QLabel("Marque a las personas que trabajarán en el servicio y escriba su participación.")
        ayuda.setObjectName("ayuda")
        ayuda.setWordWrap(True)
        v.addWidget(ayuda)
        self.tabla = QTableWidget(len(empleados), 2)
        self.tabla.setHorizontalHeaderLabels(["Nombre", "Participación (ej. Técnico)"])
        self.tabla.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tabla.verticalHeader().setVisible(False)
        for fila, (idEmpleado, nombre) in enumerate(empleados):
            nom = QTableWidgetItem(nombre)
            nom.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
            nom.setCheckState(Qt.CheckState.Unchecked)
            nom.setData(Qt.ItemDataRole.UserRole, idEmpleado)
            self.tabla.setItem(fila, 0, nom)
            self.tabla.setItem(fila, 1, QTableWidgetItem("Técnico"))
        v.addWidget(self.tabla, 1)
        self.tabla.setCurrentCell(-1, -1)
        botones = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        botones.button(QDialogButtonBox.StandardButton.Save).setText("Asignar")
        botones.button(QDialogButtonBox.StandardButton.Cancel).setText("Cancelar")
        botones.accepted.connect(self.aceptar)
        botones.rejected.connect(self.reject)
        v.addWidget(botones)

    def valores(self):
        salida = []
        for fila in range(self.tabla.rowCount()):
            nom = self.tabla.item(fila, 0)
            if nom.checkState() == Qt.CheckState.Checked:
                item_part = self.tabla.item(fila, 1)
                participacion = (item_part.text() if item_part else "").strip()
                salida.append((nom.data(Qt.ItemDataRole.UserRole), participacion))
        return salida

    def aceptar(self):
        seleccion = self.valores()
        if not seleccion:
            aviso(self, "Marque al menos una persona.")
            return
        if any(not texto_valido(p) for _, p in seleccion):
            aviso(self, "Escriba la participación de cada persona marcada (por ejemplo: Técnico).")
            return
        self.accept()


class DialogoCotizacion(QDialog):
    def __init__(self, parent, empleados, tasa, idServicio, cliente, problema,
                 re_cotizacion=False):
        super().__init__(parent)
        self.setWindowTitle("Volver a cotizar" if re_cotizacion else "Nueva cotización")
        self.setMinimumSize(940, 600)
        self.tasa = tasa
        self.conceptos = []
        v = QVBoxLayout(self)
        v.setContentsMargins(24, 22, 24, 20)
        v.setSpacing(12)
        lbl_titulo = QLabel("Volver a cotizar" if re_cotizacion else "Nueva cotización")
        lbl_titulo.setObjectName("dialogoTitulo")
        v.addWidget(lbl_titulo)
        resumen = QLabel(f"Servicio #{idServicio}")
        resumen.setObjectName("seccion")
        v.addWidget(resumen)
        lbl_cliente = QLabel(f"Cliente: {cliente}")
        lbl_cliente.setObjectName("ayuda")
        v.addWidget(lbl_cliente)
        lbl_problema = QLabel(f"Problema: {problema}")
        lbl_problema.setObjectName("ayuda")
        lbl_problema.setWordWrap(True)
        v.addWidget(lbl_problema)

        fila_quien = QHBoxLayout()
        fila_quien.addWidget(etiqueta_campo("Elaborada por"))
        self.quien = QComboBox()
        for idEmpleado, nombre in empleados:
            self.quien.addItem(nombre, idEmpleado)
        fila_quien.addWidget(self.quien, 1)
        v.addLayout(fila_quien)

        campos = QHBoxLayout()
        campos.setSpacing(10)
        campos.addWidget(etiqueta_campo("Descripción"), 1)
        campos.addWidget(etiqueta_campo("Unidad"))
        campos.addWidget(etiqueta_campo("Cantidad"))
        campos.addWidget(etiqueta_campo("Precio unitario"))
        campos.addSpacing(96)
        v.addLayout(campos)

        captura = QHBoxLayout()
        captura.setSpacing(10)
        self.in_desc = QLineEdit()
        self.in_desc.setPlaceholderText("Ej. Mano de obra")
        self.in_unidad = QComboBox()
        self.in_unidad.setEditable(True)
        self.in_unidad.setFixedWidth(110)
        for unidad in ("SERVICIO", "PZA", "HORA", "DÍA", "KG", "M2", "M3", "LOTE", "VIAJE", "OTRO"):
            self.in_unidad.addItem(unidad)
        self.in_unidad.setCurrentText("SERVICIO")
        self.in_cant = QLineEdit()
        self.in_cant.setPlaceholderText("Ej. 2")
        self.in_cant.setFixedWidth(100)
        self.in_precio = QLineEdit()
        self.in_precio.setPlaceholderText("Ej. 1500.00")
        self.in_precio.setFixedWidth(120)
        btn_agregar = QPushButton("Agregar")
        btn_agregar.setObjectName("primario")
        btn_agregar.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_agregar.clicked.connect(self.agregar)
        self.in_precio.returnPressed.connect(self.agregar)
        captura.addWidget(self.in_desc, 1)
        captura.addWidget(self.in_unidad)
        captura.addWidget(self.in_cant)
        captura.addWidget(self.in_precio)
        captura.addWidget(btn_agregar)
        v.addLayout(captura)

        self.tabla = tabla_solo_lectura(5)
        self.tabla.setHorizontalHeaderLabels(
            ["Descripción", "Unidad", "Cantidad", "Precio", "Importe"])
        ajustar_columnas(self.tabla, {
            0: QHeaderView.ResizeMode.Stretch,
            1: QHeaderView.ResizeMode.ResizeToContents,
            2: QHeaderView.ResizeMode.ResizeToContents,
            3: QHeaderView.ResizeMode.ResizeToContents,
            4: QHeaderView.ResizeMode.ResizeToContents,
        })
        v.addWidget(self.tabla, 1)

        pie = QHBoxLayout()
        btn_quitar = QPushButton("Quitar seleccionado")
        btn_quitar.setObjectName("secundario")
        btn_quitar.clicked.connect(self.quitar)
        self.lbl_totales = QLabel("")
        self.lbl_totales.setObjectName("seccion")
        pie.addWidget(btn_quitar)
        pie.addStretch()
        pie.addWidget(self.lbl_totales)
        v.addLayout(pie)

        botones = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        botones.button(QDialogButtonBox.StandardButton.Save).setText("Generar cotización")
        botones.button(QDialogButtonBox.StandardButton.Cancel).setText("Cancelar")
        botones.accepted.connect(self.aceptar)
        botones.rejected.connect(self.reject)
        v.addWidget(botones)
        self.refrescar()

    def agregar(self):
        desc = self.in_desc.text().strip()
        if not texto_valido(desc):
            aviso(self, "Escriba una descripción con al menos 3 letras.")
            return
        unidad = self.in_unidad.currentText().strip().upper()
        if not unidad:
            aviso(self, "Escriba la unidad de medida del concepto (PZA, SERVICIO, HORA, etc.).")
            return
        cantidad = numero_desde_texto(self.in_cant.text())
        if cantidad is None or not (0 < cantidad <= LIMITE_MONTO):
            aviso(self, "La cantidad debe ser un número mayor a cero (ejemplo: 2).")
            return
        precio = numero_desde_texto(self.in_precio.text())
        if precio is None or not (0 < precio <= LIMITE_MONTO):
            aviso(self, "El precio debe ser un número mayor a cero (ejemplo: 1500 o 1,500.50).")
            return
        self.conceptos.append((desc, unidad, cantidad, precio, dinero(cantidad * precio)))
        self.in_desc.clear()
        self.in_cant.clear()
        self.in_precio.clear()
        self.in_desc.setFocus()
        self.refrescar()

    def quitar(self):
        fila = self.tabla.currentRow()
        if fila < 0:
            aviso(self, "Seleccione el concepto que desea quitar.")
            return
        del self.conceptos[fila]
        self.refrescar()

    def refrescar(self):
        self.tabla.setRowCount(0)
        for desc, unidad, cantidad, precio, importe in self.conceptos:
            fila = self.tabla.rowCount()
            self.tabla.insertRow(fila)
            valores = [desc, unidad, f"{cantidad:g}", f"${precio:,.2f}", f"${importe:,.2f}"]
            for col, valor in enumerate(valores):
                poner_celda(self.tabla, fila, col, valor)
        if not self.conceptos:
            self.lbl_totales.setText("Sin conceptos agregados")
            return
        subtotal, iva, total = calcular_totales(self.conceptos, self.tasa)
        self.lbl_totales.setText(
            f"Subtotal ${subtotal:,.2f}    IVA {self.tasa * 100:g}% ${iva:,.2f}    Total ${total:,.2f}")

    def valores(self):
        return self.quien.currentData(), list(self.conceptos)

    def aceptar(self):
        if not self.conceptos:
            aviso(self, "Agregue al menos un concepto a la cotización.")
            return
        self.accept()


class DialogoRevision(QDialog):
    def __init__(self, parent, supervisores):
        super().__init__(parent)
        self.setWindowTitle("Revisión del servicio")
        self.setMinimumWidth(460)
        v = QVBoxLayout(self)
        v.setContentsMargins(24, 22, 24, 20)
        v.setSpacing(12)
        lbl_titulo = QLabel("Revisión del servicio")
        lbl_titulo.setObjectName("dialogoTitulo")
        v.addWidget(lbl_titulo)
        self.supervisor = QComboBox()
        for idEmpleado, nombre in supervisores:
            self.supervisor.addItem(nombre, idEmpleado)
        self.aprobar = QRadioButton("Aprobar el servicio")
        self.aprobar.setChecked(True)
        self.corregir = QRadioButton("Solicitar corrección")
        v.addWidget(etiqueta_campo("Revisado por"))
        v.addWidget(self.supervisor)
        v.addWidget(self.aprobar)
        v.addWidget(self.corregir)
        v.addWidget(etiqueta_campo("Comentarios"))
        self.comentarios = QTextEdit()
        self.comentarios.setPlaceholderText("Obligatorio si solicita corrección")
        self.comentarios.setMaximumHeight(100)
        v.addWidget(self.comentarios)
        botones = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        botones.button(QDialogButtonBox.StandardButton.Save).setText("Guardar revisión")
        botones.button(QDialogButtonBox.StandardButton.Cancel).setText("Cancelar")
        botones.accepted.connect(self.aceptar)
        botones.rejected.connect(self.reject)
        v.addWidget(botones)

    def valores(self):
        texto = self.comentarios.toPlainText().strip() or None
        return self.supervisor.currentData(), self.aprobar.isChecked(), texto

    def aceptar(self):
        if self.corregir.isChecked() and not texto_valido(self.comentarios.toPlainText().strip()):
            aviso(self, "Para solicitar una corrección, escriba qué debe corregirse.")
            return
        self.accept()


class DialogoDetalleCotizacion(QDialog):
    def __init__(self, parent, idCotizacion):
        super().__init__(parent)
        self.setWindowTitle(f"Conceptos de COT-{int(idCotizacion):05d}")
        self.setMinimumSize(760, 420)
        v = QVBoxLayout(self)
        v.setContentsMargins(24, 22, 24, 20)
        v.setSpacing(10)
        titulo = QLabel(f"Conceptos de COT-{int(idCotizacion):05d}")
        titulo.setObjectName("dialogoTitulo")
        v.addWidget(titulo)
        tabla = tabla_solo_lectura(5)
        tabla.setHorizontalHeaderLabels(
            ["Descripción", "Unidad", "Cantidad", "Precio", "Importe"])
        ajustar_columnas(tabla, {
            0: QHeaderView.ResizeMode.Stretch,
            1: QHeaderView.ResizeMode.ResizeToContents,
            2: QHeaderView.ResizeMode.ResizeToContents,
            3: QHeaderView.ResizeMode.ResizeToContents,
            4: QHeaderView.ResizeMode.ResizeToContents,
        })
        try:
            con = conectar()
            try:
                filas = con.execute(
                    """SELECT descripcion, unidad, cantidad, precio, importe
                       FROM tg_detalle_cotizacion WHERE idCotizacion = ? ORDER BY idDetalle;""",
                    (idCotizacion,)).fetchall()
                if not filas:
                    filas = con.execute(
                        """SELECT descripcion, COALESCE(unidad, 'SERVICIO') AS unidad, cantidad,
                                  precioUnitario AS precio, importe
                           FROM detallecotizacion WHERE idCotizacion = ? ORDER BY idDetalle;""",
                        (idCotizacion,)).fetchall()
            finally:
                con.close()
        except sqlite3.Error:
            filas = []
        total = 0.0
        for f in filas:
            fila = tabla.rowCount()
            tabla.insertRow(fila)
            valores = [f["descripcion"], f["unidad"], f"{f['cantidad']:g}",
                       f"${f['precio']:,.2f}", f"${f['importe']:,.2f}"]
            for col, valor in enumerate(valores):
                poner_celda(tabla, fila, col, valor)
            total += float(f["importe"] or 0)
        alinear_tabla(tabla, (1, 2), izquierda=(0,), derecha=(3, 4))
        v.addWidget(tabla, 1)
        if not filas:
            vacio = QLabel("Esta cotización no tiene conceptos guardados.")
            vacio.setObjectName("vacio")
            vacio.setAlignment(Qt.AlignmentFlag.AlignCenter)
            v.addWidget(vacio)
        pie = QHBoxLayout()
        pie.addStretch()
        lbl = QLabel(f"Suma de conceptos: ${total:,.2f}")
        lbl.setObjectName("seccion")
        pie.addWidget(lbl)
        v.addLayout(pie)
        botones = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        botones.button(QDialogButtonBox.StandardButton.Close).setText("Cerrar")
        botones.rejected.connect(self.accept)
        v.addWidget(botones)


class DialogoCotizaciones(QDialog):
    def __init__(self, parent, idServicio):
        super().__init__(parent)
        self.setWindowTitle(f"Cotizaciones del servicio {idServicio}")
        self.setMinimumSize(780, 430)
        self.filas = []
        v = QVBoxLayout(self)
        v.setContentsMargins(24, 22, 24, 20)
        v.setSpacing(10)
        lbl_titulo = QLabel(f"Cotizaciones del servicio #{idServicio}")
        lbl_titulo.setObjectName("dialogoTitulo")
        v.addWidget(lbl_titulo)
        cliente = "-"
        problema = "-"
        try:
            con = conectar()
            try:
                info = con.execute(
                    """SELECT c.nombre AS cliente, s.descripcionProblema AS problema
                       FROM servicio s JOIN cliente c ON c.idCliente = s.idCliente
                       WHERE s.idServicio = ?;""", (idServicio,)).fetchone()
                if info is not None:
                    cliente = info["cliente"]
                    problema = info["problema"]
            finally:
                con.close()
        except sqlite3.Error:
            pass
        lbl_cliente = QLabel(f"Cliente: {cliente}")
        lbl_cliente.setObjectName("ayuda")
        v.addWidget(lbl_cliente)
        lbl_problema = QLabel(f"Problema: {problema}")
        lbl_problema.setObjectName("ayuda")
        lbl_problema.setWordWrap(True)
        v.addWidget(lbl_problema)
        self.tabla = tabla_solo_lectura(5)
        self.tabla.setHorizontalHeaderLabels(["Folio", "Fecha", "Total", "Estado", "Archivo"])
        self.tabla.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        v.addWidget(self.tabla, 1)
        self.vacio = QLabel("Este servicio todavía no tiene cotizaciones.")
        self.vacio.setObjectName("vacio")
        self.vacio.setAlignment(Qt.AlignmentFlag.AlignCenter)
        v.addWidget(self.vacio)
        try:
            con = conectar()
            try:
                self.filas = cotizaciones_servicio_db(con, idServicio)
            finally:
                con.close()
        except sqlite3.Error:
            self.filas = []
        for f in self.filas:
            fila = self.tabla.rowCount()
            self.tabla.insertRow(fila)
            archivo = os.path.basename(f["archivoCotizacion"]) if f["archivoCotizacion"] else "-"
            estado = str(f["estado"] or "").strip().upper()
            etiqueta = "Aceptada" if estado == "ACEPTADA" else (
                "Anulada" if estado == "ANULADA" else "Registrada")
            valores = [f"COT-{f['idCotizacion']:05d}", fecha_legible(f["fechaCotizacion"]),
                       f"${f['total']:,.2f}", etiqueta, archivo]
            for col, valor in enumerate(valores):
                self.tabla.setItem(fila, col, QTableWidgetItem(valor))
        alinear_tabla(self.tabla, (0, 1, 3), izquierda=(4,), derecha=(2,))
        self.tabla.setVisible(bool(self.filas))
        self.vacio.setVisible(not self.filas)
        acciones = QHBoxLayout()
        btn_abrir = QPushButton("Abrir Excel")
        btn_abrir.setObjectName("primario")
        btn_abrir.clicked.connect(self.abrir)
        btn_conceptos = QPushButton("Ver conceptos")
        btn_conceptos.setObjectName("secundario")
        btn_conceptos.setToolTip("Muestra los conceptos, unidades y precios de la cotización.")
        btn_conceptos.clicked.connect(self.ver_conceptos)
        btn_verificar = QPushButton("Verificar archivo")
        btn_verificar.setObjectName("secundario")
        btn_verificar.setToolTip("Comprueba que el Excel no haya sido modificado después de generarse.")
        btn_verificar.clicked.connect(self.verificar)
        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.setObjectName("secundario")
        btn_cerrar.clicked.connect(self.accept)
        acciones.addWidget(btn_abrir)
        acciones.addWidget(btn_conceptos)
        acciones.addWidget(btn_verificar)
        acciones.addStretch()
        acciones.addWidget(btn_cerrar)
        v.addLayout(acciones)

    def _seleccionada(self):
        fila = self.tabla.currentRow()
        if fila < 0 or fila >= len(self.filas):
            aviso(self, "Seleccione una cotización de la lista.")
            return None
        return self.filas[fila]

    def ver_conceptos(self):
        f = self._seleccionada()
        if f is None:
            return
        DialogoDetalleCotizacion(self, f["idCotizacion"]).exec()

    def abrir(self):
        f = self._seleccionada()
        if f is None:
            return
        if not f["archivoCotizacion"] or not os.path.exists(f["archivoCotizacion"]):
            aviso(self, "No encontramos el archivo Excel de esta cotización. Puede que se haya movido o borrado.")
            return
        abrir_archivo(f["archivoCotizacion"])

    def verificar(self):
        f = self._seleccionada()
        if f is None:
            return
        if not f["archivoCotizacion"] or not os.path.exists(f["archivoCotizacion"]):
            aviso(self, "No encontramos el archivo Excel de esta cotización.")
            return
        if not f["hashArchivo"]:
            aviso(self, "Esta cotización se creó antes de activar la verificación, por eso no se puede comprobar.")
            return
        try:
            actual = sha256_archivo(f["archivoCotizacion"])
        except OSError:
            aviso(self, "No se pudo leer el archivo para verificarlo. "
                        "Puede estar abierto en otro programa.")
            return
        if actual == f["hashArchivo"]:
            listo(self, "El archivo es el original y no ha sido modificado.")
        else:
            aviso(self, "Atención: el archivo cambió después de generarse. "
                        "Avise al gerente antes de usarlo.")


def _sin_marca_de_clic(args):
    """Los botones envían un bool (marcado) al hacer clic; los métodos no lo usan."""
    if len(args) == 1 and isinstance(args[0], bool):
        return ()
    return args


def protegido(metodo):
    @functools.wraps(metodo)
    def envoltura(self, *args, **kwargs):
        try:
            return metodo(self, *_sin_marca_de_clic(args), **kwargs)
        except ReglaError as e:
            aviso(self, str(e))
        except sqlite3.OperationalError as e:
            if "locked" in str(e).lower() or "busy" in str(e).lower():
                aviso(self, MSG_OCUPADA)
            else:
                fallo(self, describir_error_bd(e) or MSG_GENERAL, SUGERENCIA_DB, str(e))
        except sqlite3.IntegrityError as e:
            fallo(self, describir_error_bd(e) or "No se pudo guardar porque un dato ya existe o no cumple una regla.",
                  SUGERENCIA_DB, str(e))
        except OSError as e:
            fallo(self, MSG_ARCHIVO, SUGERENCIA_ARCHIVO, str(e))
        except sqlite3.Error as e:
            fallo(self, describir_error_bd(e) or MSG_GENERAL, SUGERENCIA_DB, str(e))
        except Exception as e:  # evita que un fallo inesperado cierre el programa
            fallo(self, MSG_GENERAL, SUGERENCIA_INESPERADO, str(e))
        return None
    return envoltura


def con_espera(metodo):
    @functools.wraps(metodo)
    def envoltura(self, *args, **kwargs):
        self.iniciar_espera()
        try:
            return protegido(metodo)(self, *args, **kwargs)
        finally:
            self.terminar_espera()
    return envoltura


class Indicador(QWidget):
    def __init__(self):
        super().__init__()
        self.setFixedSize(18, 18)
        self.angulo = 0
        self.timer = QTimer(self)
        self.timer.timeout.connect(self._girar)
        self.hide()

    def _girar(self):
        self.angulo = (self.angulo + 30) % 360
        self.update()

    def paintEvent(self, _evento):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        pluma = QPen(QColor("#1F6FB2"))
        pluma.setWidth(3)
        pluma.setCapStyle(Qt.PenCapStyle.RoundCap)
        p.setPen(pluma)
        p.drawArc(3, 3, 12, 12, -self.angulo * 16, 270 * 16)
        p.end()

    def iniciar(self):
        self.show()
        self.timer.start(70)

    def detener(self):
        self.timer.stop()
        self.hide()


class Avatar(QWidget):
    def __init__(self, nombre="", foto="", tamano=36, parent=None):
        super().__init__(parent)
        self.nombre = nombre
        self.foto = foto
        self.tamano = tamano
        self.setFixedSize(tamano, tamano)

    def set_datos(self, nombre, foto):
        self.nombre = nombre or ""
        self.foto = foto or ""
        self.update()

    def paintEvent(self, _evento):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        rect = self.rect().adjusted(1, 1, -1, -1)
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(QColor("#2F6FED"))
        p.drawEllipse(rect)
        pixmap = QPixmap(self.foto) if self.foto else QPixmap()
        if not pixmap.isNull():
            lado = min(pixmap.width(), pixmap.height())
            origen = QRect((pixmap.width() - lado) // 2, (pixmap.height() - lado) // 2,
                           lado, lado)
            camino = QPainterPath()
            camino.addEllipse(QRectF(rect))
            p.setClipPath(camino)
            p.drawPixmap(rect, pixmap.copy(origen))
        else:
            p.setPen(QColor("#FFFFFF"))
            fuente = QFont(self.font())
            fuente.setPixelSize(max(10, self.tamano // 3))
            fuente.setBold(True)
            p.setFont(fuente)
            p.drawText(rect, Qt.AlignmentFlag.AlignCenter, iniciales(self.nombre))
        p.end()


class DialogoConfigCorreo(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Configurar Gmail")
        self.setMinimumWidth(480)
        ajustes = QSettings("TheGonzalez", "Servicios")
        v = QVBoxLayout(self)
        v.setContentsMargins(26, 24, 26, 22)
        v.setSpacing(10)
        titulo = QLabel("Envío de avisos por Gmail")
        titulo.setObjectName("dialogoTitulo")
        v.addWidget(titulo)
        ayuda = QLabel("Se usa una cuenta de Gmail y una contraseña de aplicación. "
                       "No escriba su contraseña normal de Gmail. "
                       "Si existen las variables TG_GMAIL_USUARIO y TG_GMAIL_CLAVE, "
                       "tienen prioridad sobre estos datos.")
        ayuda.setObjectName("ayuda")
        ayuda.setWordWrap(True)
        v.addWidget(ayuda)
        self.correo = QLineEdit(str(ajustes.value("correo/usuario", "") or ""))
        self.correo.setPlaceholderText("cuenta@gmail.com")
        self.clave = QLineEdit(str(ajustes.value("correo/clave", "") or ""))
        self.clave.setPlaceholderText("Contraseña de aplicación de Gmail")
        self.clave.setEchoMode(QLineEdit.EchoMode.Password)
        self.admin = QLineEdit(str(ajustes.value("correo/admin", "") or ""))
        self.admin.setPlaceholderText("Correo del administrador (opcional)")
        self.lbl_error = QLabel("")
        self.lbl_error.setObjectName("error")
        self.lbl_error.setWordWrap(True)
        self.lbl_error.hide()
        for etiqueta, campo in (("Cuenta de Gmail", self.correo),
                                ("Contraseña de aplicación", self.clave),
                                ("Correo del administrador", self.admin)):
            v.addWidget(etiqueta_campo(etiqueta))
            v.addWidget(campo)
        v.addWidget(self.lbl_error)
        botones = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        botones.button(QDialogButtonBox.StandardButton.Save).setText("Guardar")
        botones.button(QDialogButtonBox.StandardButton.Cancel).setText("Cancelar")
        botones.accepted.connect(self.aceptar)
        botones.rejected.connect(self.reject)
        v.addWidget(botones)
        self.correo.setFocus()

    def aceptar(self):
        self.lbl_error.hide()
        if self.correo.text().strip() and not correo_ok(self.correo.text().strip()):
            self.lbl_error.setText("La cuenta de Gmail no parece válida.")
            self.lbl_error.show()
            return
        if self.admin.text().strip() and not correo_ok(self.admin.text().strip()):
            self.lbl_error.setText("El correo del administrador no parece válido.")
            self.lbl_error.show()
            return
        ajustes = QSettings("TheGonzalez", "Servicios")
        ajustes.setValue("correo/usuario", self.correo.text().strip())
        ajustes.setValue("correo/clave", self.clave.text())
        ajustes.setValue("correo/admin", self.admin.text().strip())
        self.accept()


class DialogoSolicitudRol(QDialog):
    def __init__(self, parent, rol_actual, nombre):
        super().__init__(parent)
        self.setWindowTitle("Solicitar cambio de puesto")
        self.setFixedWidth(500)
        self.rol_actual = str(rol_actual or "").strip().upper()
        v = QVBoxLayout(self)
        v.setContentsMargins(26, 24, 26, 22)
        v.setSpacing(10)
        titulo = QLabel("Solicitar cambio de puesto")
        titulo.setObjectName("dialogoTitulo")
        v.addWidget(titulo)
        ayuda = QLabel("La solicitud se guarda en el sistema y se envía por Gmail al administrador. "
                       "El administrador la revisa y la aprueba desde el panel de administración.")
        ayuda.setObjectName("ayuda")
        ayuda.setWordWrap(True)
        v.addWidget(ayuda)
        self.rol = QComboBox()
        for codigo in ROLES_SISTEMA:
            if codigo != self.rol_actual:
                self.rol.addItem(ROLES_TEXTO.get(codigo, codigo.capitalize()), codigo)
        indice = self.rol.findData("GERENTE" if self.rol_actual == "COTIZADOR" else "COTIZADOR")
        if indice >= 0:
            self.rol.setCurrentIndex(indice)
        self.motivo = QTextEdit()
        self.motivo.setFixedHeight(100)
        self.motivo.setPlaceholderText("Explique brevemente por qué solicita el cambio.")
        self.lbl_error = QLabel("")
        self.lbl_error.setObjectName("error")
        self.lbl_error.setWordWrap(True)
        self.lbl_error.hide()
        v.addWidget(etiqueta_campo("Puesto actual"))
        actual = QLineEdit(ROLES_TEXTO.get(self.rol_actual, self.rol_actual.capitalize()))
        actual.setReadOnly(True)
        v.addWidget(actual)
        v.addWidget(etiqueta_campo("Puesto solicitado"))
        v.addWidget(self.rol)
        v.addWidget(etiqueta_campo("Motivo"))
        v.addWidget(self.motivo)
        v.addWidget(self.lbl_error)
        self.lbl_envio = QLabel("")
        self.lbl_envio.setObjectName("ayuda")
        self.lbl_envio.setWordWrap(True)
        v.addWidget(self.lbl_envio)
        self.btn_config = QPushButton("Configurar Gmail")
        self.btn_config.setObjectName("enlace")
        self.btn_config.clicked.connect(self.configurar)
        v.addWidget(self.btn_config, 0, Qt.AlignmentFlag.AlignRight)
        botones = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        botones.button(QDialogButtonBox.StandardButton.Save).setText("Enviar solicitud")
        botones.button(QDialogButtonBox.StandardButton.Cancel).setText("Cancelar")
        botones.accepted.connect(self.aceptar)
        botones.rejected.connect(self.reject)
        v.addWidget(botones)
        self.motivo.setFocus()
        self.actualizar_aviso()

    def configurar(self):
        DialogoConfigCorreo(self).exec()
        self.actualizar_aviso()

    def actualizar_aviso(self):
        usuario, clave, admin = config_correo()
        if usuario and clave:
            self.lbl_envio.setText("El aviso se enviará desde " + usuario + ".")
        else:
            self.lbl_envio.setText("Gmail no está configurado. La solicitud se guardará y el "
                                   "panel de administración la mostrará igualmente.")

    def aceptar(self):
        self.lbl_error.hide()
        motivo = self.motivo.toPlainText().strip()
        nuevo = self.rol.currentData()
        if nuevo == self.rol_actual:
            self.lbl_error.setText("Seleccione un puesto diferente al actual.")
            self.lbl_error.show()
            return
        if not texto_valido(motivo):
            self.lbl_error.setText("Escriba un motivo con al menos 3 letras.")
            self.lbl_error.show()
            return
        try:
            con = conectar()
        except sqlite3.Error:
            self.lbl_error.setText(MSG_BD)
            self.lbl_error.show()
            return
        destino = ""
        try:
            with con:
                con.execute(
                    "INSERT INTO solicitud_rol (idUsuario, usuario, nombre, rolActual, rolSolicitado, "
                    "motivo, estado, fecha) VALUES (?, ?, ?, ?, ?, ?, 'PENDIENTE', ?);",
                    (SESION.get("id"), SESION.get("nombre"), SESION.get("nombreCompleto"),
                     self.rol_actual, nuevo, motivo,
                     datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                registrar_movimiento(con, "SOLICITAR CAMBIO", "usuario", SESION.get("id"),
                                     f"{self.rol_actual} -> {nuevo}")
            destino = correo_administrador(con)
        except sqlite3.Error as e:
            fallo(self, "No se pudo guardar la solicitud.", SUGERENCIA_DB, str(e))
            return
        finally:
            con.close()
        usuario_gmail, clave_gmail, _admin = config_correo()
        if destino and usuario_gmail and clave_gmail:
            cuerpo = (f"Solicitud de cambio de puesto\n\n"
                      f"Nombre: {SESION.get('nombreCompleto')}\n"
                      f"Usuario: {SESION.get('nombre')}\n"
                      f"Puesto actual: {ROLES_TEXTO.get(self.rol_actual, self.rol_actual)}\n"
                      f"Puesto solicitado: {ROLES_TEXTO.get(nuevo, nuevo)}\n"
                      f"Motivo: {motivo}\n"
                      f"Fecha: {datetime.now():%d/%m/%Y %H:%M}\n")
            try:
                enviar_correo(destino, f"Solicitud de cambio de puesto: {SESION.get('nombre')}", cuerpo)
                listo(self, "Solicitud guardada y enviada por Gmail al gerente.")
            except Exception as e:
                registrar_error("ENVÍO DE CORREO", str(e))
                aviso(self, "La solicitud se guardó, pero no se pudo enviar el correo.\n\n"
                            "Revise la cuenta de Gmail y la contraseña de aplicación en "
                            "Configurar Gmail. El panel de administración la mostrará igualmente.")
        else:
            listo(self, "Solicitud guardada. El panel de administración la mostrará.\n"
                        "Configure Gmail para enviar el aviso automáticamente.")
        self.accept()


class DialogoNotificaciones(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Avisos y actualizaciones")
        self.setMinimumSize(640, 470)
        v = QVBoxLayout(self)
        v.setContentsMargins(24, 22, 24, 20)
        v.setSpacing(10)
        titulo = QLabel("Avisos y actualizaciones")
        titulo.setObjectName("dialogoTitulo")
        v.addWidget(titulo)
        self.texto = QTextBrowser()
        self.texto.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        v.addWidget(self.texto, 1)
        botones = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        boton = botones.button(QDialogButtonBox.StandardButton.Close)
        boton.setText("Cerrar")
        boton.setObjectName("secundario")
        botones.rejected.connect(self.accept)
        v.addWidget(botones)
        self.cargar()

    def cargar(self):
        id_usuario = SESION.get("id")
        try:
            con = conectar()
            try:
                filas = con.execute(
                    """SELECT titulo, mensaje, fecha, leida FROM notificacion
                       WHERE idUsuario = ? ORDER BY idNotificacion DESC LIMIT 100;""",
                    (id_usuario,)).fetchall()
                if filas:
                    con.execute("UPDATE notificacion SET leida = 1 WHERE idUsuario = ?;",
                                (id_usuario,))
                    con.commit()
            finally:
                con.close()
        except sqlite3.Error:
            filas = []
        if not filas:
            self.texto.setHtml(
                '<p style="color:#64748B; font-size:12pt;">No tienes avisos por el momento.</p>')
            return
        bloques = []
        for fila in filas:
            mensaje = escape(str(fila["mensaje"] or "")).replace("\n", "<br>")
            bloques.append(
                '<table width="100%" cellspacing="0" cellpadding="10" '
                'style="border:1px solid #E3E8F0; background-color:#FFFFFF;">'
                f'<tr><td><b style="color:#111C2E;">{escape(str(fila["titulo"] or "Aviso"))}</b>'
                f'<span style="color:#94A3B8; font-size:9pt;"> &nbsp; {fecha_legible(fila["fecha"])}</span>'
                f'<br><span style="color:#334155;">{mensaje}</span></td></tr></table><br>')
        self.texto.setHtml("".join(bloques))


class DialogoCuenta(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Mi cuenta")
        self.setMinimumWidth(620)
        self.idUsuario = SESION.get("id")
        self.fila = None
        self.claveHash = ""
        self.rol = ""
        self.foto_actual = ""
        self.foto_nueva = None
        self.quitar_foto = False
        self.cargar()
        v = QVBoxLayout(self)
        v.setContentsMargins(28, 24, 28, 22)
        v.setSpacing(12)
        titulo = QLabel("Mi cuenta")
        titulo.setObjectName("dialogoTitulo")
        v.addWidget(titulo)

        cabecera = QHBoxLayout()
        cabecera.setSpacing(14)
        self.avatar = Avatar(self.nombre_inicial, self.foto_inicial, 76)
        cabecera.addWidget(self.avatar)
        datos = QVBoxLayout()
        datos.setSpacing(2)
        self.lbl_nombre = QLabel(self.nombre_inicial)
        self.lbl_nombre.setObjectName("seccion")
        self.lbl_usuario = QLabel("Usuario: " + (self.usuario_inicial or ""))
        self.lbl_usuario.setObjectName("ayuda")
        self.lbl_rol = QLabel("Puesto: " + ROLES_TEXTO.get(self.rol, self.rol.capitalize() or "Sin puesto"))
        self.lbl_rol.setObjectName("ayuda")
        datos.addWidget(self.lbl_nombre)
        datos.addWidget(self.lbl_usuario)
        datos.addWidget(self.lbl_rol)
        datos.addStretch()
        cabecera.addLayout(datos, 1)
        botones_foto = QVBoxLayout()
        botones_foto.setSpacing(4)
        self.btn_foto = QPushButton("Cambiar foto")
        self.btn_foto.setObjectName("secundario")
        self.btn_foto.clicked.connect(self.elegir_foto)
        self.btn_quitar_foto = QPushButton("Quitar foto")
        self.btn_quitar_foto.setObjectName("secundario")
        self.btn_quitar_foto.clicked.connect(self.eliminar_foto)
        botones_foto.addWidget(self.btn_foto)
        botones_foto.addWidget(self.btn_quitar_foto)
        cabecera.addLayout(botones_foto)
        v.addLayout(cabecera)
        separador = QFrame()
        separador.setObjectName("separador")
        separador.setFixedHeight(1)
        v.addWidget(separador)

        seccion_datos = QLabel("Datos personales")
        seccion_datos.setObjectName("seccion")
        v.addWidget(seccion_datos)
        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.AllNonFixedFieldsGrow)
        form.setHorizontalSpacing(16)
        form.setVerticalSpacing(10)
        self.nombre = QLineEdit(self.nombre_inicial)
        self.correo = QLineEdit(self.correo_inicial)
        self.telefono = QLineEdit(self.telefono_inicial)
        self.usuario = QLineEdit(self.usuario_inicial)
        self.usuario.setReadOnly(True)
        self.puesto = QLineEdit(ROLES_TEXTO.get(self.rol, self.rol.capitalize() or "Sin puesto"))
        self.puesto.setReadOnly(True)
        self.nombre.setToolTip("Solo letras y espacios.")
        self.nombre.textChanged.connect(self.actualizar_avatar)
        self.correo.setToolTip("Formato nombre@dominio.com")
        self.telefono.setToolTip("10 dígitos, sin espacios ni guiones.")
        form.addRow(etiqueta_campo("Nombre completo"), self.nombre)
        form.addRow(etiqueta_campo("Correo electrónico"), self.correo)
        form.addRow(etiqueta_campo("Teléfono"), self.telefono)
        form.addRow(etiqueta_campo("Usuario"), self.usuario)
        form.addRow(etiqueta_campo("Puesto"), self.puesto)
        v.addLayout(form)

        seccion_clave = QLabel("Seguridad")
        seccion_clave.setObjectName("seccion")
        v.addWidget(seccion_clave)
        form_clave = QFormLayout()
        form_clave.setLabelAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        form_clave.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.AllNonFixedFieldsGrow)
        form_clave.setHorizontalSpacing(16)
        form_clave.setVerticalSpacing(10)
        self.clave_actual = QLineEdit()
        self.clave_actual.setEchoMode(QLineEdit.EchoMode.Password)
        self.clave_actual.setPlaceholderText("Su contraseña actual")
        self.clave_nueva = QLineEdit()
        self.clave_nueva.setEchoMode(QLineEdit.EchoMode.Password)
        self.clave_nueva.setPlaceholderText("Mínimo 8 caracteres")
        self.clave_repetir = QLineEdit()
        self.clave_repetir.setEchoMode(QLineEdit.EchoMode.Password)
        self.clave_repetir.setPlaceholderText("Repita la contraseña nueva")
        form_clave.addRow(etiqueta_campo("Contraseña actual"), self.clave_actual)
        form_clave.addRow(etiqueta_campo("Contraseña nueva"), self.clave_nueva)
        form_clave.addRow(etiqueta_campo("Repetir contraseña"), self.clave_repetir)
        v.addLayout(form_clave)
        v.addWidget(ayuda_campo("Deje las contraseñas vacías si no desea cambiarla."))
        self.lbl_error = QTextEdit()
        self.lbl_error.setObjectName("errorBox")
        self.lbl_error.setReadOnly(True)
        self.lbl_error.setFixedHeight(64)
        self.lbl_error.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.lbl_error.hide()
        v.addWidget(self.lbl_error)

        pie = QHBoxLayout()
        self.btn_solicitud = QPushButton("Solicitar cambio de puesto")
        self.btn_solicitud.setObjectName("secundario")
        self.btn_solicitud.clicked.connect(self.solicitar_cambio)
        pie.addWidget(self.btn_solicitud)
        pie.addStretch()
        btn_cancelar = QPushButton("Cancelar")
        btn_cancelar.setObjectName("secundario")
        btn_cancelar.clicked.connect(self.reject)
        btn_guardar = QPushButton("Guardar cambios")
        btn_guardar.setObjectName("primario")
        btn_guardar.clicked.connect(self.guardar)
        pie.addWidget(btn_cancelar)
        pie.addWidget(btn_guardar)
        v.addLayout(pie)
        self.nombre.setFocus()

    def cargar(self):
        self.nombre_inicial = ""
        self.usuario_inicial = ""
        self.correo_inicial = ""
        self.telefono_inicial = ""
        self.foto_inicial = ""
        try:
            con = conectar()
            try:
                columnas = {fila[1] for fila in con.execute("PRAGMA table_info(usuario);").fetchall()}
                campos = ["nombreUsuario", "nombreCompleto", "rol", "claveHash"]
                for opcional in ("correo", "telefono", "foto"):
                    if opcional in columnas:
                        campos.append(opcional)
                self.fila = con.execute(
                    f"SELECT {', '.join(campos)} FROM usuario WHERE idUsuario = ?;",
                    (self.idUsuario,)).fetchone()
            finally:
                con.close()
        except sqlite3.Error:
            self.fila = None
        if self.fila is not None:
            self.usuario_inicial = self.fila["nombreUsuario"] or ""
            self.nombre_inicial = self.fila["nombreCompleto"] or ""
            self.rol = str(self.fila["rol"] or "").strip().upper()
            self.claveHash = self.fila["claveHash"] or ""
            self.correo_inicial = self.fila["correo"] or "" if "correo" in self.fila.keys() else ""
            self.telefono_inicial = self.fila["telefono"] or "" if "telefono" in self.fila.keys() else ""
            self.foto_actual = self.fila["foto"] or "" if "foto" in self.fila.keys() else ""
            self.foto_inicial = self.foto_actual

    def actualizar_avatar(self, texto):
        foto = ""
        if not self.quitar_foto:
            foto = self.foto_nueva or self.foto_actual
        self.avatar.set_datos(texto.strip() or self.nombre_inicial, foto)

    def elegir_foto(self):
        ruta, _filtro = QFileDialog.getOpenFileName(
            self, "Elegir foto de perfil", "", "Imágenes (*.png *.jpg *.jpeg *.bmp *.webp)")
        if not ruta:
            return
        if QPixmap(ruta).isNull():
            aviso(self, "El archivo seleccionado no es una imagen válida.")
            return
        self.foto_nueva = ruta
        self.quitar_foto = False
        self.avatar.set_datos(self.nombre.text().strip() or self.nombre_inicial, ruta)

    def eliminar_foto(self):
        self.foto_nueva = None
        self.quitar_foto = True
        self.avatar.set_datos(self.nombre.text().strip() or self.nombre_inicial, "")

    def solicitar_cambio(self):
        DialogoSolicitudRol(self, self.rol, self.nombre.text().strip()).exec()

    def error(self, texto, campo=None):
        self.lbl_error.setPlainText(texto)
        self.lbl_error.show()
        self.adjustSize()
        marcar_error(campo, True) if campo is not None else None

    def guardar(self):
        if self.fila is None:
            self.error("La cuenta ya no está disponible. Cierre y vuelva a iniciar sesión.", None)
            return
        self.lbl_error.hide()
        for campo in (self.nombre, self.correo, self.telefono, self.clave_actual,
                      self.clave_nueva, self.clave_repetir):
            marcar_error(campo, False)
        nombre = self.nombre.text().strip()
        correo = self.correo.text().strip()
        telefono = self.telefono.text().strip()
        if not nombre_valido(nombre):
            self.error("Nombre completo: use solo letras y espacios.", self.nombre)
            return
        if correo and not correo_ok(correo):
            self.error("Correo: use el formato nombre@dominio.com.", self.correo)
            return
        if telefono and not telefono_ok(telefono):
            self.error("Teléfono: debe tener exactamente 10 dígitos.", self.telefono)
            return
        actual = self.clave_actual.text()
        nueva = self.clave_nueva.text()
        repetir = self.clave_repetir.text()
        cambia_clave = bool(actual or nueva or repetir)
        if cambia_clave:
            if not actual:
                self.error("Escriba su contraseña actual para cambiarla.", self.clave_actual)
                return
            try:
                if not verificar_clave(actual, self.claveHash):
                    self.error("La contraseña actual no es correcta.", self.clave_actual)
                    return
            except (ValueError, TypeError):
                self.error("No se pudo comprobar la contraseña actual.", self.clave_actual)
                return
            if len(nueva) < 8:
                self.error("La contraseña nueva debe tener al menos 8 caracteres.", self.clave_nueva)
                return
            if nueva != repetir:
                self.error("Las contraseñas nuevas no coinciden.", self.clave_repetir)
                return
        foto = self.foto_actual
        if self.foto_nueva:
            try:
                foto = copiar_foto(self.foto_nueva, self.usuario_inicial)
            except OSError:
                self.error("No se pudo guardar la foto. Revise los permisos de la carpeta.", None)
                return
        elif self.quitar_foto:
            foto = ""
        try:
            con = conectar()
        except sqlite3.Error:
            self.error(MSG_BD, None)
            return
        try:
            asegurar_esquema(con)
            with con:
                con.execute(
                    "UPDATE usuario SET nombreCompleto = ?, correo = ?, telefono = ?, foto = ? "
                    "WHERE idUsuario = ?;",
                    (nombre, correo or None, telefono or None, foto or None, self.idUsuario))
                if cambia_clave:
                    cambiar_clave_usuario(con, self.idUsuario, nueva)
                registrar_movimiento(con, "ACTUALIZAR CUENTA", "usuario", self.idUsuario,
                                     "Perfil y contraseña" if cambia_clave else "Perfil")
                fila_id = con.execute("SELECT idUsuario FROM usuario WHERE nombreUsuario = ?;",
                                      (self.usuario_inicial,)).fetchone()
                if fila_id is not None:
                    self.idUsuario = fila_id[0]
        except sqlite3.Error as e:
            fallo(self, "No se pudieron guardar los cambios de la cuenta.", SUGERENCIA_DB, str(e))
            return
        finally:
            con.close()
        SESION.update({"id": self.idUsuario, "nombreCompleto": nombre,
                       "correo": correo or None, "telefono": telefono or None,
                       "foto": foto or None})
        self.accept()


class DialogoNuevoUsuario(QDialog):
    def __init__(self, parent=None, titulo="Crear cuenta", roles=ROLES_AUTORREGISTRO, ayuda=""):
        super().__init__(parent)
        self.setWindowTitle(titulo)
        self.setFixedWidth(470)
        v = QVBoxLayout(self)
        v.setContentsMargins(26, 24, 26, 22)
        v.setSpacing(10)
        lbl_titulo = QLabel(titulo)
        lbl_titulo.setObjectName("dialogoTitulo")
        v.addWidget(lbl_titulo)
        if ayuda:
            lbl_ayuda = QLabel(ayuda)
            lbl_ayuda.setObjectName("ayuda")
            lbl_ayuda.setWordWrap(True)
            v.addWidget(lbl_ayuda)
        v.addWidget(ayuda_campo("Todos los campos son obligatorios. Siga las indicaciones "
                                "que aparecen debajo de cada uno."))
        self.nombre = QLineEdit()
        self.nombre.setPlaceholderText("Ej. Juan Pérez López")
        self.nombre.setToolTip("Solo letras y espacios. Ejemplo: Juan Pérez López")
        self.usuario = QLineEdit()
        self.usuario.setPlaceholderText("Ej. jperez")
        self.usuario.setToolTip("De 3 a 20 caracteres: letras, números o guion bajo (_).")
        self.rol = QComboBox()
        for codigo in roles:
            self.rol.addItem(ROLES_TEXTO.get(codigo, codigo.capitalize()), codigo)
        self.rol.setToolTip("El puesto define a qué secciones puede entrar la cuenta.")
        self.clave = QLineEdit()
        self.clave.setPlaceholderText("Mínimo 8 caracteres")
        self.clave.setEchoMode(QLineEdit.EchoMode.Password)
        self.clave.setToolTip("Al menos 8 caracteres. Use una contraseña que pueda recordar.")
        self.clave2 = QLineEdit()
        self.clave2.setPlaceholderText("Repita la contraseña")
        self.clave2.setEchoMode(QLineEdit.EchoMode.Password)
        self.clave2.setToolTip("Escriba exactamente la misma contraseña.")
        self.lbl_error = QTextEdit()
        self.lbl_error.setObjectName("errorBox")
        self.lbl_error.setReadOnly(True)
        self.lbl_error.setFixedHeight(96)
        self.lbl_error.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.lbl_error.hide()

        v.addWidget(etiqueta_campo("Nombre completo"))
        v.addWidget(self.nombre)
        v.addWidget(etiqueta_campo("Usuario"))
        v.addWidget(self.usuario)
        v.addWidget(ayuda_campo("De 3 a 20 caracteres: letras, números o guion bajo (_). "
                                "Ejemplo: jperez"))
        v.addWidget(etiqueta_campo("Puesto"))
        v.addWidget(self.rol)
        v.addWidget(etiqueta_campo("Contraseña"))
        v.addWidget(self.clave)
        v.addWidget(ayuda_campo("Mínimo 8 caracteres."))
        v.addWidget(etiqueta_campo("Repetir contraseña"))
        v.addWidget(self.clave2)
        v.addWidget(self.lbl_error)
        botones = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        botones.button(QDialogButtonBox.StandardButton.Save).setText("Crear cuenta")
        botones.button(QDialogButtonBox.StandardButton.Cancel).setText("Cancelar")
        botones.accepted.connect(self.aceptar)
        botones.rejected.connect(self.reject)
        v.addWidget(botones)
        self.nombre.setFocus()

    def valores(self):
        return {
            "nombre": self.nombre.text().strip(),
            "usuario": self.usuario.text().strip(),
            "rol": self.rol.currentData(),
            "clave": self.clave.text(),
            "clave2": self.clave2.text(),
        }

    def aceptar(self):
        datos = self.valores()
        campos = {"nombre": self.nombre, "usuario": self.usuario, "rol": self.rol,
                  "clave": self.clave, "clave2": self.clave2}
        for widget in campos.values():
            marcar_error(widget, False)
        errores = errores_usuario(datos)
        if errores:
            self.lbl_error.setPlainText("Corrija lo siguiente:\n" + texto_errores(errores))
            self.lbl_error.show()
            self.layout().activate()
            self.resize(self.width(), self.sizeHint().height() + 4)
            marcar_errores(errores, campos)
            return
        self.accept()


class DialogoClave(QDialog):
    def __init__(self, parent, usuario):
        super().__init__(parent)
        self.setWindowTitle("Restablecer contraseña")
        self.setFixedWidth(420)
        v = QVBoxLayout(self)
        v.setContentsMargins(26, 24, 26, 22)
        v.setSpacing(10)
        lbl_titulo = QLabel("Restablecer contraseña")
        lbl_titulo.setObjectName("dialogoTitulo")
        lbl_ayuda = QLabel(f"Cuenta: {usuario}. Escriba la nueva contraseña.")
        lbl_ayuda.setObjectName("ayuda")
        lbl_ayuda.setWordWrap(True)
        v.addWidget(lbl_titulo)
        v.addWidget(lbl_ayuda)
        v.addWidget(ayuda_campo("Use al menos 8 caracteres. La contraseña anterior dejará de funcionar."))
        self.clave = QLineEdit()
        self.clave.setPlaceholderText("Nueva contraseña (mínimo 8 caracteres)")
        self.clave.setEchoMode(QLineEdit.EchoMode.Password)
        self.clave2 = QLineEdit()
        self.clave2.setPlaceholderText("Repita la nueva contraseña")
        self.clave2.setEchoMode(QLineEdit.EchoMode.Password)
        self.lbl_error = QLabel("")
        self.lbl_error.setObjectName("error")
        self.lbl_error.setWordWrap(True)
        self.lbl_error.hide()
        v.addWidget(etiqueta_campo("Nueva contraseña"))
        v.addWidget(self.clave)
        v.addWidget(etiqueta_campo("Repetir contraseña"))
        v.addWidget(self.clave2)
        v.addWidget(self.lbl_error)
        botones = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        botones.button(QDialogButtonBox.StandardButton.Save).setText("Guardar contraseña")
        botones.button(QDialogButtonBox.StandardButton.Cancel).setText("Cancelar")
        botones.accepted.connect(self.aceptar)
        botones.rejected.connect(self.reject)
        v.addWidget(botones)
        self.clave.setFocus()

    def valores(self):
        return self.clave.text(), self.clave2.text()

    def aceptar(self):
        clave, clave2 = self.valores()
        marcar_error(self.clave, False)
        marcar_error(self.clave2, False)
        if len(clave) < 8:
            self.lbl_error.setText("La contraseña debe tener al menos 8 caracteres.")
            self.lbl_error.show()
            marcar_error(self.clave, True)
            return
        if clave != clave2:
            self.lbl_error.setText("Las contraseñas no coinciden. Escríbalas de nuevo.")
            self.lbl_error.show()
            marcar_error(self.clave2, True)
            return
        self.accept()


class DialogoAcceso(QDialog):
    def __init__(self, usuario_inicial=""):
        super().__init__()
        self.setWindowTitle("The Gonzalez - Acceso")
        self.setObjectName("dialogoAcceso")
        self.setFixedWidth(420)
        self.sesion = None
        self.intentos = 3
        exterior = QVBoxLayout(self)
        exterior.setContentsMargins(16, 16, 16, 16)
        tarjeta = QFrame()
        tarjeta.setObjectName("tarjetaAcceso")
        v = QVBoxLayout(tarjeta)
        v.setContentsMargins(30, 26, 30, 24)
        v.setSpacing(10)
        exterior.addWidget(tarjeta)
        monograma = QLabel("TG")
        monograma.setObjectName("marcaAcceso")
        monograma.setFixedSize(42, 42)
        monograma.setAlignment(Qt.AlignmentFlag.AlignCenter)
        v.addWidget(monograma, 0, Qt.AlignmentFlag.AlignHCenter)
        titulo = QLabel("The Gonzalez")
        titulo.setObjectName("loginTitulo")
        titulo.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        ayuda = QLabel("Gestión de servicios. Escriba su usuario y contraseña para entrar.")
        ayuda.setObjectName("loginAyuda")
        ayuda.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        ayuda.setWordWrap(True)
        self.usuario = QLineEdit()
        self.usuario.setPlaceholderText("Usuario")
        if usuario_inicial:
            self.usuario.setText(usuario_inicial)
        self.clave = QLineEdit()
        self.clave.setPlaceholderText("Contraseña")
        self.clave.setEchoMode(QLineEdit.EchoMode.Password)
        self.lbl_error = QLabel("")
        self.lbl_error.setObjectName("error")
        self.lbl_error.setWordWrap(True)
        self.lbl_error.hide()
        self.boton = QPushButton("Entrar")
        self.boton.setObjectName("primario")
        self.boton.setDefault(True)
        self.boton.setCursor(Qt.CursorShape.PointingHandCursor)
        self.boton.clicked.connect(self.intentar)
        self.clave.returnPressed.connect(self.intentar)
        v.addWidget(titulo)
        v.addWidget(ayuda)
        v.addSpacing(6)
        v.addWidget(etiqueta_campo("Usuario"))
        v.addWidget(self.usuario)
        v.addWidget(etiqueta_campo("Contraseña"))
        v.addWidget(self.clave)
        v.addWidget(self.lbl_error)
        v.addWidget(self.boton)
        self.boton_cuenta = QPushButton("Crear cuenta")
        self.boton_cuenta.setObjectName("enlace")
        self.boton_cuenta.setCursor(Qt.CursorShape.PointingHandCursor)
        self.boton_cuenta.setToolTip("Cree un usuario operativo para ingresar al sistema.")
        self.boton_cuenta.clicked.connect(self.crear_cuenta)
        v.addWidget(self.boton_cuenta)
        if usuario_inicial:
            self.clave.setFocus()
        else:
            self.usuario.setFocus()

    def _error(self, texto):
        self.lbl_error.setText(texto)
        self.lbl_error.show()

    def crear_cuenta(self):
        dlg = DialogoNuevoUsuario(
            self, "Crear cuenta", ROLES_AUTORREGISTRO,
            "La cuenta se crea con perfil de Cotizador. El gerente puede cambiar el puesto "
            "y desactivarla desde la sección Usuarios.")
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        datos = dlg.valores()
        try:
            con = conectar()
        except sqlite3.Error:
            self._error(MSG_BD)
            return
        try:
            if con.execute("SELECT 1 FROM usuario WHERE nombreUsuario = ?;",
                           (datos["usuario"],)).fetchone():
                self._error("Ese usuario ya existe. Elija otro.")
                return
            reiniciar_secuencia(con, "usuario")
            crear_usuario(con, datos["usuario"], datos["nombre"], datos["rol"], datos["clave"])
        except (sqlite3.Error, ValueError) as e:
            self._error(str(e) if isinstance(e, ValueError) and str(e)
                        else "No se pudo crear la cuenta. Intente de nuevo.")
            return
        finally:
            con.close()
        self.lbl_error.hide()
        self.usuario.setText(datos["usuario"])
        self.clave.clear()
        self.clave.setFocus()
        listo(self, f"La cuenta {datos['usuario']} se creó correctamente.\n"
                    "Escriba su contraseña para entrar.")

    def intentar(self):
        if not self.boton.isEnabled():
            return
        usuario = self.usuario.text().strip()
        clave = self.clave.text()
        if not usuario or not clave:
            self._error("Escriba su usuario y contraseña.")
            return
        self.boton.setEnabled(False)
        self.lbl_error.hide()
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            con = conectar()
            try:
                fila = con.execute(
                    """SELECT idUsuario, nombreUsuario, nombreCompleto, claveHash, rol
                       FROM usuario WHERE nombreUsuario = ? AND activo = 1;""", (usuario,)).fetchone()
            finally:
                con.close()
        except sqlite3.Error:
            self._error(MSG_BD)
            return
        finally:
            QApplication.restoreOverrideCursor()
            self.boton.setEnabled(True)
        valida = False
        if fila:
            try:
                valida = bool(verificar_clave(clave, fila["claveHash"]))
            except (ValueError, TypeError):
                valida = False
        if valida:
            self.sesion = {"id": fila["idUsuario"], "nombre": fila["nombreUsuario"],
                           "nombreCompleto": fila["nombreCompleto"] or fila["nombreUsuario"],
                           "rol": str(fila["rol"] or "").strip().upper()}
            self.accept()
            return
        self.intentos -= 1
        self.clave.clear()
        self.clave.setFocus()
        if self.intentos <= 0:
            self._error("Demasiados intentos fallidos. El programa se cerrará.")
            self.boton.setEnabled(False)
            QTimer.singleShot(2500, self.reject)
        elif self.intentos == 1:
            self._error("Usuario o contraseña incorrectos. Le queda 1 intento.")
        else:
            self._error(f"Usuario o contraseña incorrectos. Le quedan {self.intentos} intentos.")


class DialogoTexto(QDialog):
    def __init__(self, parent, titulo, etiqueta, ayuda="", multilinea=False, valor=""):
        super().__init__(parent)
        self.setWindowTitle(titulo)
        self.setMinimumWidth(450)
        v = QVBoxLayout(self)
        v.setContentsMargins(24, 22, 24, 20)
        v.setSpacing(10)
        lbl_titulo = QLabel(titulo)
        lbl_titulo.setObjectName("dialogoTitulo")
        v.addWidget(lbl_titulo)
        if ayuda:
            lbl_ayuda = QLabel(ayuda)
            lbl_ayuda.setObjectName("ayuda")
            lbl_ayuda.setWordWrap(True)
            v.addWidget(lbl_ayuda)
        v.addWidget(etiqueta_campo(etiqueta))
        if multilinea:
            self.campo = QTextEdit()
            self.campo.setFixedHeight(110)
            self.campo.setPlainText(valor)
        else:
            self.campo = QLineEdit()
            self.campo.setText(valor)
        v.addWidget(self.campo)
        botones = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        botones.button(QDialogButtonBox.StandardButton.Save).setText("Guardar")
        botones.button(QDialogButtonBox.StandardButton.Cancel).setText("Cancelar")
        botones.accepted.connect(self.aceptar)
        botones.rejected.connect(self.reject)
        v.addWidget(botones)
        self.campo.setFocus()

    def texto(self):
        if isinstance(self.campo, QTextEdit):
            return self.campo.toPlainText().strip()
        return self.campo.text().strip()

    def aceptar(self):
        if not self.texto():
            aviso(self, "Escriba el texto solicitado.")
            return
        self.accept()


class DialogoSeleccion(QDialog):
    def __init__(self, parent, titulo, etiqueta, opciones):
        super().__init__(parent)
        self.setWindowTitle(titulo)
        self.setMinimumWidth(460)
        v = QVBoxLayout(self)
        v.setContentsMargins(24, 22, 24, 20)
        v.setSpacing(10)
        lbl_titulo = QLabel(titulo)
        lbl_titulo.setObjectName("dialogoTitulo")
        v.addWidget(lbl_titulo)
        v.addWidget(etiqueta_campo(etiqueta))
        self.lista = QListWidget()
        self.lista.setMinimumHeight(170)
        for opcion in opciones:
            self.lista.addItem(opcion)
        if opciones:
            self.lista.setCurrentRow(0)
        self.lista.itemDoubleClicked.connect(lambda _item: self.aceptar())
        v.addWidget(self.lista)
        botones = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        botones.button(QDialogButtonBox.StandardButton.Ok).setText("Aceptar")
        botones.button(QDialogButtonBox.StandardButton.Cancel).setText("Cancelar")
        botones.accepted.connect(self.aceptar)
        botones.rejected.connect(self.reject)
        v.addWidget(botones)
        self.lista.setFocus()

    def indice(self):
        return self.lista.currentRow()

    def aceptar(self):
        if self.indice() < 0:
            aviso(self, "Seleccione una opción de la lista.")
            return
        self.accept()


AYUDA_SECCIONES = [
    ("Primeros pasos", """
<p style="color:#475569; font-size:11pt;">Configure y use el sistema en tres pasos.</p>
<table width="100%" cellspacing="0" cellpadding="12" style="background-color:#F7F9FC;">
<tr><td width="8%" style="color:#2F6FED; font-size:16pt; font-weight:700;">1</td>
<td><b>Cuenta inicial.</b> La primera vez, el programa pide crear la cuenta del gerente.
Esa cuenta administra el sistema y crea las demás.</td></tr>
<tr><td style="color:#2F6FED; font-size:16pt; font-weight:700;">2</td>
<td><b>Acceso.</b> Escriba su usuario y contraseña. También puede crear una cuenta
operativa con el enlace <b>Crear cuenta</b> de la pantalla de acceso.</td></tr>
<tr><td style="color:#2F6FED; font-size:16pt; font-weight:700;">3</td>
<td><b>Trabajo diario.</b> Registre primero al cliente, luego al empleado y por último
el servicio. El servicio avanza por sus estados hasta finalizar.</td></tr>
</table>

<h3 style="color:#111C2E; font-size:12pt;">Secciones del programa</h3>
<table width="100%" cellspacing="0" cellpadding="8">
<tr><td width="26%"><b>Resumen</b></td><td>Indicadores y últimos servicios registrados.</td></tr>
<tr><td><b>Clientes</b></td><td>Alta, búsqueda, cambios y baja de clientes.</td></tr>
<tr><td><b>Empleados</b></td><td>Personal, teléfonos y puestos de trabajo.</td></tr>
<tr><td><b>Servicios</b></td><td>Registro y seguimiento del trabajo, cotización y facturación.</td></tr>
<tr><td><b>Usuarios</b></td><td>Cuentas y permisos. Solo la ve el gerente.</td></tr>
</table>

<table width="100%" cellspacing="0" cellpadding="10" style="background-color:#EEF4FF;">
<tr><td><b>Consejo:</b> si un formulario no se guarda, revise los campos marcados en rojo.
El mensaje indica la regla y un ejemplo.</td></tr>
</table>

<h3 style="color:#111C2E; font-size:12pt;">Avisos y cierre de sesión</h3>
<ul>
<li><b>Avisos:</b> muestra los mensajes del administrador, cambios de puesto y
actualizaciones. El número indica los que no ha leído.</li>
<li><b>Cerrar sesión:</b> marca su cuenta como desconectada y cierra el programa.
Así el administrador puede cambiar su puesto sin conflictos.</li>
<li>Al volver a entrar, el nuevo puesto se aplica y verá los avisos pendientes.</li>
</ul>
"""),
    ("Mi cuenta", """
<p style="color:#475569; font-size:11pt;">Desde el botón <b>Mi cuenta</b> puede actualizar sus
datos personales, su foto y su contraseña.</p>
<table width="100%" cellspacing="0" cellpadding="7" style="border:1px solid #E3E8F0;">
<tr style="background-color:#F7F9FC;">
<td width="24%"><b>Campo</b></td><td><b>Cómo funciona</b></td></tr>
<tr><td><b>Foto</b></td><td>Opcional. Elija una imagen desde su equipo; se guarda en la carpeta
<b>perfiles</b> del programa. Puede quitarla cuando quiera.</td></tr>
<tr><td><b>Nombre</b></td><td>Solo letras y espacios.</td></tr>
<tr><td><b>Correo</b></td><td>Formato nombre@dominio.com. Se usa para avisos y solicitudes.</td></tr>
<tr><td><b>Teléfono</b></td><td>10 dígitos, sin espacios ni guiones.</td></tr>
<tr><td><b>Contraseña</b></td><td>Para cambiarla escriba la actual y la nueva dos veces.
Mínimo 8 caracteres. Si no desea cambiarla, deje esos campos vacíos.</td></tr>
</table>

<h3 style="color:#111C2E; font-size:12pt;">Solicitar cambio de puesto</h3>
<p>Use <b>Solicitar cambio de puesto</b>, elija el puesto y escriba el motivo. La solicitud
se guarda y se envía por Gmail al administrador. El administrador la aprueba o rechaza desde el
<b>Panel de administración</b>.</p>
<table width="100%" cellspacing="0" cellpadding="10" style="background-color:#EEF4FF;">
<tr><td><b>Gmail:</b> para enviar el aviso configure una cuenta y una contraseña de
aplicación con el botón <b>Configurar Gmail</b>. El correo se envía al correo del
administrador del panel. Si no la configura, la solicitud igual queda registrada y
el administrador la verá.</td></tr>
</table>
"""),
    ("Clientes", """
<p style="color:#475569; font-size:11pt;">El nombre es el único dato obligatorio.
Los demás campos se pueden dejar vacíos.</p>
<table width="100%" cellspacing="0" cellpadding="7" style="border:1px solid #E3E8F0;">
<tr style="background-color:#F7F9FC;">
<td width="17%"><b>Campo</b></td><td width="53%"><b>Cómo llenarlo</b></td><td width="30%"><b>Ejemplo</b></td></tr>
<tr><td><b>Nombre</b></td><td>Solo letras y espacios. Sin números ni símbolos.</td><td>Juan Rolquez</td></tr>
<tr><td><b>Teléfono</b></td><td>Exactamente 10 dígitos, sin espacios ni guiones.</td><td>8441234567</td></tr>
<tr><td><b>Correo</b></td><td>Formato nombre@dominio.com</td><td>juan@correo.com</td></tr>
<tr><td><b>Dirección</b></td><td>Opcional. Calle, número y colonia.</td><td>Av. Juárez 120</td></tr>
<tr><td><b>RFC</b></td><td>Opcional. 12 o 13 caracteres: 3 o 4 letras, 6 dígitos de fecha y 3 de homoclave.</td><td>RUIA010101AB1</td></tr>
</table>

<h3 style="color:#111C2E; font-size:12pt;">Antes de guardar</h3>
<ul>
<li>Verifique que el cliente no esté ya en la lista; use <b>Buscar</b>.</li>
<li>Si el nombre se repite, el programa preguntará si desea registrarlo de todos modos.</li>
<li>Si aparece un campo en rojo, corrija ese dato: el mensaje explica la regla.</li>
</ul>

<h3 style="color:#111C2E; font-size:12pt;">Dar de baja</h3>
<p>Seleccione el cliente y use <b>Dar de baja</b>. No se permite si tiene servicios
en proceso; primero finalice o cancele esos servicios. El historial se conserva.</p>

<h3 style="color:#111C2E; font-size:12pt;">Números e historial</h3>
<p>Dar de baja oculta al cliente de la lista, pero conserva su historial y su número
para que las facturas y servicios anteriores sigan siendo correctos. Un cliente nuevo
recibe el siguiente número; el número no se reutiliza mientras el registro exista.
Si la tabla queda completamente vacía, la numeración vuelve a empezar en 1.</p>
"""),
    ("Empleados", """
<p style="color:#475569; font-size:11pt;">Registre al personal y asígnele uno o más puestos.</p>
<table width="100%" cellspacing="0" cellpadding="7" style="border:1px solid #E3E8F0;">
<tr style="background-color:#F7F9FC;">
<td width="17%"><b>Campo</b></td><td width="53%"><b>Cómo llenarlo</b></td><td width="30%"><b>Ejemplo</b></td></tr>
<tr><td><b>Nombre</b></td><td>Obligatorio. Solo letras y espacios.</td><td>Juan</td></tr>
<tr><td><b>Apellido paterno</b></td><td>Obligatorio. Solo letras y espacios.</td><td>Pérez</td></tr>
<tr><td><b>Apellido materno</b></td><td>Opcional. Solo letras y espacios.</td><td>López</td></tr>
<tr><td><b>Teléfono</b></td><td>10 dígitos, sin espacios ni guiones.</td><td>8441234567</td></tr>
<tr><td><b>Correo</b></td><td>Formato nombre@dominio.com</td><td>juan@correo.com</td></tr>
<tr><td><b>Puestos</b></td><td>Seleccione uno o más con el mouse.</td><td>Técnico</td></tr>
</table>
<table width="100%" cellspacing="0" cellpadding="10" style="background-color:#EEF4FF;">
<tr><td><b>Importante:</b> no se puede dar de baja a un empleado que está asignado
a un servicio en proceso.</td></tr>
</table>

<h3 style="color:#111C2E; font-size:12pt;">Números e historial</h3>
<p>La baja conserva el historial del empleado y su número. El siguiente empleado
recibe el número siguiente; no se reutiliza mientras el registro exista.</p>
"""),
    ("Servicios", """
<p style="color:#475569; font-size:11pt;">Registre el trabajo solicitado y acompáñelo
por sus estados hasta la facturación.</p>
<table width="100%" cellspacing="0" cellpadding="7" style="border:1px solid #E3E8F0;">
<tr style="background-color:#F7F9FC;">
<td width="17%"><b>Campo</b></td><td width="53%"><b>Cómo llenarlo</b></td><td width="30%"><b>Ejemplo</b></td></tr>
<tr><td><b>Cliente</b></td><td>Escriba parte del nombre y elija el cliente correcto de la lista. Si no aparece, regístrelo antes.</td><td>Juan Rolquez</td></tr>
<tr><td><b>Problema</b></td><td>Obligatorio. Al menos 3 letras.</td><td>Falla la bomba de agua</td></tr>
<tr><td><b>Descripción</b></td><td>Opcional. Detalle del trabajo.</td><td>Cambio de sello</td></tr>
<tr><td><b>Observaciones</b></td><td>Opcional. Notas adicionales.</td><td>Revisar presión</td></tr>
</table>

<h3 style="color:#111C2E; font-size:12pt;">Estados del servicio</h3>
<p>Registrado, Asignado, Cotización generada, Aceptado, En proceso, Pendiente de revisión,
Corrección solicitada, Aprobado, Facturado y Finalizado.</p>
<p>Las acciones disponibles aparecen en el botón <b>Acciones</b> según el estado y su puesto.</p>

<h3 style="color:#111C2E; font-size:12pt;">Cotizaciones</h3>
<ul>
<li>La cotización muestra <b>servicio, cliente y problema</b> antes de capturar los conceptos.</li>
<li>Cada concepto lleva <b>unidad de medida</b>: SERVICIO, PZA, HORA, DÍA, KG, M2, LOTE, etc.</li>
<li>Las cotizaciones nacen como <b>Registrada</b>. Las antiguas “Pendiente” se convierten
solas al abrir el programa; no se borra ninguna.</li>
<li>Mientras el servicio esté en <b>Cotización generada</b> puede usar <b>Volver a cotizar</b>.
Las cotizaciones anteriores se conservan.</li>
<li>El Excel es formal: título, datos del cliente y servicio, tabla de conceptos,
subtotal, IVA, total e <b>importe con letra</b>.</li>
<li>Al facturar se valida que exista una cotización <b>aceptada</b> y se usa su importe.</li>
</ul>

<table width="100%" cellspacing="0" cellpadding="10" style="background-color:#EEF4FF;">
<tr><td><b>Consejo:</b> antes de facturar, confirme que el servicio esté aprobado en revisión.</td></tr>
</table>
"""),
    ("Usuarios", """
<p style="color:#475569; font-size:11pt;">Solo el gerente ve esta sección. Aquí se crean,
activan, desactivan y corrigen las cuentas de acceso.</p>
<table width="100%" cellspacing="0" cellpadding="7" style="border:1px solid #E3E8F0;">
<tr style="background-color:#F7F9FC;">
<td width="17%"><b>Campo</b></td><td width="53%"><b>Cómo llenarlo</b></td><td width="30%"><b>Ejemplo</b></td></tr>
<tr><td><b>Nombre</b></td><td>Solo letras y espacios.</td><td>Juan Pérez</td></tr>
<tr><td><b>Usuario</b></td><td>De 3 a 20 caracteres: letras, números o guion bajo.</td><td>jperez</td></tr>
<tr><td><b>Contraseña</b></td><td>Mínimo 8 caracteres. Debe repetirse igual.</td><td>Clave segura</td></tr>
<tr><td><b>Puesto</b></td><td>Define los permisos de la cuenta.</td><td>Cotizador</td></tr>
</table>

<h3 style="color:#111C2E; font-size:12pt;">Puestos y permisos</h3>
<table width="100%" cellspacing="0" cellpadding="8">
<tr><td width="22%"><b>Gerente</b></td><td>Administra todo: clientes, empleados, servicios y usuarios.</td></tr>
<tr><td><b>Cotizador</b></td><td>Clientes, servicios y cotizaciones.</td></tr>
<tr><td><b>Supervisor</b></td><td>Revisión de trabajos y comentarios al personal.</td></tr>
<tr><td><b>Contador</b></td><td>Facturación y finalización de servicios.</td></tr>
</table>

<h3 style="color:#111C2E; font-size:12pt;">Seguridad</h3>
<ul>
<li>No comparta su contraseña ni use la cuenta de otra persona.</li>
<li>La cuenta con la que está trabajando no se puede desactivar.</li>
<li>Debe quedar al menos un gerente activo.</li>
<li>Una cuenta desactivada ya no puede entrar al sistema.</li>
</ul>

<table width="100%" cellspacing="0" cellpadding="10" style="background-color:#EEF4FF;">
<tr><td><b>Panel de administración:</b> tiene una cuenta de administrador propia,
independiente de los usuarios de este programa. No use aquí su contraseña de
Gerente para ese panel.</td></tr>
</table>

<h3 style="color:#111C2E; font-size:12pt;">Números e historial</h3>
<ul>
<li><b>Desactivar</b> no borra la cuenta: conserva el historial y su número.</li>
<li>Para liberar el número, primero desactive la cuenta y luego use
<b>Eliminar</b>.</li>
<li>Si la tabla de usuarios queda vacía, la numeración vuelve a empezar en 1.</li>
<li>El historial de acciones de la bitácora se conserva aunque se elimine la cuenta.</li>
</ul>

<table width="100%" cellspacing="0" cellpadding="10" style="background-color:#FEF3F2;">
<tr><td><b>Eliminar definitivamente</b> no se puede deshacer. Úselo solo con cuentas
inactivas y cuando esté seguro.</td></tr>
</table>
"""),
    ("Errores frecuentes", """
<p style="color:#475569; font-size:11pt;">Todos los mensajes están en lenguaje sencillo.
Si necesita más detalle, use el botón <b>Posible causa</b>.</p>
<table width="100%" cellspacing="0" cellpadding="8" style="border:1px solid #E3E8F0;">
<tr style="background-color:#F7F9FC;">
<td width="38%"><b>Mensaje</b></td><td><b>Qué hacer</b></td></tr>
<tr><td><b>No se pudo completar la operación</b></td><td>Puede ser un problema de permisos,
de espacio en disco o de la base de datos. Cierre otras ventanas del programa y vuelva a
intentar. Si continúa, avise al administrador.</td></tr>
<tr><td><b>La base de datos está en uso</b></td><td>Otra ventana del programa la está usando.
Cierre las demás ventanas y espere unos segundos.</td></tr>
<tr><td><b>La base de datos es de solo lectura</b></td><td>Mueva el programa y la base de datos
a una carpeta donde tenga permiso de escritura (por ejemplo, Documentos).</td></tr>
<tr><td><b>Un dato ya existe</b></td><td>Use otro nombre de usuario o revise si el cliente
ya está registrado.</td></tr>
<tr><td><b>Corrija los siguientes datos</b></td><td>Los campos en rojo tienen un error.
Lea la regla y el ejemplo que aparecen en el mensaje.</td></tr>
<tr><td><b>Usuario o contraseña incorrectos</b></td><td>Verifique mayúsculas y minúsculas.
Después de tres intentos el programa se cierra por seguridad.</td></tr>
<tr><td><b>No se encontró el archivo Excel</b></td><td>El archivo se movió o se borró.
Genere la cotización de nuevo.</td></tr>
<tr><td><b>El archivo cambió después de generarse</b></td><td>El Excel se modificó.
Avise al gerente antes de usarlo.</td></tr>
<tr><td><b>Su cuenta no tiene permiso para esta acción</b></td><td>Su cuenta no
tiene permisos para esa sección. Pida al gerente que cambie su puesto en Usuarios.</td></tr>
</table>

<table width="100%" cellspacing="0" cellpadding="10" style="background-color:#EEF4FF;">
<tr><td><b>Detalle técnico:</b> se guarda en el archivo <b>errores.log</b>, junto al programa,
para que el administrador pueda revisarlo. No se muestra en pantalla.</td></tr>
</table>
"""),
    ("Diagnóstico", None),
    ("Recomendaciones", """
<p style="color:#475569; font-size:11pt;">Buenas prácticas para mantener la información
ordenada y segura.</p>
<ul>
<li>Registre primero al cliente y luego el servicio.</li>
<li>Use el botón <b>Buscar</b> antes de crear un registro para evitar duplicados.</li>
<li>Revise el estado del servicio antes de facturar o finalizar.</li>
<li>Cierre el programa al terminar; el respaldo se hace automáticamente al iniciar.</li>
<li>No abra el archivo de la base de datos con otros programas.</li>
<li>Mantenga una copia de seguridad de la carpeta del programa.</li>
<li>Asigne a cada persona solo el puesto que necesita.</li>
</ul>

<table width="100%" cellspacing="0" cellpadding="10" style="background-color:#F7F9FC;">
<tr><td><b>Atajo:</b> presione <b>F1</b> para abrir esta ayuda desde cualquier sección.</td></tr>
</table>
"""),
]


class DialogoAyuda(QDialog):
    def __init__(self, parent=None, seccion=0):
        super().__init__(parent)
        self.setWindowTitle("Centro de ayuda - The Gonzalez")
        self.resize(940, 660)
        raiz = QHBoxLayout(self)
        raiz.setContentsMargins(0, 0, 0, 0)
        raiz.setSpacing(0)

        lateral = QFrame()
        lateral.setObjectName("ayudaLateral")
        lateral.setFixedWidth(248)
        vl = QVBoxLayout(lateral)
        vl.setContentsMargins(18, 24, 18, 18)
        vl.setSpacing(4)
        titulo = QLabel("Centro de ayuda")
        titulo.setObjectName("ayudaTitulo")
        subtitulo = QLabel("The Gonzalez")
        subtitulo.setObjectName("ayudaSub")
        vl.addWidget(titulo)
        vl.addWidget(subtitulo)
        vl.addSpacing(16)
        self.lista = QListWidget()
        self.lista.setObjectName("ayudaNav")
        self.lista.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        for nombre, _html in AYUDA_SECCIONES:
            self.lista.addItem(nombre)
        vl.addWidget(self.lista, 1)
        pie = QLabel(f"Versión {VERSION}. Ayuda con F1.")
        pie.setObjectName("ayudaPie")
        pie.setWordWrap(True)
        vl.addWidget(pie)
        raiz.addWidget(lateral)

        contenido = QWidget()
        vc = QVBoxLayout(contenido)
        vc.setContentsMargins(30, 24, 30, 18)
        vc.setSpacing(12)
        self.lbl_seccion = QLabel("")
        self.lbl_seccion.setObjectName("ayudaSeccionTitulo")
        self.texto = QTextBrowser()
        self.texto.setObjectName("ayudaTexto")
        self.texto.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.texto.setOpenExternalLinks(False)
        vc.addWidget(self.lbl_seccion)
        vc.addWidget(self.texto, 1)
        botones = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        boton_cerrar = botones.button(QDialogButtonBox.StandardButton.Close)
        boton_cerrar.setText("Cerrar")
        boton_cerrar.setObjectName("secundario")
        botones.rejected.connect(self.accept)
        vc.addWidget(botones)
        raiz.addWidget(contenido, 1)

        self.lista.currentRowChanged.connect(self.mostrar_seccion)
        indice = max(0, min(int(seccion), len(AYUDA_SECCIONES) - 1))
        self.lista.setCurrentRow(indice)

    def mostrar_seccion(self, indice):
        if indice < 0 or indice >= len(AYUDA_SECCIONES):
            return
        titulo, html = AYUDA_SECCIONES[indice]
        if html is None:
            html = diagnostico_html()
        self.lbl_seccion.setText(titulo)
        self.texto.setHtml(html)
        self.texto.verticalScrollBar().setValue(0)


class DialogoMensaje(QDialog):
    def __init__(self, parent, titulo, texto, detalle=None):
        super().__init__(parent)
        self.setWindowTitle(titulo)
        self.setMinimumWidth(480)
        self.detalle = detalle
        v = QVBoxLayout(self)
        v.setContentsMargins(24, 22, 24, 20)
        v.setSpacing(12)
        lbl_titulo = QLabel(titulo)
        lbl_titulo.setObjectName("dialogoTitulo")
        v.addWidget(lbl_titulo)
        cuerpo = QLabel(texto)
        cuerpo.setWordWrap(True)
        cuerpo.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        v.addWidget(cuerpo)
        botones = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        boton_cerrar = botones.button(QDialogButtonBox.StandardButton.Close)
        boton_cerrar.setText("Cerrar")
        botones.rejected.connect(self.accept)
        if detalle:
            self.boton_copiar = botones.addButton("Copiar detalle", QDialogButtonBox.ButtonRole.ActionRole)
            self.boton_copiar.clicked.connect(self.copiar_detalle)
        v.addWidget(botones)

    def copiar_detalle(self):
        QApplication.clipboard().setText(self.detalle)
        self.boton_copiar.setText("Detalle copiado")


class DialogoCliente(QDialog):
    def __init__(self, parent, titulo, datos=None):
        super().__init__(parent)
        self.setWindowTitle(titulo)
        self.setMinimumWidth(430)
        v = QVBoxLayout(self)
        v.setContentsMargins(24, 22, 24, 20)
        v.setSpacing(12)
        lbl_titulo = QLabel(titulo)
        lbl_titulo.setObjectName("dialogoTitulo")
        v.addWidget(lbl_titulo)
        v.addWidget(ayuda_campo("Nombre obligatorio. Teléfono de 10 dígitos. RFC de 12 o 13 "
                                "caracteres. Correo y dirección son opcionales."))
        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.AllNonFixedFieldsGrow)
        form.setHorizontalSpacing(16)
        form.setVerticalSpacing(10)
        v.addLayout(form)
        self.nombre = QLineEdit()
        self.nombre.setPlaceholderText("Ej. Pedro Ramírez")
        self.nombre.setToolTip("Solo letras y espacios.")
        self.telefono = QLineEdit()
        self.telefono.setPlaceholderText("Ej. 8441234567")
        self.telefono.setToolTip("Exactamente 10 dígitos, sin espacios ni guiones.")
        self.correo = QLineEdit()
        self.correo.setPlaceholderText("Ej. cliente@correo.com")
        self.correo.setToolTip("Formato nombre@dominio.com")
        self.direccion = QLineEdit()
        self.direccion.setPlaceholderText("Ej. Av. Juárez 120")
        self.direccion.setToolTip("Calle, número y colonia. Opcional.")
        self.rfc = QLineEdit()
        self.rfc.setPlaceholderText("Ej. RUIA010101AB1")
        self.rfc.setToolTip("12 o 13 caracteres: letras, fecha y homoclave. Opcional.")
        if datos:
            self.nombre.setText(datos["nombre"] or "")
            self.telefono.setText(datos["telefono"] or "")
            self.correo.setText(datos["correo"] or "")
            self.direccion.setText(datos["direccion"] or "")
            self.rfc.setText(datos["rfc"] or "")
        form.addRow(etiqueta_campo("Nombre"), self.nombre)
        form.addRow(etiqueta_campo("Teléfono"), self.telefono)
        form.addRow(etiqueta_campo("Correo"), self.correo)
        form.addRow(etiqueta_campo("Dirección"), self.direccion)
        form.addRow(etiqueta_campo("RFC"), self.rfc)
        botones = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        botones.button(QDialogButtonBox.StandardButton.Save).setText("Guardar")
        botones.button(QDialogButtonBox.StandardButton.Cancel).setText("Cancelar")
        botones.accepted.connect(self.aceptar)
        botones.rejected.connect(self.reject)
        v.addWidget(botones)

    def valores(self):
        return {
            "nombre": self.nombre.text().strip(),
            "telefono": self.telefono.text().strip() or None,
            "correo": self.correo.text().strip() or None,
            "direccion": self.direccion.text().strip() or None,
            "rfc": self.rfc.text().strip().upper() or None,
        }

    def aceptar(self):
        datos = self.valores()
        campos = {"nombre": self.nombre, "telefono": self.telefono, "correo": self.correo,
                  "direccion": self.direccion, "rfc": self.rfc}
        for widget in campos.values():
            marcar_error(widget, False)
        errores = errores_cliente(datos)
        if errores:
            aviso(self, "Corrija los siguientes datos:\n\n" + texto_errores(errores))
            marcar_errores(errores, campos)
            return
        self.accept()


class DialogoEmpleado(QDialog):
    def __init__(self, parent, titulo, datos=None, roles_actuales=()):
        super().__init__(parent)
        self.setWindowTitle(titulo)
        self.setMinimumWidth(430)
        v = QVBoxLayout(self)
        v.setContentsMargins(24, 22, 24, 20)
        v.setSpacing(12)
        lbl_titulo = QLabel(titulo)
        lbl_titulo.setObjectName("dialogoTitulo")
        v.addWidget(lbl_titulo)
        v.addWidget(ayuda_campo("Nombre y apellido paterno obligatorios. Teléfono de 10 dígitos. "
                                "Seleccione al menos un puesto."))
        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.AllNonFixedFieldsGrow)
        form.setHorizontalSpacing(16)
        form.setVerticalSpacing(10)
        v.addLayout(form)
        self.nombre = QLineEdit()
        self.nombre.setPlaceholderText("Ej. Juan")
        self.nombre.setToolTip("Solo letras y espacios.")
        self.ap = QLineEdit()
        self.ap.setPlaceholderText("Ej. Pérez")
        self.ap.setToolTip("Solo letras y espacios. Obligatorio.")
        self.am = QLineEdit()
        self.am.setPlaceholderText("Ej. López")
        self.am.setToolTip("Solo letras y espacios. Opcional.")
        self.telefono = QLineEdit()
        self.telefono.setPlaceholderText("Ej. 8441234567")
        self.telefono.setToolTip("Exactamente 10 dígitos, sin espacios ni guiones.")
        self.correo = QLineEdit()
        self.correo.setPlaceholderText("Ej. juan@correo.com")
        self.correo.setToolTip("Formato nombre@dominio.com")
        self.roles = QListWidget()
        self.roles.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        self.roles.setMaximumHeight(160)
        self.roles.setToolTip("Marque uno o más puestos con el mouse.")
        for idRol, nombreRol in ROLES_LISTA:
            item = QListWidgetItem(ROLES_TEXTO.get(nombreRol, nombreRol.capitalize()))
            item.setData(Qt.ItemDataRole.UserRole, idRol)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Checked if idRol in roles_actuales
                               else Qt.CheckState.Unchecked)
            self.roles.addItem(item)
        if datos:
            self.nombre.setText(datos["nombre"] or "")
            self.ap.setText(datos["aPaterno"] or "")
            self.am.setText(datos["aMaterno"] or "")
            self.telefono.setText(datos["telefono"] or "")
            self.correo.setText(datos["correo"] or "")
        form.addRow(etiqueta_campo("Nombre"), self.nombre)
        form.addRow(etiqueta_campo("Apellido paterno"), self.ap)
        form.addRow(etiqueta_campo("Apellido materno"), self.am)
        form.addRow(etiqueta_campo("Teléfono"), self.telefono)
        form.addRow(etiqueta_campo("Correo"), self.correo)
        form.addRow(etiqueta_campo("Puestos"), self.roles)
        botones = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        botones.button(QDialogButtonBox.StandardButton.Save).setText("Guardar")
        botones.button(QDialogButtonBox.StandardButton.Cancel).setText("Cancelar")
        botones.accepted.connect(self.aceptar)
        botones.rejected.connect(self.reject)
        v.addWidget(botones)

    def valores(self):
        return {
            "nombre": self.nombre.text().strip(),
            "ap": self.ap.text().strip(),
            "am": self.am.text().strip() or None,
            "telefono": self.telefono.text().strip() or None,
            "correo": self.correo.text().strip() or None,
            "roles": [self.roles.item(i).data(Qt.ItemDataRole.UserRole)
                      for i in range(self.roles.count())
                      if self.roles.item(i).checkState() == Qt.CheckState.Checked
                      or self.roles.item(i).isSelected()],
        }

    def aceptar(self):
        datos = self.valores()
        campos = {"nombre": self.nombre, "ap": self.ap, "am": self.am,
                  "telefono": self.telefono, "correo": self.correo, "roles": self.roles}
        for widget in campos.values():
            marcar_error(widget, False)
        errores = errores_empleado(datos)
        if errores:
            aviso(self, "Corrija los siguientes datos:\n\n" + texto_errores(errores))
            marcar_errores(errores, campos)
            return
        self.accept()


class DialogoServicio(QDialog):
    def __init__(self, parent, datos):
        super().__init__(parent)
        self.setWindowTitle("Modificar servicio")
        self.setMinimumWidth(470)
        v = QVBoxLayout(self)
        v.setContentsMargins(24, 22, 24, 20)
        v.setSpacing(12)
        lbl_titulo = QLabel("Modificar servicio")
        lbl_titulo.setObjectName("dialogoTitulo")
        v.addWidget(lbl_titulo)
        v.addWidget(ayuda_campo("El problema es obligatorio: describa con al menos 3 letras qué "
                                "falla o qué se requiere."))
        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.AllNonFixedFieldsGrow)
        form.setHorizontalSpacing(16)
        form.setVerticalSpacing(10)
        v.addLayout(form)
        self.problema = QTextEdit(datos["descripcionProblema"] or "")
        self.problema.setToolTip("Obligatorio. Ejemplo: Falla la bomba de agua.")
        self.descripcion = QTextEdit(datos["descripcionServicio"] or "")
        self.descripcion.setToolTip("Detalle del trabajo a realizar. Opcional.")
        self.observaciones = QTextEdit(datos["observaciones"] or "")
        self.observaciones.setToolTip("Notas adicionales. Opcional.")
        for w in (self.problema, self.descripcion, self.observaciones):
            w.setMaximumHeight(90)
        form.addRow(etiqueta_campo("Problema"), self.problema)
        form.addRow(etiqueta_campo("Descripción del servicio"), self.descripcion)
        form.addRow(etiqueta_campo("Observaciones"), self.observaciones)
        botones = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        botones.button(QDialogButtonBox.StandardButton.Save).setText("Guardar")
        botones.button(QDialogButtonBox.StandardButton.Cancel).setText("Cancelar")
        botones.accepted.connect(self.aceptar)
        botones.rejected.connect(self.reject)
        v.addWidget(botones)

    def valores(self):
        return {
            "problema": self.problema.toPlainText().strip(),
            "descripcion": self.descripcion.toPlainText().strip() or None,
            "observaciones": self.observaciones.toPlainText().strip() or None,
        }

    def aceptar(self):
        datos = self.valores()
        campos = {"problema": self.problema}
        marcar_error(self.problema, False)
        errores = errores_servicio(datos)
        if errores:
            aviso(self, texto_errores(errores))
            marcar_errores(errores, campos)
            return
        self.accept()


class PanelContenedor(QFrame):
    def __init__(self, titulo, ayuda=""):
        super().__init__()
        self.setObjectName("card")
        self.layout_principal = QVBoxLayout(self)
        self.layout_principal.setContentsMargins(22, 20, 22, 20)
        self.layout_principal.setSpacing(10)
        self.lbl_titulo = QLabel(titulo)
        self.lbl_titulo.setObjectName("seccion")
        self.layout_principal.addWidget(self.lbl_titulo)
        if ayuda:
            lbl_ayuda = QLabel(ayuda)
            lbl_ayuda.setObjectName("ayuda")
            lbl_ayuda.setWordWrap(True)
            self.layout_principal.addWidget(lbl_ayuda)


def iniciales(nombre):
    partes = [p for p in str(nombre or "").split() if p]
    if not partes:
        return "US"
    if len(partes) == 1:
        return partes[0][:2].upper()
    return (partes[0][0] + partes[-1][0]).upper()


def _pluma_icono(p, color):
    pluma = QPen(QColor(color))
    pluma.setWidthF(1.7)
    pluma.setCapStyle(Qt.PenCapStyle.RoundCap)
    pluma.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
    p.setPen(pluma)
    p.setBrush(Qt.BrushStyle.NoBrush)


def icono_nav(tipo, color="#B9C6D8"):
    pixmap = QPixmap(20, 20)
    pixmap.fill(Qt.GlobalColor.transparent)
    p = QPainter(pixmap)
    p.setRenderHint(QPainter.RenderHint.Antialiasing)
    _pluma_icono(p, color)
    if tipo == "resumen":
        for x, y in ((2, 2), (12, 2), (2, 12), (12, 12)):
            p.drawRoundedRect(x, y, 6, 6, 2, 2)
    elif tipo == "clientes":
        p.drawEllipse(7, 2, 6, 6)
        p.drawArc(3, 10, 14, 10, 0, 180 * 16)
    elif tipo == "empleados":
        p.drawEllipse(2, 4, 5, 5)
        p.drawArc(0, 10, 9, 8, 0, 180 * 16)
        p.drawEllipse(11, 4, 5, 5)
        p.drawArc(9, 10, 9, 8, 0, 180 * 16)
    elif tipo == "servicios":
        p.drawRoundedRect(4, 3, 12, 15, 2, 2)
        p.drawRoundedRect(7, 1, 6, 4, 1, 1)
        for y in (9, 12, 15):
            p.drawLine(7, y, 13, y)
    elif tipo == "usuarios":
        ruta = QPainterPath()
        ruta.moveTo(10, 2)
        ruta.lineTo(17, 5)
        ruta.lineTo(17, 10)
        ruta.cubicTo(17, 14, 10, 18, 10, 18)
        ruta.cubicTo(10, 18, 3, 14, 3, 10)
        ruta.lineTo(3, 5)
        ruta.closeSubpath()
        p.drawPath(ruta)
    elif tipo == "movimientos":
        p.drawRoundedRect(3, 3, 14, 15, 2, 2)
        for y in (7, 10, 13, 16):
            p.drawLine(6, y, 10, y)
            p.drawLine(12, y, 15, y)
    elif tipo == "facturas":
        p.drawRoundedRect(3, 2, 11, 16, 2, 2)
        p.drawLine(6, 6, 11, 6)
        p.drawLine(6, 9, 11, 9)
        p.drawLine(6, 12, 9, 12)
        p.drawText(QRectF(10, 9, 8, 9), Qt.AlignmentFlag.AlignCenter, "$")
    p.end()
    return QIcon(pixmap)


def icono_buscar(color="#94A3B8"):
    pixmap = QPixmap(18, 18)
    pixmap.fill(Qt.GlobalColor.transparent)
    p = QPainter(pixmap)
    p.setRenderHint(QPainter.RenderHint.Antialiasing)
    _pluma_icono(p, color)
    p.drawEllipse(3, 3, 9, 9)
    p.drawLine(12, 12, 16, 16)
    p.end()
    return QIcon(pixmap)


def tarjeta_indicador(titulo, tono="azul"):
    tarjeta = QFrame()
    tarjeta.setObjectName("kpiCard")
    tarjeta.setProperty("tono", tono)
    tarjeta.setMinimumWidth(146)
    v = QVBoxLayout(tarjeta)
    v.setContentsMargins(16, 14, 16, 14)
    v.setSpacing(6)
    lbl_t = QLabel(titulo)
    lbl_t.setObjectName("kpiTitulo")
    lbl_t.setWordWrap(True)
    lbl_t.setMinimumHeight(32)
    lbl_t.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
    lbl_v = QLabel("0")
    lbl_v.setObjectName("kpiValor")
    v.addWidget(lbl_t)
    v.addWidget(lbl_v)
    v.addStretch()
    return tarjeta, lbl_v


ESTADOS_TEXTO = {
    "REGISTRADO": "Registrado",
    "ASIGNADO": "Asignado",
    "COTIZACION_GENERADA": "Cotización generada",
    "ACEPTADO": "Aceptado",
    "EN_PROCESO": "En proceso",
    "PENDIENTE_REVISION": "Pendiente de revisión",
    "CORRECCION_SOLICITADA": "Corrección solicitada",
    "APROBADO": "Aprobado",
    "FACTURADO": "Facturado",
    "FINALIZADO": "Finalizado",
    "CANCELADO": "Cancelado",
}


def texto_estado(nombre):
    return ESTADOS_TEXTO.get(nombre, str(nombre).replace("_", " ").capitalize())


def fecha_legible(valor):
    texto = str(valor or "").strip()
    for formato in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(texto, formato).strftime("%d/%m/%Y")
        except ValueError:
            continue
    return texto


def texto_roles(cadena):
    if not cadena:
        return "Sin puesto"
    partes = [p.strip() for p in str(cadena).split(",") if p.strip()]
    return ", ".join(ROLES_TEXTO.get(p, p.capitalize()) for p in partes)


def ajustar_columnas(tabla, modos):
    encabezado = tabla.horizontalHeader()
    for columna, modo in modos.items():
        encabezado.setSectionResizeMode(columna, modo)


def poner_celda(tabla, fila, columna, valor):
    texto = "" if valor is None else str(valor)
    item = QTableWidgetItem(texto)
    if texto:
        item.setToolTip(texto)
    tabla.setItem(fila, columna, item)
    return item


def alinear_tabla(tabla, columnas=(), alineacion=Qt.AlignmentFlag.AlignCenter,
                  izquierda=(), derecha=()):
    for columna in columnas:
        encabezado = tabla.horizontalHeaderItem(columna)
        if encabezado is not None:
            encabezado.setTextAlignment(alineacion)
    for columna in izquierda:
        encabezado = tabla.horizontalHeaderItem(columna)
        if encabezado is not None:
            encabezado.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
    for columna in derecha:
        encabezado = tabla.horizontalHeaderItem(columna)
        if encabezado is not None:
            encabezado.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
    for fila in range(tabla.rowCount()):
        for columna in columnas:
            item = tabla.item(fila, columna)
            if item is not None:
                item.setTextAlignment(alineacion)
        for columna in derecha:
            item = tabla.item(fila, columna)
            if item is not None:
                item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)


def tabla_solo_lectura(columnas):
    tabla = QTableWidget(0, columnas)
    tabla.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
    tabla.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
    tabla.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
    tabla.setAlternatingRowColors(True)
    tabla.verticalHeader().setVisible(False)
    tabla.verticalHeader().setDefaultSectionSize(36)
    tabla.setShowGrid(False)
    return tabla


def id_de_fila(tabla):
    fila = tabla.currentRow()
    if fila < 0:
        return None
    item = tabla.item(fila, 0)
    if item is None:
        return None
    try:
        return int(item.text())
    except (TypeError, ValueError):
        return None


def restaurar_seleccion(tabla, idBuscado):
    if idBuscado is None:
        return
    for fila in range(tabla.rowCount()):
        item = tabla.item(fila, 0)
        if item is not None and item.text() == str(idBuscado):
            tabla.selectRow(fila)
            return


def texto_de_fila(tabla, columna):
    fila = tabla.currentRow()
    if fila < 0:
        return ""
    item = tabla.item(fila, columna)
    return item.text().strip() if item is not None else ""


def etiqueta_campo(texto):
    lbl = QLabel(texto)
    lbl.setObjectName("campo")
    return lbl


QSS = """
QWidget { color: #1B2432; font-size: 13px; }
QWidget#central { background: #F4F6FA; }

/* Barra lateral */
QFrame#sidebar { background: #111C2E; border: none; }
QLabel#marca {
    background: #2F6FED; color: #FFFFFF; border-radius: 8px;
    font-size: 15px; font-weight: 700;
}
QLabel#marcaNombre { color: #FFFFFF; font-size: 15px; font-weight: 700; }
QLabel#marcaSub { color: #8FA0B8; font-size: 11px; }
QLabel#navEtiqueta {
    color: #71839C; font-size: 10px; font-weight: 700; letter-spacing: 1px;
}
QPushButton#nav {
    background: transparent; color: #B9C6D8; border: none;
    border-left: 3px solid transparent; border-radius: 6px;
    padding: 10px 14px; text-align: left;
    font-size: 13px; font-weight: 600;
}
QPushButton#nav:hover { background: rgba(255, 255, 255, 0.06); color: #FFFFFF; }
QPushButton#nav:checked {
    background: rgba(47, 111, 237, 0.16); color: #FFFFFF;
    border-left: 3px solid #2F6FED;
}
QFrame#pieSesion { background: rgba(255, 255, 255, 0.04); border-radius: 8px; }
QLabel#sesionNombre { color: #FFFFFF; font-size: 12px; font-weight: 600; }
QLabel#sesionRol { color: #8FA0B8; font-size: 11px; }
QLabel#sesionFecha { color: #71839C; font-size: 11px; }

/* Encabezado del contenido */
QFrame#topbar { background: #FFFFFF; border-bottom: 1px solid #E3E8F0; }
QLabel#bienvenida { color: #111C2E; font-size: 21px; font-weight: 700; }
QLabel#fecha { color: #64748B; font-size: 12px; }
QLabel#sesionVersion { color: #71839C; font-size: 10px; }
QFrame#acentoTitulo { background: #2F6FED; border-radius: 2px; }

/* Tarjetas y textos */
QFrame#card { background: #FFFFFF; border: 1px solid #E3E8F0; border-radius: 10px; }
QFrame#kpiCard {
    background: #FFFFFF; border: 1px solid #E3E8F0;
    border-top: 3px solid #2F6FED; border-radius: 10px;
}
QFrame#kpiCard[tono="azul"] { border-top: 3px solid #2F6FED; }
QFrame#kpiCard[tono="verde"] { border-top: 3px solid #1B7F4B; }
QFrame#kpiCard[tono="ambar"] { border-top: 3px solid #B45309; }
QFrame#kpiCard[tono="morado"] { border-top: 3px solid #6D28D9; }
QFrame#kpiCard[tono="gris"] { border-top: 3px solid #64748B; }
QLabel#subvista { color: #64748B; font-size: 11px; }
QLabel#avatar { background: #2F6FED; color: #FFFFFF; border-radius: 17px; font-size: 12px; font-weight: 700; }
QFrame#tarjetaAcceso { background: #FFFFFF; border: 1px solid #E3E8F0; border-radius: 12px; }
QFrame#separador { background: #E3E8F0; border: none; }
QLabel#marcaAcceso {
    background: #2F6FED; color: #FFFFFF; border-radius: 10px;
    font-size: 16px; font-weight: 700;
}
QDialog#dialogoAcceso, QDialog#dialogoPrimerUso { background: #F4F6FA; }
QLabel#seccion { color: #111C2E; font-size: 14px; font-weight: 700; }
QLabel#ayuda { color: #64748B; font-size: 12px; }
QLabel#campo { color: #334155; font-size: 12px; font-weight: 600; }
QLabel#kpiTitulo { color: #64748B; font-size: 12px; font-weight: 600; }
QLabel#kpiValor { color: #111C2E; font-size: 28px; font-weight: 700; }
QLabel#vacio { color: #7A8798; font-size: 13px; padding: 24px; }
QLabel#loginTitulo { color: #111C2E; font-size: 19px; font-weight: 700; }
QLabel#loginAyuda { color: #64748B; font-size: 12px; }
QLabel#error { color: #B42318; font-size: 12px; font-weight: 600; }
QLabel#ayudaCampo { color: #7A8798; font-size: 11px; }
QTextEdit#errorBox {
    background: #FEF3F2; border: 1px solid #E4B4AF; border-radius: 6px;
    color: #B42318; font-size: 12px; padding: 6px;
}
QLineEdit[error="true"], QTextEdit[error="true"], QComboBox[error="true"] {
    border: 1px solid #B42318; background: #FFF8F7;
}
QLineEdit[error="true"]:focus, QTextEdit[error="true"]:focus, QComboBox[error="true"]:focus {
    border: 1px solid #B42318;
}
QTextBrowser { background: #FFFFFF; border: 1px solid #E3E8F0; border-radius: 8px; padding: 6px; }
QFrame#ayudaLateral { background: #F7F9FC; border-right: 1px solid #E3E8F0; }
QLabel#ayudaTitulo { color: #111C2E; font-size: 16px; font-weight: 700; }
QLabel#ayudaSub { color: #64748B; font-size: 11px; }
QLabel#ayudaPie { color: #94A3B8; font-size: 10px; }
QLabel#ayudaSeccionTitulo { color: #111C2E; font-size: 19px; font-weight: 700; }
QListWidget#ayudaNav { background: transparent; border: none; outline: 0; }
QListWidget#ayudaNav::item {
    color: #334155; padding: 10px 12px; border-radius: 6px;
    font-size: 12px; font-weight: 600;
}
QListWidget#ayudaNav::item:hover { background: #EDF2F9; }
QListWidget#ayudaNav::item:selected { background: #E7EFFC; color: #1D4ED8; }
QTextBrowser#ayudaTexto { background: #FFFFFF; border: none; }
QLabel#estadoCarga { color: #475569; font-size: 12px; padding-right: 8px; }

/* Campos */
QLineEdit, QTextEdit, QListWidget, QComboBox {
    background: #FFFFFF; border: 1px solid #CBD5E1; border-radius: 6px;
    padding: 7px 9px; font-size: 13px; color: #1B2432;
    selection-background-color: #D7E4FB; selection-color: #111C2E;
}
QLineEdit:focus, QTextEdit:focus, QListWidget:focus, QComboBox:focus {
    border: 1px solid #2F6FED;
}
QLineEdit:disabled, QTextEdit:disabled, QComboBox:disabled {
    background: #F1F4F8; color: #94A3B8;
}
QLineEdit:read-only { background: #F1F4F8; color: #64748B; }
QComboBox { padding-right: 24px; }

QComboBox QAbstractItemView {
    background: #FFFFFF; border: 1px solid #CBD5E1; outline: none;
    selection-background-color: #EEF3FB; selection-color: #111C2E;
}
QListWidget::item { padding: 5px; border-radius: 4px; }
QListWidget::item:selected { background: #EEF3FB; color: #111C2E; }
QRadioButton, QCheckBox { color: #1B2432; font-size: 13px; spacing: 8px; }

/* Botones */
QPushButton#primario {
    background: #2F6FED; color: #FFFFFF; border: 1px solid #2F6FED;
    border-radius: 6px; padding: 9px 18px; font-size: 13px; font-weight: 600;
}
QPushButton#primario:hover { background: #255BC7; border-color: #255BC7; }
QPushButton#primario:pressed { background: #1E4CA8; border-color: #1E4CA8; }
QPushButton#primario:disabled {
    background: #B9CDF2; border-color: #B9CDF2; color: #FFFFFF;
}
QPushButton#secundario {
    background: #FFFFFF; color: #334155; border: 1px solid #CBD5E1;
    border-radius: 6px; padding: 8px 16px; font-size: 13px; font-weight: 600;
}
QPushButton#secundario:hover { background: #F1F5FA; border-color: #94A3B8; }
QPushButton#secundario:disabled { color: #94A3B8; border-color: #E2E8F0; }
QPushButton#enlace {
    background: transparent; color: #2F6FED; border: none;
    padding: 6px 2px; font-size: 12px; font-weight: 600; text-align: center;
}
QPushButton#enlace:hover { color: #255BC7; text-decoration: underline; }
QPushButton#peligro {
    background: #FFFFFF; color: #B42318; border: 1px solid #E4B4AF;
    border-radius: 6px; padding: 8px 16px; font-size: 13px; font-weight: 600;
}
QPushButton#peligro:hover { background: #FDF2F1; border-color: #B42318; }
QPushButton#peligro:disabled { color: #D8A5A0; border-color: #F0D6D3; }

/* Tablas */
QTableWidget {
    background: #FFFFFF; alternate-background-color: #FAFBFD;
    border: 1px solid #E3E8F0; border-radius: 8px; font-size: 13px;
    gridline-color: #EDF1F6;
    selection-background-color: #E7EFFC; selection-color: #111C2E;
}
QTableWidget::item { padding: 4px 6px; }
QHeaderView::section {
    background: #F7F9FC; color: #5B6B80; padding: 9px 8px;
    border: none; border-bottom: 1px solid #E3E8F0;
    font-size: 11px; font-weight: 700;
}
QTableCornerButton::section { background: #F7F9FC; border: none; }

/* Menús y barra de estado */
QMenu { background: #FFFFFF; border: 1px solid #E3E8F0; padding: 6px; }
QMenu::item { padding: 8px 26px 8px 12px; border-radius: 4px; color: #1B2432; font-size: 13px; }
QMenu::item:selected { background: #EEF3FB; }
QMenu::item:disabled { color: #94A3B8; }
QMenu::separator { height: 1px; background: #E3E8F0; margin: 5px 8px; }
QStatusBar { background: #FFFFFF; border-top: 1px solid #E3E8F0; color: #5B6B80; font-size: 12px; }
QStatusBar::item { border: none; }
QStatusBar QLabel { color: #5B6B80; font-size: 12px; }

/* Diálogos */
QDialog { background: #FFFFFF; }
QLabel#dialogoTitulo { color: #111C2E; font-size: 16px; font-weight: 700; }
QMessageBox { background: #FFFFFF; }
QMessageBox QLabel { color: #1B2432; font-size: 13px; }
QDialogButtonBox QPushButton {
    background: #FFFFFF; color: #334155; border: 1px solid #CBD5E1;
    border-radius: 6px; padding: 8px 16px; min-width: 84px;
    font-size: 13px; font-weight: 600;
}
QDialogButtonBox QPushButton:hover { background: #F1F5FA; border-color: #94A3B8; }
QDialogButtonBox QPushButton:default {
    background: #2F6FED; color: #FFFFFF; border-color: #2F6FED;
}
QDialogButtonBox QPushButton:default:hover { background: #255BC7; }

/* Barras de desplazamiento */
QScrollBar:vertical { background: transparent; width: 10px; margin: 2px; }
QScrollBar::handle:vertical { background: #C7D2E0; border-radius: 5px; min-height: 28px; }
QScrollBar::handle:vertical:hover { background: #A9B7C9; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: transparent; }
QScrollBar:horizontal { background: transparent; height: 10px; margin: 2px; }
QScrollBar::handle:horizontal { background: #C7D2E0; border-radius: 5px; min-width: 28px; }
QScrollBar::handle:horizontal:hover { background: #A9B7C9; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0; }
QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal { background: transparent; }

QToolTip {
    background: #111C2E; color: #FFFFFF; border: none;
    padding: 6px 8px; font-size: 12px;
}
"""


class DialogoPrimerUso(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Configuración inicial")
        self.setObjectName("dialogoPrimerUso")
        self.setFixedWidth(500)
        self.cuenta = None
        exterior = QVBoxLayout(self)
        exterior.setContentsMargins(16, 16, 16, 16)
        tarjeta = QFrame()
        tarjeta.setObjectName("tarjetaAcceso")
        v = QVBoxLayout(tarjeta)
        v.setContentsMargins(30, 26, 30, 24)
        v.setSpacing(10)
        exterior.addWidget(tarjeta)
        monograma = QLabel("TG")
        monograma.setObjectName("marcaAcceso")
        monograma.setFixedSize(42, 42)
        monograma.setAlignment(Qt.AlignmentFlag.AlignCenter)
        v.addWidget(monograma, 0, Qt.AlignmentFlag.AlignHCenter)
        titulo = QLabel("Bienvenido a The Gonzalez")
        titulo.setObjectName("loginTitulo")
        titulo.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        ayuda = QLabel("Cree la cuenta del gerente. Esta persona administra el sistema "
                       "y puede crear las demás cuentas.")
        ayuda.setObjectName("loginAyuda")
        ayuda.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        ayuda.setWordWrap(True)
        self.nombre = QLineEdit()
        self.nombre.setPlaceholderText("Ej. Juan Pérez López")
        self.nombre.setToolTip("Solo letras y espacios. Ejemplo: Juan Pérez López")
        self.usuario = QLineEdit()
        self.usuario.setPlaceholderText("Ej. jperez")
        self.usuario.setToolTip("De 3 a 20 caracteres: letras, números o guion bajo (_).")
        self.clave = QLineEdit()
        self.clave.setPlaceholderText("Mínimo 8 caracteres")
        self.clave.setEchoMode(QLineEdit.EchoMode.Password)
        self.clave.setToolTip("Al menos 8 caracteres. Anótela en un lugar seguro.")
        self.clave2 = QLineEdit()
        self.clave2.setPlaceholderText("Repita la contraseña")
        self.clave2.setEchoMode(QLineEdit.EchoMode.Password)
        self.clave2.setToolTip("Escriba exactamente la misma contraseña.")
        self.lbl_error = QTextEdit()
        self.lbl_error.setObjectName("errorBox")
        self.lbl_error.setReadOnly(True)
        self.lbl_error.setFixedHeight(96)
        self.lbl_error.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.lbl_error.hide()
        boton = QPushButton("Crear cuenta y entrar")
        boton.setObjectName("primario")
        boton.setDefault(True)
        boton.setCursor(Qt.CursorShape.PointingHandCursor)
        boton.clicked.connect(self.aceptar)
        self.clave2.returnPressed.connect(self.aceptar)
        v.addWidget(titulo)
        v.addWidget(ayuda)
        v.addWidget(ayuda_campo("Todos los campos son obligatorios."))
        v.addWidget(etiqueta_campo("Nombre completo"))
        v.addWidget(self.nombre)
        v.addWidget(etiqueta_campo("Usuario"))
        v.addWidget(self.usuario)
        v.addWidget(ayuda_campo("De 3 a 20 caracteres: letras, números o guion bajo (_). "
                                "Ejemplo: jperez"))
        v.addWidget(etiqueta_campo("Contraseña"))
        v.addWidget(self.clave)
        v.addWidget(ayuda_campo("Mínimo 8 caracteres."))
        v.addWidget(etiqueta_campo("Repetir contraseña"))
        v.addWidget(self.clave2)
        v.addWidget(self.lbl_error)
        v.addWidget(boton)
        self.nombre.setFocus()

    def _error(self, texto):
        self.lbl_error.setPlainText(texto)
        self.lbl_error.show()

    def aceptar(self):
        nombre = self.nombre.text().strip()
        usuario = self.usuario.text().strip()
        clave = self.clave.text()
        self.lbl_error.hide()
        campos = {"nombre": self.nombre, "usuario": self.usuario,
                  "clave": self.clave, "clave2": self.clave2}
        for widget in campos.values():
            marcar_error(widget, False)
        errores = errores_usuario({"nombre": nombre, "usuario": usuario, "rol": "GERENTE",
                                   "clave": clave, "clave2": self.clave2.text()})
        if errores:
            self._error("Corrija lo siguiente:\n" + texto_errores(errores))
            self.layout().activate()
            self.resize(self.width(), self.sizeHint().height() + 4)
            marcar_errores(errores, campos)
            return
        try:
            con = conectar()
        except sqlite3.Error:
            self._error(MSG_BD)
            return
        try:
            if con.execute("SELECT 1 FROM usuario WHERE nombreUsuario = ?;", (usuario,)).fetchone():
                self._error("Ese usuario ya existe. Elija otro.")
                return
            reiniciar_secuencia(con, "usuario")
            crear_usuario(con, usuario, nombre, "GERENTE", clave)
        except (sqlite3.Error, ValueError) as e:
            self._error(str(e) if isinstance(e, ValueError) and str(e) else
                        "No se pudo crear la cuenta. Intente de nuevo.")
            return
        finally:
            con.close()
        self.cuenta = usuario
        self.accept()


class VentanaPrincipal(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("The Gonzalez - Gestión de servicios")
        self.setMinimumSize(1180, 720)
        self._esperas = 0

        central = QWidget()
        central.setObjectName("central")
        self.setCentralWidget(central)
        raiz = QHBoxLayout(central)
        raiz.setContentsMargins(0, 0, 0, 0)
        raiz.setSpacing(0)

        # Barra lateral de navegación
        sidebar = QFrame()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(232)
        lateral = QVBoxLayout(sidebar)
        lateral.setContentsMargins(16, 22, 16, 18)
        lateral.setSpacing(0)

        marca = QHBoxLayout()
        marca.setSpacing(10)
        monograma = QLabel("TG")
        monograma.setObjectName("marca")
        monograma.setFixedSize(38, 38)
        monograma.setAlignment(Qt.AlignmentFlag.AlignCenter)
        textos = QVBoxLayout()
        textos.setSpacing(1)
        lbl_marca = QLabel("The Gonzalez")
        lbl_marca.setObjectName("marcaNombre")
        lbl_sub = QLabel("Gestión de servicios")
        lbl_sub.setObjectName("marcaSub")
        textos.addWidget(lbl_marca)
        textos.addWidget(lbl_sub)
        marca.addWidget(monograma)
        marca.addLayout(textos)
        marca.addStretch()
        lateral.addLayout(marca)
        lateral.addSpacing(30)

        lbl_nav = QLabel("NAVEGACIÓN")
        lbl_nav.setObjectName("navEtiqueta")
        lateral.addWidget(lbl_nav)
        lateral.addSpacing(8)

        self.btn_inicio = QPushButton("Resumen")
        self.btn_clientes = QPushButton("Clientes")
        self.btn_empleados = QPushButton("Empleados")
        self.btn_servicios = QPushButton("Servicios")
        self.btn_facturas = QPushButton("Facturas")
        self.btn_movimientos = QPushButton("Movimientos")
        self.btn_usuarios = QPushButton("Usuarios")
        self.botones_nav = [self.btn_inicio, self.btn_clientes, self.btn_empleados,
                            self.btn_servicios, self.btn_facturas, self.btn_movimientos,
                            self.btn_usuarios]
        self.tipos_nav = ("resumen", "clientes", "empleados", "servicios", "facturas",
                          "movimientos", "usuarios")
        for boton, tipo in zip(self.botones_nav, self.tipos_nav):
            boton.setObjectName("nav")
            boton.setCheckable(True)
            boton.setCursor(Qt.CursorShape.PointingHandCursor)
            boton.setMinimumHeight(40)
            boton.setIcon(icono_nav(tipo))
            boton.setIconSize(QSize(18, 18))
            lateral.addWidget(boton)
            lateral.addSpacing(2)
        if not puede(PERMISO_GERENTE):
            self.btn_usuarios.hide()
        if not puede(PERMISO_GERENTE):
            self.btn_movimientos.hide()
        if not puede(PERMISO_FACTURACION):
            self.btn_facturas.hide()
        lateral.addStretch()

        pie = QFrame()
        pie.setObjectName("pieSesion")
        p = QVBoxLayout(pie)
        p.setContentsMargins(12, 12, 12, 10)
        p.setSpacing(3)
        nombre_sesion = SESION.get("nombreCompleto") or "Usuario"
        fila_sesion = QHBoxLayout()
        fila_sesion.setSpacing(10)
        self.avatar_sesion = Avatar(nombre_sesion, SESION.get("foto") or "", 34)
        textos_sesion = QVBoxLayout()
        textos_sesion.setSpacing(1)
        self.lbl_sesion_nombre = QLabel(nombre_sesion)
        self.lbl_sesion_nombre.setObjectName("sesionNombre")
        rol = str(SESION.get("rol") or "").strip().upper()
        self.lbl_sesion_rol = QLabel(ROLES_TEXTO.get(rol, rol.capitalize() or "Sin puesto"))
        self.lbl_sesion_rol.setObjectName("sesionRol")
        textos_sesion.addWidget(self.lbl_sesion_nombre)
        textos_sesion.addWidget(self.lbl_sesion_rol)
        fila_sesion.addWidget(self.avatar_sesion)
        fila_sesion.addLayout(textos_sesion)
        fila_sesion.addStretch()
        p.addLayout(fila_sesion)
        p.addSpacing(4)
        lbl_fecha = QLabel(date.today().strftime("%d/%m/%Y"))
        lbl_fecha.setObjectName("sesionFecha")
        lbl_version = QLabel(f"Versión {VERSION}")
        lbl_version.setObjectName("sesionVersion")
        p.addWidget(lbl_fecha)
        p.addWidget(lbl_version)
        lateral.addWidget(pie)
        raiz.addWidget(sidebar)

        # Contenido principal
        contenido = QWidget()
        columna = QVBoxLayout(contenido)
        columna.setContentsMargins(0, 0, 0, 0)
        columna.setSpacing(0)

        topbar = QFrame()
        topbar.setObjectName("topbar")
        t = QHBoxLayout(topbar)
        t.setContentsMargins(32, 14, 32, 14)
        self.lbl_vista = QLabel("Resumen general")
        self.lbl_vista.setObjectName("bienvenida")
        self.lbl_subvista = QLabel("Indicadores y últimos servicios registrados")
        self.lbl_subvista.setObjectName("subvista")
        barra_titulo = QFrame()
        barra_titulo.setObjectName("acentoTitulo")
        barra_titulo.setFixedSize(4, 28)
        bloque_titulo = QVBoxLayout()
        bloque_titulo.setSpacing(1)
        bloque_titulo.addWidget(self.lbl_vista)
        bloque_titulo.addWidget(self.lbl_subvista)
        t.addWidget(barra_titulo)
        t.addSpacing(12)
        t.addLayout(bloque_titulo)
        t.addStretch()
        self.btn_cuenta = QPushButton("Mi cuenta")
        self.btn_cuenta.setObjectName("secundario")
        self.btn_cuenta.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_cuenta.setToolTip("Cambiar foto, datos personales, correo, teléfono y contraseña.")
        self.btn_cuenta.clicked.connect(self.abrir_cuenta)
        t.addWidget(self.btn_cuenta)
        t.addSpacing(8)
        self.btn_avisos = QPushButton("Avisos")
        self.btn_avisos.setObjectName("secundario")
        self.btn_avisos.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_avisos.setToolTip("Avisos, cambios de puesto y actualizaciones del sistema.")
        self.btn_avisos.clicked.connect(self.abrir_avisos)
        t.addWidget(self.btn_avisos)
        t.addSpacing(8)
        self.btn_salir = QPushButton("Cerrar sesión")
        self.btn_salir.setObjectName("secundario")
        self.btn_salir.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_salir.setToolTip("Cerrar sesión y marcar al usuario como desconectado.")
        self.btn_salir.clicked.connect(self.cerrar_sesion)
        t.addWidget(self.btn_salir)
        t.addSpacing(8)
        self.btn_ayuda = QPushButton("Ayuda")
        self.btn_ayuda.setObjectName("secundario")
        self.btn_ayuda.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_ayuda.setToolTip("Cómo llenar los formularios y qué hacer ante un error (F1).")
        self.btn_ayuda.clicked.connect(self.mostrar_ayuda)
        t.addWidget(self.btn_ayuda)
        t.addSpacing(16)
        lbl_fecha_hoy = QLabel(date.today().strftime("%d/%m/%Y"))
        lbl_fecha_hoy.setObjectName("fecha")
        t.addWidget(lbl_fecha_hoy)
        columna.addWidget(topbar)

        self.vistas = QStackedWidget()
        columna.addWidget(self.vistas, 1)
        raiz.addWidget(contenido, 1)

        self.indicador = Indicador()
        self.lbl_cargando = QLabel("Cargando información...")
        self.lbl_cargando.setObjectName("estadoCarga")
        self.lbl_cargando.hide()
        self.statusBar().addPermanentWidget(self.lbl_cargando)
        self.statusBar().addPermanentWidget(self.indicador)
        QShortcut(QKeySequence("F1"), self, activated=self.mostrar_ayuda)
        self.timer_avisos = QTimer(self)
        self.timer_avisos.timeout.connect(self.actualizar_avisos)
        self.timer_avisos.start(60000)

        self.setup_inicio()
        self.setup_clientes()
        self.setup_empleados()
        self.setup_servicios()
        self.setup_usuarios()
        self.setup_movimientos()
        self.setup_facturas()

        for indice, boton in enumerate(self.botones_nav):
            boton.clicked.connect(lambda _=False, i=indice: self.cambiar_vista(i))

        self.actualizar_tabla_clientes()
        self.actualizar_tabla_empleados()
        self.actualizar_tabla_servicios()
        self.actualizar_sesion()
        self.actualizar_avisos()
        self.cambiar_vista(0)

    def iniciar_espera(self):
        self._esperas += 1
        if self._esperas == 1:
            QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
            self.indicador.iniciar()
            self.lbl_cargando.show()
        QApplication.processEvents()

    def terminar_espera(self):
        self._esperas = max(0, self._esperas - 1)
        if self._esperas == 0:
            QApplication.restoreOverrideCursor()
            self.indicador.detener()
            self.lbl_cargando.hide()

    def abrir_avisos(self):
        DialogoNotificaciones(self).exec()
        self.actualizar_avisos()

    def actualizar_avisos(self):
        id_usuario = SESION.get("id")
        if not id_usuario:
            return
        try:
            con = conectar()
            try:
                total = con.execute(
                    "SELECT COUNT(*) FROM notificacion WHERE idUsuario = ? AND leida = 0;",
                    (id_usuario,)).fetchone()[0]
            finally:
                con.close()
        except sqlite3.Error:
            total = 0
        self.btn_avisos.setText(f"Avisos ({total})" if total else "Avisos")
        self.btn_avisos.setToolTip(
            f"{total} aviso(s) sin leer. Cambios de puesto y actualizaciones." if total
            else "Avisos, cambios de puesto y actualizaciones del sistema.")

    def cerrar_sesion(self):
        if not confirmar(self, "Cerrar sesión",
                         "¿Desea cerrar la sesión actual?\n\n"
                         "El usuario quedará marcado como desconectado para que el "
                         "administrador pueda cambiar su puesto."):
            return
        marcar_conectado(False)
        self.close()

    def abrir_cuenta(self):
        dlg = DialogoCuenta(self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            self.actualizar_sesion()
            self.statusBar().showMessage("Datos de la cuenta actualizados.", 5000)

    def actualizar_sesion(self):
        try:
            con = conectar()
            try:
                columnas = {fila[1] for fila in con.execute("PRAGMA table_info(usuario);").fetchall()}
                campos = ["nombreUsuario", "nombreCompleto", "rol"]
                if "foto" in columnas:
                    campos.append("foto")
                fila = con.execute(
                    f"SELECT {', '.join(campos)} FROM usuario WHERE idUsuario = ?;",
                    (SESION.get("id"),)).fetchone()
            finally:
                con.close()
        except sqlite3.Error:
            return
        if fila is None:
            return
        SESION["nombreCompleto"] = fila["nombreCompleto"] or fila["nombreUsuario"]
        SESION["rol"] = str(fila["rol"] or "").strip().upper()
        SESION["foto"] = fila["foto"] if "foto" in fila.keys() else None
        self.lbl_sesion_nombre.setText(SESION["nombreCompleto"])
        self.lbl_sesion_rol.setText(ROLES_TEXTO.get(SESION["rol"], SESION["rol"].capitalize() or "Sin puesto"))
        self.avatar_sesion.set_datos(SESION["nombreCompleto"], SESION["foto"] or "")

    def closeEvent(self, evento):
        marcar_conectado(False)
        QSettings("TheGonzalez", "Servicios").setValue("ventana/geometria", self.saveGeometry())
        super().closeEvent(evento)

    def mostrar_ayuda(self):
        seccion = {0: 0, 1: 2, 2: 3, 3: 4, 4: 5}.get(self.vistas.currentIndex(), 0)
        DialogoAyuda(self, seccion).exec()

    def cambiar_vista(self, index):
        titulos = ("Resumen general", "Clientes", "Empleados", "Servicios", "Facturas",
                   "Movimientos", "Usuarios")
        subtitulos = (
            "Indicadores y últimos servicios registrados",
            "Datos de contacto y registro de clientes",
            "Personal activo y puestos de trabajo",
            "Registro, seguimiento y facturación de servicios",
            "Facturas emitidas e importes cobrados",
            "Quién hizo qué, cuándo y con qué importe",
            "Cuentas de acceso y permisos",
        )
        self.vistas.setCurrentIndex(index)
        self.lbl_vista.setText(titulos[index])
        self.lbl_subvista.setText(subtitulos[index])
        for i, b in enumerate(self.botones_nav):
            activo = i == index
            b.setChecked(activo)
            b.setIcon(icono_nav(self.tipos_nav[i], "#FFFFFF" if activo else "#B9C6D8"))
        if index == 0:
            self.actualizar_inicio()
        elif index == 1:
            self.actualizar_tabla_clientes()
        elif index == 2:
            self.actualizar_tabla_empleados()
        elif index == 3:
            self.cargar_clientes_combo()
            self.actualizar_tabla_servicios()
        elif index == 4:
            self.actualizar_tabla_facturas()
        elif index == 5:
            self.actualizar_tabla_movimientos()
        elif index == 6:
            self.actualizar_tabla_usuarios()

    def conectar_vacio(self, tabla, etiqueta, texto):
        etiqueta.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        etiqueta.setAlignment(Qt.AlignmentFlag.AlignCenter)
        etiqueta.setWordWrap(True)

        def actualizar():
            vacia = tabla.rowCount() == 0
            etiqueta.setText(texto)
            etiqueta.setVisible(vacia)
            tabla.setVisible(not vacia)
        tabla.model().rowsInserted.connect(lambda *a: actualizar())
        tabla.model().rowsRemoved.connect(lambda *a: actualizar())
        actualizar()

    def setup_inicio(self):
        vista = QWidget()
        v = QVBoxLayout(vista)
        v.setContentsMargins(32, 22, 32, 24)
        v.setSpacing(18)

        fila = QHBoxLayout()
        fila.setSpacing(16)
        self.kpi = {}
        for clave, titulo, tono in [("clientes", "Clientes activos", "azul"),
                                    ("empleados", "Empleados activos", "verde"),
                                    ("porcotizar", "Servicios por cotizar", "ambar"),
                                    ("abiertos", "Servicios en proceso", "morado"),
                                    ("finalizados", "Servicios finalizados", "gris")]:
            tarjeta, valor = tarjeta_indicador(titulo, tono)
            self.kpi[clave] = valor
            fila.addWidget(tarjeta)
        v.addLayout(fila)

        panel = PanelContenedor("Servicios recientes",
                                "Los últimos 8 servicios registrados. Cotice y facture desde la sección Servicios.")
        self.tabla_recientes = tabla_solo_lectura(5)
        self.tabla_recientes.setHorizontalHeaderLabels(["ID", "Fecha", "Cliente", "Estado", "Problema"])
        ajustar_columnas(self.tabla_recientes, {
            0: QHeaderView.ResizeMode.ResizeToContents,
            1: QHeaderView.ResizeMode.ResizeToContents,
            2: QHeaderView.ResizeMode.ResizeToContents,
            3: QHeaderView.ResizeMode.ResizeToContents,
            4: QHeaderView.ResizeMode.Stretch,
        })
        self.vacio_recientes = QLabel("Todavía no hay servicios registrados.")
        self.vacio_recientes.setObjectName("vacio")
        self.vacio_recientes.setAlignment(Qt.AlignmentFlag.AlignCenter)
        panel.layout_principal.addWidget(self.tabla_recientes, 1)
        panel.layout_principal.addWidget(self.vacio_recientes)
        self.conectar_vacio(self.tabla_recientes, self.vacio_recientes, "Todavía no hay servicios registrados.")
        v.addWidget(panel, 1)
        self.vistas.addWidget(vista)

    @con_espera
    def actualizar_inicio(self):
        con = conectar()
        try:
            cifras = {
                "clientes": con.execute("SELECT COUNT(*) FROM cliente WHERE activo = 1;").fetchone()[0],
                "empleados": con.execute("SELECT COUNT(*) FROM empleado WHERE activo = 1;").fetchone()[0],
                "porcotizar": con.execute(
                    """SELECT COUNT(*) FROM servicio s JOIN estado_servicio es ON es.idEstado = s.idEstado
                       WHERE es.nombreEstado IN ('REGISTRADO', 'ASIGNADO');""").fetchone()[0],
                "abiertos": con.execute(
                    """SELECT COUNT(*) FROM servicio s JOIN estado_servicio es ON es.idEstado = s.idEstado
                       WHERE es.nombreEstado IN ('COTIZACION_GENERADA', 'ACEPTADO', 'EN_PROCESO',
                                                 'PENDIENTE_REVISION', 'CORRECCION_SOLICITADA', 'APROBADO',
                                                 'FACTURADO');""").fetchone()[0],
                "finalizados": con.execute(
                    """SELECT COUNT(*) FROM servicio s JOIN estado_servicio es ON es.idEstado = s.idEstado
                       WHERE es.nombreEstado = 'FINALIZADO';""").fetchone()[0],
            }
            recientes = con.execute(
                """SELECT s.idServicio, s.fechaRegistro, c.nombre, es.nombreEstado, s.descripcionProblema
                   FROM servicio s
                   JOIN cliente c ON c.idCliente = s.idCliente
                   JOIN estado_servicio es ON es.idEstado = s.idEstado
                   ORDER BY s.idServicio DESC LIMIT 8;""").fetchall()
        finally:
            con.close()
        for clave, valor in cifras.items():
            self.kpi[clave].setText(str(valor))
        self.tabla_recientes.setRowCount(0)
        for s in recientes:
            row = self.tabla_recientes.rowCount()
            self.tabla_recientes.insertRow(row)
            valores = [s[0], fecha_legible(s[1]), s[2], texto_estado(s[3]), s[4]]
            for col, dato in enumerate(valores):
                poner_celda(self.tabla_recientes, row, col, dato)
        alinear_tabla(self.tabla_recientes, (0, 1, 3), izquierda=(2, 4))
        self.statusBar().showMessage(
            f"Clientes activos: {cifras['clientes']}. En proceso: {cifras['abiertos']}. "
            f"Por cotizar: {cifras['porcotizar']}.")

    def setup_clientes(self):
        vista = QWidget()
        layout = QHBoxLayout(vista)
        layout.setContentsMargins(32, 22, 32, 24)
        layout.setSpacing(20)

        panel_reg = PanelContenedor("Nuevo cliente",
                                    "El nombre es obligatorio. Teléfono: 10 dígitos. "
                                    "RFC: 12 o 13 caracteres. Los demás datos son opcionales.")
        panel_reg.setMinimumWidth(370)
        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.AllNonFixedFieldsGrow)
        form.setVerticalSpacing(10)
        self.cli_nombre = QLineEdit()
        self.cli_nombre.setPlaceholderText("Ej. Juan Pérez López")
        self.cli_nombre.setToolTip("Obligatorio. Solo letras y espacios.")
        self.cli_tel = QLineEdit()
        self.cli_tel.setPlaceholderText("Ej. 8441234567")
        self.cli_tel.setToolTip("Exactamente 10 dígitos, sin espacios ni guiones.")
        self.cli_correo = QLineEdit()
        self.cli_correo.setPlaceholderText("Ej. juan@correo.com")
        self.cli_correo.setToolTip("Formato nombre@dominio.com")
        self.cli_dir = QLineEdit()
        self.cli_dir.setPlaceholderText("Ej. Av. Juárez 120")
        self.cli_dir.setToolTip("Calle, número y colonia. Opcional.")
        self.cli_rfc = QLineEdit()
        self.cli_rfc.setPlaceholderText("Ej. RUIA010101AB1")
        self.cli_rfc.setToolTip("12 o 13 caracteres: letras, fecha y homoclave. Opcional.")
        form.addRow(etiqueta_campo("Nombre"), self.cli_nombre)
        form.addRow(etiqueta_campo("Teléfono"), self.cli_tel)
        form.addRow(etiqueta_campo("Correo"), self.cli_correo)
        form.addRow(etiqueta_campo("Dirección"), self.cli_dir)
        form.addRow(etiqueta_campo("RFC"), self.cli_rfc)
        btn_guardar = QPushButton("Guardar cliente")
        btn_guardar.setObjectName("primario")
        btn_guardar.setToolTip("Revisa los datos y pide confirmación antes de guardar.")
        btn_guardar.clicked.connect(self.guardar_cliente)
        btn_guardar.setMinimumWidth(170)
        fila_boton = QHBoxLayout()
        fila_boton.addStretch()
        fila_boton.addWidget(btn_guardar)
        panel_reg.layout_principal.addLayout(form)
        panel_reg.layout_principal.addLayout(fila_boton)
        panel_reg.layout_principal.addStretch()

        panel_list = PanelContenedor("Lista de clientes",
                                     "Seleccione un cliente para modificarlo o darlo de baja.")
        busqueda = QHBoxLayout()
        self.cli_buscar = QLineEdit()
        self.cli_buscar.setPlaceholderText("Escriba un nombre para buscar")
        self.cli_buscar.addAction(icono_buscar(), QLineEdit.ActionPosition.LeadingPosition)
        self.cli_buscar.setClearButtonEnabled(True)
        self.cli_buscar.returnPressed.connect(self.buscar_clientes)
        btn_buscar = QPushButton("Buscar")
        btn_buscar.setObjectName("secundario")
        btn_buscar.clicked.connect(self.buscar_clientes)
        busqueda.addWidget(self.cli_buscar)
        busqueda.addWidget(btn_buscar)

        self.tabla_clientes = tabla_solo_lectura(4)
        self.tabla_clientes.setHorizontalHeaderLabels(["ID", "Nombre", "Teléfono", "Correo"])
        ajustar_columnas(self.tabla_clientes, {
            0: QHeaderView.ResizeMode.ResizeToContents,
            1: QHeaderView.ResizeMode.Stretch,
            2: QHeaderView.ResizeMode.ResizeToContents,
            3: QHeaderView.ResizeMode.Stretch,
        })
        self.vacio_clientes = QLabel("No hay clientes con ese nombre. Puede registrarlo en el formulario.")
        self.vacio_clientes.setObjectName("vacio")
        self.vacio_clientes.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.conectar_vacio(self.tabla_clientes, self.vacio_clientes,
                            "No hay clientes con ese nombre. Puede registrarlo en el formulario.")

        acciones = QHBoxLayout()
        self.btn_cli_modificar = QPushButton("Modificar")
        self.btn_cli_modificar.setObjectName("secundario")
        self.btn_cli_modificar.setToolTip("Cambia los datos del cliente seleccionado.")
        self.btn_cli_eliminar = QPushButton("Dar de baja")
        self.btn_cli_eliminar.setObjectName("peligro")
        self.btn_cli_eliminar.setToolTip("Quita al cliente de la lista. Su historial se conserva.")
        self.btn_cli_modificar.clicked.connect(self.modificar_cliente)
        self.btn_cli_eliminar.clicked.connect(self.eliminar_cliente)
        acciones.addStretch()
        acciones.addWidget(self.btn_cli_modificar)
        acciones.addWidget(self.btn_cli_eliminar)

        panel_list.layout_principal.addLayout(busqueda)
        panel_list.layout_principal.addWidget(self.tabla_clientes, 1)
        panel_list.layout_principal.addWidget(self.vacio_clientes)
        panel_list.layout_principal.addLayout(acciones)
        layout.addWidget(panel_reg, 1)
        layout.addWidget(panel_list, 2)
        self.vistas.addWidget(vista)

    def setup_empleados(self):
        vista = QWidget()
        layout = QHBoxLayout(vista)
        layout.setContentsMargins(32, 22, 32, 24)
        layout.setSpacing(20)

        panel_reg = PanelContenedor("Nuevo empleado",
                                    "Nombre y apellido paterno obligatorios. Teléfono: 10 dígitos. "
                                    "Elija al menos un puesto.")
        panel_reg.setMinimumWidth(370)
        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.AllNonFixedFieldsGrow)
        form.setVerticalSpacing(10)
        self.emp_nombre = QLineEdit()
        self.emp_nombre.setPlaceholderText("Ej. Juan")
        self.emp_nombre.setToolTip("Obligatorio. Solo letras y espacios.")
        self.emp_ap = QLineEdit()
        self.emp_ap.setPlaceholderText("Ej. Pérez")
        self.emp_ap.setToolTip("Obligatorio. Solo letras y espacios.")
        self.emp_am = QLineEdit()
        self.emp_am.setPlaceholderText("Ej. López")
        self.emp_am.setToolTip("Opcional. Solo letras y espacios.")
        self.emp_tel = QLineEdit()
        self.emp_tel.setPlaceholderText("Ej. 8441234567")
        self.emp_tel.setToolTip("Exactamente 10 dígitos, sin espacios ni guiones.")
        self.emp_correo = QLineEdit()
        self.emp_correo.setPlaceholderText("Ej. juan@correo.com")
        self.emp_correo.setToolTip("Formato nombre@dominio.com")
        form.addRow(etiqueta_campo("Nombre"), self.emp_nombre)
        form.addRow(etiqueta_campo("Apellido paterno"), self.emp_ap)
        form.addRow(etiqueta_campo("Apellido materno"), self.emp_am)
        form.addRow(etiqueta_campo("Teléfono"), self.emp_tel)
        form.addRow(etiqueta_campo("Correo"), self.emp_correo)
        self.emp_roles = QListWidget()
        self.emp_roles.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        self.emp_roles.setMaximumHeight(150)
        self.emp_roles.setToolTip("Marque uno o más puestos con el mouse.")
        self.cargar_roles_listado()
        btn_guardar = QPushButton("Guardar empleado")
        btn_guardar.setObjectName("primario")
        btn_guardar.setToolTip("Revisa los datos y pide confirmación antes de guardar.")
        btn_guardar.clicked.connect(self.guardar_empleado)
        panel_reg.layout_principal.addLayout(form)
        panel_reg.layout_principal.addWidget(etiqueta_campo("Puestos (marque uno o más)"))
        panel_reg.layout_principal.addWidget(self.emp_roles)
        btn_guardar.setMinimumWidth(170)
        fila_boton = QHBoxLayout()
        fila_boton.addStretch()
        fila_boton.addWidget(btn_guardar)
        panel_reg.layout_principal.addLayout(fila_boton)
        panel_reg.layout_principal.addStretch()

        panel_list = PanelContenedor("Lista de empleados",
                                     "Personal activo con sus puestos. Seleccione uno para modificarlo o darlo de baja.")
        busqueda = QHBoxLayout()
        self.emp_buscar = QLineEdit()
        self.emp_buscar.setPlaceholderText("Escriba un nombre para buscar")
        self.emp_buscar.addAction(icono_buscar(), QLineEdit.ActionPosition.LeadingPosition)
        self.emp_buscar.setClearButtonEnabled(True)
        self.emp_buscar.returnPressed.connect(self.actualizar_tabla_empleados)
        btn_buscar = QPushButton("Buscar")
        btn_buscar.setObjectName("secundario")
        btn_buscar.clicked.connect(self.actualizar_tabla_empleados)
        busqueda.addWidget(self.emp_buscar)
        busqueda.addWidget(btn_buscar)

        self.tabla_empleados = tabla_solo_lectura(4)
        self.tabla_empleados.setHorizontalHeaderLabels(["ID", "Nombre", "Teléfono", "Puestos"])
        ajustar_columnas(self.tabla_empleados, {
            0: QHeaderView.ResizeMode.ResizeToContents,
            1: QHeaderView.ResizeMode.Stretch,
            2: QHeaderView.ResizeMode.ResizeToContents,
            3: QHeaderView.ResizeMode.Stretch,
        })
        self.vacio_empleados = QLabel("No hay empleados con ese nombre. Puede registrarlo en el formulario.")
        self.vacio_empleados.setObjectName("vacio")
        self.vacio_empleados.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.conectar_vacio(self.tabla_empleados, self.vacio_empleados,
                            "No hay empleados con ese nombre. Puede registrarlo en el formulario.")

        acciones = QHBoxLayout()
        self.btn_emp_modificar = QPushButton("Modificar")
        self.btn_emp_modificar.setObjectName("secundario")
        self.btn_emp_modificar.setToolTip("Cambia los datos y puestos del empleado seleccionado.")
        self.btn_emp_eliminar = QPushButton("Dar de baja")
        self.btn_emp_eliminar.setObjectName("peligro")
        self.btn_emp_eliminar.setToolTip("No se permite si el empleado está en un servicio en proceso.")
        self.btn_emp_modificar.clicked.connect(self.modificar_empleado)
        self.btn_emp_eliminar.clicked.connect(self.eliminar_empleado)
        acciones.addStretch()
        acciones.addWidget(self.btn_emp_modificar)
        acciones.addWidget(self.btn_emp_eliminar)

        panel_list.layout_principal.addLayout(busqueda)
        panel_list.layout_principal.addWidget(self.tabla_empleados, 1)
        panel_list.layout_principal.addWidget(self.vacio_empleados)
        panel_list.layout_principal.addLayout(acciones)
        layout.addWidget(panel_reg, 1)
        layout.addWidget(panel_list, 2)
        self.vistas.addWidget(vista)

    def setup_servicios(self):
        vista = QWidget()
        layout = QHBoxLayout(vista)
        layout.setContentsMargins(32, 22, 32, 24)
        layout.setSpacing(20)

        panel_reg = PanelContenedor("Nuevo servicio",
                                    "Escriba un cliente ya registrado y describa el problema "
                                    "con al menos 3 letras.")
        panel_reg.setMinimumWidth(370)
        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.AllNonFixedFieldsGrow)
        form.setVerticalSpacing(10)
        self.srv_cliente = QComboBox()
        self.srv_cliente.setEditable(True)
        self.srv_cliente.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.srv_cliente.setMaxVisibleItems(12)
        self.srv_cliente.lineEdit().setPlaceholderText("Escriba o elija un cliente")
        self.srv_cliente.setToolTip("Escriba parte del nombre y elija el cliente correcto de la lista.")
        completador = self.srv_cliente.completer()
        completador.setCompletionMode(QCompleter.CompletionMode.PopupCompletion)
        completador.setFilterMode(Qt.MatchFlag.MatchContains)
        completador.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        form.addRow(etiqueta_campo("Cliente"), self.srv_cliente)
        self.srv_problema = QTextEdit()
        self.srv_problema.setPlaceholderText("Ej. Falla la bomba de agua")
        self.srv_problema.setToolTip("Obligatorio. Al menos 3 letras.")
        self.srv_problema.setMaximumHeight(90)
        self.srv_descripcion = QTextEdit()
        self.srv_descripcion.setMaximumHeight(80)
        self.srv_obs = QTextEdit()
        self.srv_obs.setMaximumHeight(70)
        btn_guardar = QPushButton("Guardar servicio")
        btn_guardar.setObjectName("primario")
        btn_guardar.setToolTip("Revisa los datos y pide confirmación antes de registrar el servicio.")
        btn_guardar.clicked.connect(self.guardar_servicio)
        panel_reg.layout_principal.addLayout(form)
        panel_reg.layout_principal.addWidget(etiqueta_campo("Problema"))
        panel_reg.layout_principal.addWidget(self.srv_problema)
        panel_reg.layout_principal.addWidget(etiqueta_campo("Descripción del servicio (opcional)"))
        panel_reg.layout_principal.addWidget(self.srv_descripcion)
        panel_reg.layout_principal.addWidget(etiqueta_campo("Observaciones (opcional)"))
        panel_reg.layout_principal.addWidget(self.srv_obs)
        btn_guardar.setMinimumWidth(170)
        fila_boton = QHBoxLayout()
        fila_boton.addStretch()
        fila_boton.addWidget(btn_guardar)
        panel_reg.layout_principal.addLayout(fila_boton)
        panel_reg.layout_principal.addStretch()

        panel_list = PanelContenedor("Lista de servicios",
                                     "Estado actual de cada servicio. Seleccione uno para ver las acciones disponibles.")
        busqueda = QHBoxLayout()
        self.srv_buscar = QLineEdit()
        self.srv_buscar.setPlaceholderText("Buscar por cliente o problema")
        self.srv_buscar.addAction(icono_buscar(), QLineEdit.ActionPosition.LeadingPosition)
        self.srv_buscar.setClearButtonEnabled(True)
        self.srv_buscar.returnPressed.connect(self.actualizar_tabla_servicios)
        btn_buscar = QPushButton("Buscar")
        btn_buscar.setObjectName("secundario")
        btn_buscar.clicked.connect(self.actualizar_tabla_servicios)
        busqueda.addWidget(self.srv_buscar)
        busqueda.addWidget(btn_buscar)

        self.tabla_servicios = tabla_solo_lectura(6)
        self.tabla_servicios.setHorizontalHeaderLabels(
            ["ID", "Fecha", "Cliente", "Estado", "Problema", "Última cotización"])
        ajustar_columnas(self.tabla_servicios, {
            0: QHeaderView.ResizeMode.ResizeToContents,
            1: QHeaderView.ResizeMode.ResizeToContents,
            2: QHeaderView.ResizeMode.ResizeToContents,
            3: QHeaderView.ResizeMode.ResizeToContents,
            4: QHeaderView.ResizeMode.Stretch,
            5: QHeaderView.ResizeMode.ResizeToContents,
        })
        self.vacio_servicios = QLabel("No hay servicios que mostrar. Registre uno en el formulario.")
        self.vacio_servicios.setObjectName("vacio")
        self.vacio_servicios.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.conectar_vacio(self.tabla_servicios, self.vacio_servicios,
                            "No hay servicios que mostrar. Registre uno en el formulario.")

        acciones = QHBoxLayout()
        self.menu_acciones = QMenu(self)
        self.menu_acciones.aboutToShow.connect(self.llenar_menu_acciones)
        self.btn_acciones = QPushButton("Acciones")
        self.btn_acciones.setObjectName("primario")
        self.btn_acciones.setToolTip("Muestra lo que puede hacer con el servicio seleccionado.")
        self.btn_acciones.setMenu(self.menu_acciones)
        self.btn_srv_modificar = QPushButton("Modificar")
        self.btn_srv_modificar.setObjectName("secundario")
        self.btn_srv_modificar.setToolTip("Cambia el problema, la descripción o las observaciones.")
        self.btn_srv_eliminar = QPushButton("Cancelar servicio")
        self.btn_srv_eliminar.setObjectName("peligro")
        self.btn_srv_eliminar.setToolTip("Cancela el servicio indicando el motivo. El registro se conserva.")
        self.btn_srv_modificar.clicked.connect(self.modificar_servicio)
        self.btn_srv_eliminar.clicked.connect(self.cancelar_servicio)
        acciones.addStretch()
        acciones.addWidget(self.btn_acciones)
        acciones.addWidget(self.btn_srv_modificar)
        acciones.addWidget(self.btn_srv_eliminar)

        panel_list.layout_principal.addLayout(busqueda)
        panel_list.layout_principal.addWidget(self.tabla_servicios, 1)
        panel_list.layout_principal.addWidget(self.vacio_servicios)
        panel_list.layout_principal.addLayout(acciones)
        layout.addWidget(panel_reg, 1)
        layout.addWidget(panel_list, 2)
        self.vistas.addWidget(vista)
        self.cargar_clientes_combo()

    def cargar_clientes_combo(self):
        try:
            con = conectar()
            try:
                filas = con.execute(
                    "SELECT idCliente, nombre FROM cliente WHERE activo = 1 ORDER BY nombre;"
                ).fetchall()
            finally:
                con.close()
        except sqlite3.Error:
            filas = []
        self.srv_cliente.blockSignals(True)
        self.srv_cliente.clear()
        for fila in filas:
            self.srv_cliente.addItem(fila["nombre"], fila["idCliente"])
            indice = self.srv_cliente.count() - 1
            self.srv_cliente.setItemData(indice, f"Cliente ID {fila['idCliente']}",
                                         Qt.ItemDataRole.ToolTipRole)
        self.srv_cliente.setCurrentIndex(-1)
        self.srv_cliente.setEditText("")
        self.srv_cliente.lineEdit().setPlaceholderText(
            "Escriba o elija un cliente" if filas else "No hay clientes registrados")
        completador = self.srv_cliente.completer()
        if completador is not None:
            completador.setCompletionMode(QCompleter.CompletionMode.PopupCompletion)
            completador.setFilterMode(Qt.MatchFlag.MatchContains)
            completador.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        self.srv_cliente.blockSignals(False)

    def setup_usuarios(self):
        vista = QWidget()
        layout = QHBoxLayout(vista)
        layout.setContentsMargins(32, 22, 32, 24)
        layout.setSpacing(20)

        panel_reg = PanelContenedor("Nuevo usuario",
                                    "Nombre: solo letras. Usuario: 3 a 20 caracteres. "
                                    "Contraseña: mínimo 8. El puesto define sus permisos.")
        panel_reg.setMinimumWidth(370)
        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.AllNonFixedFieldsGrow)
        form.setVerticalSpacing(10)
        self.usu_nombre = QLineEdit()
        self.usu_nombre.setPlaceholderText("Ej. Juan Pérez López")
        self.usu_nombre.setToolTip("Solo letras y espacios.")
        self.usu_usuario = QLineEdit()
        self.usu_usuario.setPlaceholderText("Ej. jperez")
        self.usu_usuario.setToolTip("De 3 a 20 caracteres: letras, números o guion bajo.")
        self.usu_rol = QComboBox()
        for codigo in ROLES_SISTEMA:
            self.usu_rol.addItem(ROLES_TEXTO.get(codigo, codigo.capitalize()), codigo)
        indice_cotizador = self.usu_rol.findData("COTIZADOR")
        if indice_cotizador >= 0:
            self.usu_rol.setCurrentIndex(indice_cotizador)
        self.usu_clave = QLineEdit()
        self.usu_clave.setPlaceholderText("Mínimo 8 caracteres")
        self.usu_clave.setEchoMode(QLineEdit.EchoMode.Password)
        self.usu_clave.setToolTip("Al menos 8 caracteres.")
        self.usu_clave2 = QLineEdit()
        self.usu_clave2.setPlaceholderText("Repita la contraseña")
        self.usu_clave2.setEchoMode(QLineEdit.EchoMode.Password)
        self.usu_clave2.setToolTip("Escriba exactamente la misma contraseña.")
        form.addRow(etiqueta_campo("Nombre completo"), self.usu_nombre)
        form.addRow(etiqueta_campo("Usuario"), self.usu_usuario)
        form.addRow(etiqueta_campo("Puesto"), self.usu_rol)
        form.addRow(etiqueta_campo("Contraseña"), self.usu_clave)
        form.addRow(etiqueta_campo("Repetir contraseña"), self.usu_clave2)
        btn_guardar = QPushButton("Crear usuario")
        btn_guardar.setObjectName("primario")
        btn_guardar.setToolTip("Revisa los datos y pide confirmación antes de guardar.")
        btn_guardar.clicked.connect(self.guardar_usuario)
        btn_guardar.setMinimumWidth(170)
        fila_boton = QHBoxLayout()
        fila_boton.addStretch()
        fila_boton.addWidget(btn_guardar)
        panel_reg.layout_principal.addLayout(form)
        panel_reg.layout_principal.addLayout(fila_boton)
        panel_reg.layout_principal.addStretch()

        panel_list = PanelContenedor("Usuarios del sistema",
                                     "Solo el gerente puede crear, desactivar o restablecer cuentas.")
        busqueda = QHBoxLayout()
        self.usu_buscar = QLineEdit()
        self.usu_buscar.setPlaceholderText("Buscar por usuario o nombre")
        self.usu_buscar.addAction(icono_buscar(), QLineEdit.ActionPosition.LeadingPosition)
        self.usu_buscar.setClearButtonEnabled(True)
        self.usu_buscar.returnPressed.connect(self.actualizar_tabla_usuarios)
        btn_buscar = QPushButton("Buscar")
        btn_buscar.setObjectName("secundario")
        btn_buscar.clicked.connect(self.actualizar_tabla_usuarios)
        busqueda.addWidget(self.usu_buscar)
        busqueda.addWidget(btn_buscar)

        self.tabla_usuarios = tabla_solo_lectura(5)
        self.tabla_usuarios.setHorizontalHeaderLabels(["ID", "Usuario", "Nombre", "Puesto", "Estado"])
        ajustar_columnas(self.tabla_usuarios, {
            0: QHeaderView.ResizeMode.ResizeToContents,
            1: QHeaderView.ResizeMode.ResizeToContents,
            2: QHeaderView.ResizeMode.Stretch,
            3: QHeaderView.ResizeMode.ResizeToContents,
            4: QHeaderView.ResizeMode.ResizeToContents,
        })
        self.vacio_usuarios = QLabel("No hay usuarios que coincidan con la búsqueda.")
        self.vacio_usuarios.setObjectName("vacio")
        self.vacio_usuarios.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.conectar_vacio(self.tabla_usuarios, self.vacio_usuarios,
                            "No hay usuarios que coincidan con la búsqueda.")

        acciones = QHBoxLayout()
        self.btn_usu_clave = QPushButton("Restablecer clave")
        self.btn_usu_clave.setObjectName("secundario")
        self.btn_usu_clave.setToolTip("Asigna una contraseña nueva a la cuenta seleccionada.")
        self.btn_usu_clave.clicked.connect(self.restablecer_clave)
        self.btn_usu_estado = QPushButton("Desactivar")
        self.btn_usu_estado.setObjectName("peligro")
        self.btn_usu_estado.setToolTip("Impide que la cuenta seleccionada pueda entrar al sistema.")
        self.btn_usu_estado.clicked.connect(self.cambiar_estado_usuario)
        self.tabla_usuarios.itemSelectionChanged.connect(self.actualizar_boton_usuarios)
        self.btn_usu_eliminar = QPushButton("Eliminar")
        self.btn_usu_eliminar.setObjectName("peligro")
        self.btn_usu_eliminar.setToolTip("Elimina definitivamente una cuenta inactiva. No se puede deshacer.")
        self.btn_usu_eliminar.clicked.connect(self.eliminar_usuario)
        acciones.addStretch()
        acciones.addWidget(self.btn_usu_clave)
        acciones.addWidget(self.btn_usu_estado)
        acciones.addWidget(self.btn_usu_eliminar)

        panel_list.layout_principal.addLayout(busqueda)
        panel_list.layout_principal.addWidget(self.tabla_usuarios, 1)
        panel_list.layout_principal.addWidget(self.vacio_usuarios)
        panel_list.layout_principal.addLayout(acciones)
        layout.addWidget(panel_reg, 1)
        layout.addWidget(panel_list, 2)
        self.vistas.addWidget(vista)

    def setup_movimientos(self):
        vista = QWidget()
        layout = QVBoxLayout(vista)
        layout.setContentsMargins(32, 22, 32, 24)
        layout.setSpacing(16)
        panel = PanelContenedor(
            "Movimientos del sistema",
            "Registro de todo lo que se hizo en el programa: quién, con qué puesto, cuándo, "
            "qué acción y sobre qué registro. Es la misma información que consulta el panel de administración.")
        filtros = QHBoxLayout()
        self.mov_buscar = QLineEdit()
        self.mov_buscar.setPlaceholderText("Buscar por usuario, acción o detalle")
        self.mov_buscar.addAction(icono_buscar(), QLineEdit.ActionPosition.LeadingPosition)
        self.mov_buscar.setClearButtonEnabled(True)
        self.mov_buscar.returnPressed.connect(self.actualizar_tabla_movimientos)
        self.mov_periodo = QComboBox()
        for etiqueta, dias in (("Hoy", 1), ("Últimos 7 días", 7), ("Últimos 30 días", 30),
                               ("Últimos 90 días", 90), ("Todo el historial", 0)):
            self.mov_periodo.addItem(etiqueta, dias)
        self.mov_periodo.setCurrentIndex(2)
        self.mov_rol = QComboBox()
        self.mov_rol.addItem("Todos los puestos", "")
        for codigo in ROLES_SISTEMA:
            self.mov_rol.addItem(ROLES_TEXTO.get(codigo, codigo.capitalize()), codigo)
        btn_buscar = QPushButton("Buscar")
        btn_buscar.setObjectName("secundario")
        btn_buscar.clicked.connect(self.actualizar_tabla_movimientos)
        filtros.addWidget(self.mov_buscar, 1)
        filtros.addWidget(self.mov_periodo)
        filtros.addWidget(self.mov_rol)
        filtros.addWidget(btn_buscar)

        self.tabla_movimientos = tabla_solo_lectura(6)
        self.tabla_movimientos.setHorizontalHeaderLabels(
            ["Fecha", "Usuario", "Puesto", "Acción", "Detalle", "Importe"])
        ajustar_columnas(self.tabla_movimientos, {
            0: QHeaderView.ResizeMode.ResizeToContents,
            1: QHeaderView.ResizeMode.ResizeToContents,
            2: QHeaderView.ResizeMode.ResizeToContents,
            3: QHeaderView.ResizeMode.ResizeToContents,
            4: QHeaderView.ResizeMode.Stretch,
            5: QHeaderView.ResizeMode.ResizeToContents,
        })
        self.vacio_movimientos = QLabel("No hay movimientos que coincidan con el filtro.")
        self.vacio_movimientos.setObjectName("vacio")
        self.vacio_movimientos.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.conectar_vacio(self.tabla_movimientos, self.vacio_movimientos,
                            "No hay movimientos que coincidan con el filtro.")
        self.lbl_mov_resumen = QLabel("")
        self.lbl_mov_resumen.setObjectName("ayuda")
        panel.layout_principal.addLayout(filtros)
        panel.layout_principal.addWidget(self.tabla_movimientos, 1)
        panel.layout_principal.addWidget(self.vacio_movimientos)
        panel.layout_principal.addWidget(self.lbl_mov_resumen)
        layout.addWidget(panel, 1)
        self.vistas.addWidget(vista)

    @requiere(PERMISO_GERENTE)
    @protegido
    @con_espera
    def actualizar_tabla_movimientos(self):
        texto = self.mov_buscar.text().strip()
        dias = int(self.mov_periodo.currentData() or 0)
        rol = str(self.mov_rol.currentData() or "")
        consulta = ["SELECT fecha, usuario, rol, accion, detalle, importe FROM movimiento WHERE 1 = 1"]
        valores = []
        if texto:
            consulta.append("AND (usuario LIKE ? OR accion LIKE ? OR COALESCE(detalle, '') LIKE ?)")
            patron = f"%{texto}%"
            valores.extend([patron, patron, patron])
        if dias:
            consulta.append("AND fecha >= date('now', 'localtime', ?)")
            valores.append(f"-{dias - 1} days")
        if rol:
            consulta.append("AND UPPER(TRIM(rol)) = ?")
            valores.append(rol.upper())
        consulta.append("ORDER BY idMovimiento DESC LIMIT 500")
        con = conectar()
        try:
            filas = con.execute(" ".join(consulta), valores).fetchall()
        finally:
            con.close()
        self.tabla_movimientos.setRowCount(0)
        for f in filas:
            row = self.tabla_movimientos.rowCount()
            self.tabla_movimientos.insertRow(row)
            puesto = str(f["rol"] or "").strip().upper()
            poner_celda(self.tabla_movimientos, row, 0, fecha_legible(str(f["fecha"] or "")[:10]))
            poner_celda(self.tabla_movimientos, row, 1, f["usuario"] or "-")
            poner_celda(self.tabla_movimientos, row, 2, ROLES_TEXTO.get(puesto, puesto.capitalize() or "-"))
            poner_celda(self.tabla_movimientos, row, 3, f["accion"] or "-")
            poner_celda(self.tabla_movimientos, row, 4, f["detalle"] or "-")
            importe = f["importe"]
            poner_celda(self.tabla_movimientos, row, 5,
                        f"${float(importe):,.2f}" if importe is not None else "")
        alinear_tabla(self.tabla_movimientos, (0, 2, 5), izquierda=(1, 3, 4), derecha=(5,))
        self.lbl_mov_resumen.setText(
            f"Movimientos mostrados: {len(filas)}. Se conservan todos los registros; "
            "aquí solo se muestran los últimos 500 del filtro.")

    def setup_facturas(self):
        vista = QWidget()
        layout = QVBoxLayout(vista)
        layout.setContentsMargins(32, 22, 32, 24)
        layout.setSpacing(16)
        panel = PanelContenedor(
            "Facturas emitidas",
            "Cada factura usa el importe de la cotización validada del servicio. "
            "No se puede facturar dos veces el mismo servicio.")
        filtros = QHBoxLayout()
        self.fac_buscar = QLineEdit()
        self.fac_buscar.setPlaceholderText("Buscar por factura, cliente o servicio")
        self.fac_buscar.addAction(icono_buscar(), QLineEdit.ActionPosition.LeadingPosition)
        self.fac_buscar.setClearButtonEnabled(True)
        self.fac_buscar.returnPressed.connect(self.actualizar_tabla_facturas)
        btn_buscar = QPushButton("Buscar")
        btn_buscar.setObjectName("secundario")
        btn_buscar.clicked.connect(self.actualizar_tabla_facturas)
        filtros.addWidget(self.fac_buscar, 1)
        filtros.addWidget(btn_buscar)
        self.tabla_facturas = tabla_solo_lectura(8)
        self.tabla_facturas.setHorizontalHeaderLabels(
            ["Factura", "Fecha", "Cliente", "Servicio", "Subtotal", "IVA", "Total", "Estado"])
        ajustar_columnas(self.tabla_facturas, {
            0: QHeaderView.ResizeMode.ResizeToContents,
            1: QHeaderView.ResizeMode.ResizeToContents,
            2: QHeaderView.ResizeMode.Stretch,
            3: QHeaderView.ResizeMode.ResizeToContents,
            4: QHeaderView.ResizeMode.ResizeToContents,
            5: QHeaderView.ResizeMode.ResizeToContents,
            6: QHeaderView.ResizeMode.ResizeToContents,
            7: QHeaderView.ResizeMode.ResizeToContents,
        })
        self.vacio_facturas = QLabel("Todavía no hay facturas emitidas.")
        self.vacio_facturas.setObjectName("vacio")
        self.vacio_facturas.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.conectar_vacio(self.tabla_facturas, self.vacio_facturas, "Todavía no hay facturas emitidas.")
        self.lbl_fac_resumen = QLabel("")
        self.lbl_fac_resumen.setObjectName("seccion")
        panel.layout_principal.addLayout(filtros)
        panel.layout_principal.addWidget(self.tabla_facturas, 1)
        panel.layout_principal.addWidget(self.vacio_facturas)
        panel.layout_principal.addWidget(self.lbl_fac_resumen)
        layout.addWidget(panel, 1)
        self.vistas.addWidget(vista)

    @requiere(PERMISO_FACTURACION)
    @protegido
    @con_espera
    def actualizar_tabla_facturas(self):
        texto = self.fac_buscar.text().strip()
        patron = f"%{texto}%"
        con = conectar()
        try:
            filas = con.execute(
                """SELECT f.numeroFactura, f.fechaEmision, c.nombre AS cliente,
                          f.idServicio, f.subtotal, f.iva, f.total, f.estado
                   FROM factura f JOIN cliente c ON c.idCliente = f.idCliente
                   WHERE f.numeroFactura LIKE ? OR c.nombre LIKE ? OR CAST(f.idServicio AS TEXT) = ?
                   ORDER BY f.idFactura DESC;""", (patron, patron, texto)).fetchall()
        finally:
            con.close()
        self.tabla_facturas.setRowCount(0)
        total_general = 0.0
        for f in filas:
            row = self.tabla_facturas.rowCount()
            self.tabla_facturas.insertRow(row)
            poner_celda(self.tabla_facturas, row, 0, f["numeroFactura"] or "-")
            poner_celda(self.tabla_facturas, row, 1, fecha_legible(f["fechaEmision"]))
            poner_celda(self.tabla_facturas, row, 2, f["cliente"])
            poner_celda(self.tabla_facturas, row, 3, f"#{f['idServicio']}")
            poner_celda(self.tabla_facturas, row, 4, f"${float(f['subtotal'] or 0):,.2f}")
            poner_celda(self.tabla_facturas, row, 5, f"${float(f['iva'] or 0):,.2f}")
            poner_celda(self.tabla_facturas, row, 6, f"${float(f['total'] or 0):,.2f}")
            poner_celda(self.tabla_facturas, row, 7, str(f["estado"] or "GENERADA").title())
            total_general += float(f["total"] or 0)
        alinear_tabla(self.tabla_facturas, (0, 1, 3, 7), izquierda=(2,), derecha=(4, 5, 6))
        self.lbl_fac_resumen.setText(
            f"Facturas: {len(filas)}    Total facturado: ${total_general:,.2f}")

    @protegido
    def cargar_roles_listado(self):
        con = conectar()
        try:
            roles = con.execute("SELECT idRol, nombreRol FROM rol ORDER BY idRol;").fetchall()
        finally:
            con.close()
        ROLES_LISTA.clear()
        ROLES_LISTA.extend((r["idRol"], r["nombreRol"]) for r in roles)
        self.emp_roles.clear()
        for idRol, nombreRol in ROLES_LISTA:
            item = QListWidgetItem(ROLES_TEXTO.get(nombreRol, nombreRol.capitalize()))
            item.setData(Qt.ItemDataRole.UserRole, idRol)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Unchecked)
            self.emp_roles.addItem(item)

    @con_espera
    def ejecutar_db(self, funcion, *args, exito=None):
        con = conectar()
        try:
            resultado = funcion(con, *args)
        except ReglaError as e:
            aviso(self, str(e))
            return False, None
        except sqlite3.IntegrityError as e:
            fallo(self, describir_error_bd(e) or "No se pudo guardar porque un dato ya existe o no cumple una regla.",
                  SUGERENCIA_DB, str(e))
            return False, None
        except sqlite3.OperationalError as e:
            if "locked" in str(e).lower() or "busy" in str(e).lower():
                aviso(self, MSG_OCUPADA)
            else:
                fallo(self, describir_error_bd(e) or MSG_GENERAL, SUGERENCIA_DB, str(e))
            return False, None
        except OSError as e:
            fallo(self, MSG_ARCHIVO, SUGERENCIA_ARCHIVO, str(e))
            return False, None
        except sqlite3.Error as e:
            fallo(self, describir_error_bd(e) or MSG_GENERAL, SUGERENCIA_DB, str(e))
            return False, None
        except Exception as e:
            fallo(self, MSG_GENERAL, SUGERENCIA_INESPERADO, str(e))
            return False, None
        finally:
            con.close()
        if exito:
            listo(self, exito)
        return True, resultado

    @requiere(PERMISO_VENTA)
    @protegido
    def guardar_cliente(self):
        datos = {
            "nombre": self.cli_nombre.text().strip(),
            "telefono": self.cli_tel.text().strip() or None,
            "correo": self.cli_correo.text().strip() or None,
            "direccion": self.cli_dir.text().strip() or None,
            "rfc": self.cli_rfc.text().strip().upper() or None,
        }
        campos = {"nombre": self.cli_nombre, "telefono": self.cli_tel,
                  "correo": self.cli_correo, "direccion": self.cli_dir, "rfc": self.cli_rfc}
        for widget in campos.values():
            marcar_error(widget, False)
        errores = errores_cliente(datos)
        if errores:
            aviso(self, "Corrija los siguientes datos:\n\n" + texto_errores(errores))
            marcar_errores(errores, campos)
            return
        con = conectar()
        try:
            existente = con.execute("SELECT idCliente FROM cliente WHERE lower(nombre) = lower(?) AND activo = 1;",
                                    (datos["nombre"],)).fetchone()
        finally:
            con.close()
        if existente:
            if not preguntar(self, "Cliente repetido",
                             "Ya existe un cliente con ese nombre.\n\n¿Desea registrarlo de todas formas?",
                             "Sí, registrar", "No, cancelar"):
                return
        resumen = (f"Nombre: {datos['nombre']}\nTeléfono: {datos['telefono'] or '-'}\n"
                   f"Correo: {datos['correo'] or '-'}\nDirección: {datos['direccion'] or '-'}\n"
                   f"RFC: {datos['rfc'] or '-'}")
        if not confirmar(self, "Confirmar cliente", "¿Desea guardar este cliente?\n\n" + resumen):
            return
        ok, idNuevo = self.ejecutar_db(registrar_cliente_db, datos)
        if not ok:
            return
        listo(self, f"El cliente se guardó correctamente. Número de cliente: {idNuevo}")
        for campo in (self.cli_nombre, self.cli_tel, self.cli_correo, self.cli_dir, self.cli_rfc):
            campo.clear()
        self.actualizar_tabla_clientes()
        self.cargar_clientes_combo()

    @con_espera
    def actualizar_tabla_clientes(self, filas=None):
        seleccionado = id_de_fila(self.tabla_clientes)
        if filas is None:
            con = conectar()
            try:
                filas = con.execute("SELECT * FROM cliente WHERE activo = 1 ORDER BY nombre;").fetchall()
            finally:
                con.close()
        self.tabla_clientes.setRowCount(0)
        for fila in filas:
            row = self.tabla_clientes.rowCount()
            self.tabla_clientes.insertRow(row)
            poner_celda(self.tabla_clientes, row, 0, fila["idCliente"])
            poner_celda(self.tabla_clientes, row, 1, fila["nombre"])
            poner_celda(self.tabla_clientes, row, 2, fila["telefono"] or "-")
            poner_celda(self.tabla_clientes, row, 3, fila["correo"] or "-")
        alinear_tabla(self.tabla_clientes, (0, 2), izquierda=(1, 3))
        restaurar_seleccion(self.tabla_clientes, seleccionado)
        self.statusBar().showMessage(f"Clientes activos: {len(filas)}")

    @con_espera
    def buscar_clientes(self):
        texto = self.cli_buscar.text().strip()
        con = conectar()
        try:
            filas = con.execute("SELECT * FROM cliente WHERE activo = 1 AND nombre LIKE ? ORDER BY nombre;",
                                (f"%{texto}%",)).fetchall()
        finally:
            con.close()
        self.actualizar_tabla_clientes(filas)

    def cliente_seleccionado(self):
        idCliente = id_de_fila(self.tabla_clientes)
        if idCliente is None:
            aviso(self, "Primero seleccione un cliente de la lista.")
        return idCliente

    @requiere(PERMISO_VENTA)
    @protegido
    def modificar_cliente(self):
        idCliente = self.cliente_seleccionado()
        if idCliente is None:
            return
        con = conectar()
        try:
            datos = con.execute("SELECT * FROM cliente WHERE idCliente = ?;", (idCliente,)).fetchone()
        finally:
            con.close()
        if datos is None:
            aviso(self, MSG_REGISTRO)
            self.buscar_clientes()
            return
        dlg = DialogoCliente(self, "Modificar cliente", datos)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        nuevos = dlg.valores()
        resumen = (f"Nombre: {nuevos['nombre']}\nTeléfono: {nuevos['telefono'] or '-'}\n"
                   f"RFC: {nuevos['rfc'] or '-'}")
        if not confirmar(self, "Confirmar cambios", "¿Desea guardar los cambios de este cliente?\n\n" + resumen):
            return
        ok, _ = self.ejecutar_db(modificar_cliente_db, idCliente, nuevos, exito="Los cambios se guardaron.")
        if ok:
            self.buscar_clientes()
            self.cargar_clientes_combo()

    @requiere(PERMISO_VENTA)
    @protegido
    def eliminar_cliente(self):
        idCliente = self.cliente_seleccionado()
        if idCliente is None:
            return
        nombre = texto_de_fila(self.tabla_clientes, 1)
        if not confirmar(self, "Dar de baja cliente",
                         f"¿Desea dar de baja a {nombre}?\n\nSu historial se conserva y su número queda reservado."):
            return
        ok, _ = self.ejecutar_db(eliminar_cliente_db, idCliente, exito=f"{nombre} fue dado de baja.")
        if ok:
            self.buscar_clientes()
            self.cargar_clientes_combo()

    @requiere(PERMISO_GERENTE)
    @protegido
    def guardar_empleado(self):
        datos = {
            "nombre": self.emp_nombre.text().strip(),
            "ap": self.emp_ap.text().strip(),
            "am": self.emp_am.text().strip() or None,
            "telefono": self.emp_tel.text().strip() or None,
            "correo": self.emp_correo.text().strip() or None,
            "roles": [self.emp_roles.item(i).data(Qt.ItemDataRole.UserRole)
                      for i in range(self.emp_roles.count())
                      if self.emp_roles.item(i).checkState() == Qt.CheckState.Checked
                      or self.emp_roles.item(i).isSelected()],
        }
        campos = {"nombre": self.emp_nombre, "ap": self.emp_ap, "am": self.emp_am,
                  "telefono": self.emp_tel, "correo": self.emp_correo, "roles": self.emp_roles}
        for widget in campos.values():
            marcar_error(widget, False)
        errores = errores_empleado(datos)
        if errores:
            aviso(self, "Corrija los siguientes datos:\n\n" + texto_errores(errores))
            marcar_errores(errores, campos)
            return
        puestos = ", ".join(self.emp_roles.item(i).text()
                            for i in range(self.emp_roles.count())
                            if self.emp_roles.item(i).checkState() == Qt.CheckState.Checked
                            or self.emp_roles.item(i).isSelected())
        resumen = (f"Nombre: {datos['nombre']} {datos['ap']} {datos['am'] or ''}\n"
                   f"Teléfono: {datos['telefono'] or '-'}\nCorreo: {datos['correo'] or '-'}\n"
                   f"Puestos: {puestos}")
        if not confirmar(self, "Confirmar empleado", "¿Desea guardar este empleado?\n\n" + resumen):
            return
        ok, idNuevo = self.ejecutar_db(registrar_empleado_db, datos)
        if not ok:
            return
        listo(self, f"El empleado se guardó correctamente. Número: {idNuevo}")
        for campo in (self.emp_nombre, self.emp_ap, self.emp_am, self.emp_tel, self.emp_correo):
            campo.clear()
        self.emp_roles.clearSelection()
        for indice in range(self.emp_roles.count()):
            self.emp_roles.item(indice).setCheckState(Qt.CheckState.Unchecked)
        self.actualizar_tabla_empleados()

    @con_espera
    def actualizar_tabla_empleados(self):
        seleccionado = id_de_fila(self.tabla_empleados)
        texto = self.emp_buscar.text().strip()
        con = conectar()
        try:
            filas = con.execute("""
                SELECT e.idEmpleado, e.nombre, e.aPaterno, e.telefono,
                       COALESCE(GROUP_CONCAT(r.nombreRol, ', '), 'Sin puesto') AS puestos
                FROM empleado e
                LEFT JOIN puesto p ON p.idEmpleado = e.idEmpleado
                LEFT JOIN rol r ON r.idRol = p.idRol
                WHERE e.activo = 1
                  AND (e.nombre LIKE ? OR e.aPaterno LIKE ? OR COALESCE(e.aMaterno, '') LIKE ?)
                GROUP BY e.idEmpleado
                ORDER BY e.nombre;
            """, (f"%{texto}%", f"%{texto}%", f"%{texto}%")).fetchall()
        finally:
            con.close()
        self.tabla_empleados.setRowCount(0)
        for f in filas:
            row = self.tabla_empleados.rowCount()
            self.tabla_empleados.insertRow(row)
            poner_celda(self.tabla_empleados, row, 0, f["idEmpleado"])
            poner_celda(self.tabla_empleados, row, 1, f"{f['nombre']} {f['aPaterno']}")
            poner_celda(self.tabla_empleados, row, 2, f["telefono"] or "-")
            poner_celda(self.tabla_empleados, row, 3, texto_roles(f["puestos"]))
        alinear_tabla(self.tabla_empleados, (0, 2), izquierda=(1, 3))
        restaurar_seleccion(self.tabla_empleados, seleccionado)
        self.statusBar().showMessage(f"Empleados activos: {len(filas)}")

    def empleado_seleccionado(self):
        idEmpleado = id_de_fila(self.tabla_empleados)
        if idEmpleado is None:
            aviso(self, "Primero seleccione un empleado de la lista.")
        return idEmpleado

    @requiere(PERMISO_GERENTE)
    @protegido
    def modificar_empleado(self):
        idEmpleado = self.empleado_seleccionado()
        if idEmpleado is None:
            return
        con = conectar()
        try:
            datos = con.execute("SELECT * FROM empleado WHERE idEmpleado = ?;", (idEmpleado,)).fetchone()
            roles_actuales = [r[0] for r in con.execute(
                "SELECT idRol FROM puesto WHERE idEmpleado = ?;", (idEmpleado,)).fetchall()]
        finally:
            con.close()
        if datos is None:
            aviso(self, MSG_REGISTRO)
            self.actualizar_tabla_empleados()
            return
        dlg = DialogoEmpleado(self, "Modificar empleado", datos, roles_actuales)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        nuevos = dlg.valores()
        nombres = ", ".join(n for i, n in ROLES_LISTA if i in nuevos["roles"])
        resumen = f"Nombre: {nuevos['nombre']} {nuevos['ap']}\nPuestos: {nombres}"
        if not confirmar(self, "Confirmar cambios", "¿Desea guardar los cambios de este empleado?\n\n" + resumen):
            return
        ok, _ = self.ejecutar_db(modificar_empleado_db, idEmpleado, nuevos, exito="Los cambios se guardaron.")
        if ok:
            self.actualizar_tabla_empleados()

    @requiere(PERMISO_GERENTE)
    @protegido
    def eliminar_empleado(self):
        idEmpleado = self.empleado_seleccionado()
        if idEmpleado is None:
            return
        nombre = texto_de_fila(self.tabla_empleados, 1)
        if not confirmar(self, "Dar de baja empleado",
                         f"¿Desea dar de baja a {nombre}?\n\nSu historial se conserva y su número queda reservado."):
            return
        ok, _ = self.ejecutar_db(eliminar_empleado_db, idEmpleado, exito=f"{nombre} fue dado de baja.")
        if ok:
            self.actualizar_tabla_empleados()

    @protegido
    def resolver_cliente(self, texto):
        con = conectar()
        try:
            coincidencias = buscar_clientes(con, texto)
        finally:
            con.close()
        if not coincidencias:
            aviso(self, "No encontramos un cliente con ese nombre o número.\n"
                        "Revise que esté bien escrito o regístrelo primero en la sección Clientes.")
            return None
        if len(coincidencias) == 1:
            idCliente, nombre = coincidencias[0][0], coincidencias[0][1]
            if nombre.strip().lower() != texto.strip().lower():
                if not confirmar(self, "Confirmar cliente",
                                 f"El único cliente que coincide es:\n\n{nombre} (ID {idCliente})\n\n"
                                 "¿Es el cliente correcto?"):
                    return None
            return idCliente
        dlg = DialogoSeleccion(self, "Varios clientes con ese nombre",
                               "Elija el cliente correcto:",
                               [f"Cliente {c[0]} - {c[1]}" for c in coincidencias])
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return None
        return coincidencias[dlg.indice()][0]

    @requiere(PERMISO_VENTA)
    @protegido
    def guardar_servicio(self):
        cli_busq = self.srv_cliente.currentText().strip()
        datos = {
            "problema": self.srv_problema.toPlainText().strip(),
            "descripcion": self.srv_descripcion.toPlainText().strip() or None,
            "observaciones": self.srv_obs.toPlainText().strip() or None,
        }
        if not cli_busq:
            aviso(self, "Escriba el nombre o el número del cliente.")
            return
        marcar_error(self.srv_problema, False)
        errores = errores_servicio(datos)
        if errores:
            aviso(self, texto_errores(errores))
            marcar_errores(errores, {"problema": self.srv_problema})
            return
        indice = self.srv_cliente.currentIndex()
        if indice >= 0 and self.srv_cliente.itemText(indice).strip() == cli_busq:
            idCliente = self.srv_cliente.itemData(indice)
        else:
            idCliente = self.resolver_cliente(cli_busq)
        if idCliente is None:
            return
        con = conectar()
        try:
            nombre_cliente = con.execute("SELECT nombre FROM cliente WHERE idCliente = ?;", (idCliente,)).fetchone()[0]
        finally:
            con.close()
        resumen = (f"Cliente: {nombre_cliente}\nProblema: {datos['problema']}\n"
                   f"Descripción del servicio: {datos['descripcion'] or '-'}\n"
                   f"Observaciones: {datos['observaciones'] or '-'}")
        if not confirmar(self, "Confirmar servicio", "¿Desea registrar este servicio?\n\n" + resumen):
            return
        ok, folio = self.ejecutar_db(registrar_servicio_db, idCliente, datos["problema"],
                                     datos["descripcion"], datos["observaciones"])
        if not ok:
            return
        listo(self, f"El servicio se registró correctamente. Número de servicio: {folio}")
        self.srv_cliente.setCurrentIndex(-1)
        self.srv_cliente.setEditText("")
        self.srv_problema.clear()
        self.srv_descripcion.clear()
        self.srv_obs.clear()
        self.actualizar_tabla_servicios()

    @con_espera
    def actualizar_tabla_servicios(self):
        seleccionado = id_de_fila(self.tabla_servicios)
        texto = self.srv_buscar.text().strip()
        con = conectar()
        try:
            filas = con.execute("""
                SELECT s.idServicio, s.fechaRegistro, c.nombre AS cliente,
                       es.nombreEstado, s.descripcionProblema,
                       (SELECT co.idCotizacion FROM cotizacion co
                         WHERE co.idServicio = s.idServicio
                         ORDER BY co.idCotizacion DESC LIMIT 1) AS idCotizacion,
                       (SELECT co.total FROM cotizacion co
                         WHERE co.idServicio = s.idServicio
                         ORDER BY co.idCotizacion DESC LIMIT 1) AS totalCotizacion
                FROM servicio s
                JOIN cliente c ON c.idCliente = s.idCliente
                JOIN estado_servicio es ON es.idEstado = s.idEstado
                WHERE c.nombre LIKE ? OR s.descripcionProblema LIKE ?
                ORDER BY s.idServicio DESC;
            """, (f"%{texto}%", f"%{texto}%")).fetchall()
        finally:
            con.close()
        self.tabla_servicios.setRowCount(0)
        for s in filas:
            row = self.tabla_servicios.rowCount()
            self.tabla_servicios.insertRow(row)
            poner_celda(self.tabla_servicios, row, 0, s["idServicio"])
            poner_celda(self.tabla_servicios, row, 1, fecha_legible(s["fechaRegistro"]))
            poner_celda(self.tabla_servicios, row, 2, s["cliente"])
            poner_celda(self.tabla_servicios, row, 3, texto_estado(s["nombreEstado"]))
            poner_celda(self.tabla_servicios, row, 4, s["descripcionProblema"])
            if s["idCotizacion"]:
                poner_celda(self.tabla_servicios, row, 5,
                            f"COT-{int(s['idCotizacion']):05d} · ${float(s['totalCotizacion'] or 0):,.2f}")
            else:
                poner_celda(self.tabla_servicios, row, 5, "Sin cotización")
        alinear_tabla(self.tabla_servicios, (0, 1, 3), izquierda=(2, 4), derecha=(5,))
        restaurar_seleccion(self.tabla_servicios, seleccionado)
        self.statusBar().showMessage(f"Servicios registrados: {len(filas)}")

    def llenar_menu_acciones(self):
        self.menu_acciones.clear()
        idServicio = id_de_fila(self.tabla_servicios)
        if idServicio is None:
            vacio = self.menu_acciones.addAction("Seleccione un servicio de la lista")
            vacio.setEnabled(False)
            return
        con = conectar()
        try:
            estado = estado_de(con, idServicio)
        except ReglaError:
            estado = None
        finally:
            con.close()
        opciones = [("Ver cotizaciones", self.accion_ver_cotizaciones)]
        if estado in ("REGISTRADO", "ASIGNADO") and puede(PERMISO_VENTA):
            opciones = [("Asignar personal", self.accion_asignar),
                        ("Generar cotización", self.accion_cotizar)] + opciones
        if estado == "COTIZACION_GENERADA" and puede(PERMISO_VENTA):
            opciones.append(("Volver a cotizar", self.accion_cotizar))
            opciones.append(("Aceptar cotización", self.accion_aceptar))
        if estado in ("COTIZACION_GENERADA", "ACEPTADO", "EN_PROCESO", "CORRECCION_SOLICITADA") \
                and puede(PERMISO_SUPERVISION):
            opciones.insert(0, ("Asignar personal", self.accion_asignar))
        if estado == "ACEPTADO" and puede(PERMISO_SUPERVISION):
            opciones.append(("Iniciar trabajo", self.accion_iniciar))
        if estado in ("EN_PROCESO", "CORRECCION_SOLICITADA") and puede(PERMISO_SUPERVISION):
            opciones.append(("Enviar a revisión", self.accion_enviar))
        if estado == "PENDIENTE_REVISION" and puede(PERMISO_SUPERVISION):
            opciones.append(("Revisar servicio", self.accion_revisar))
        if estado == "APROBADO" and puede(PERMISO_FACTURACION):
            opciones.append(("Facturar", self.accion_facturar))
        if estado == "FACTURADO" and puede(PERMISO_FACTURACION):
            opciones.append(("Finalizar servicio", self.finalizar_servicio))
        if estado in ESTADOS_RETROALIMENTACION and puede(PERMISO_SUPERVISION):
            opciones.append(("Enviar comentario al personal", self.retroalimentar_servicio))
        for texto, funcion in opciones:
            accion = self.menu_acciones.addAction(texto)
            accion.triggered.connect(lambda _=False, f=funcion: f())

    @protegido
    def accion_ver_cotizaciones(self):
        idServicio = self.servicio_seleccionado()
        if idServicio is None:
            return
        DialogoCotizaciones(self, idServicio).exec()

    @requiere(PERMISO_VENTA)
    @protegido
    def accion_asignar(self):
        idServicio = self.servicio_seleccionado()
        if idServicio is None:
            return
        con = conectar()
        try:
            empleados = empleados_activos(con)
        finally:
            con.close()
        if not empleados:
            aviso(self, "No hay personal registrado. Pida al gerente que registre al personal.")
            return
        dlg = DialogoAsignar(self, [(e[0], e[1]) for e in empleados])
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        asignaciones = dlg.valores()
        nombres = {e[0]: e[1] for e in empleados}
        lista = "\n".join(f"{nombres[i]} ({p})" for i, p in asignaciones)
        if not confirmar(self, "Confirmar asignación", f"¿Desea asignar al servicio {idServicio}?\n\n{lista}"):
            return
        ok, agregados = self.ejecutar_db(asignar_personal_db, idServicio, asignaciones)
        if ok:
            if agregados:
                listo(self, f"Personal asignado: {plural(agregados, 'persona', 'personas')}. "
                            "El servicio pasa a Asignado.")
            else:
                aviso(self, "Esas personas ya estaban asignadas a este servicio.")
            self.actualizar_tabla_servicios()

    @requiere(PERMISO_VENTA)
    @protegido
    def accion_cotizar(self):
        idServicio = self.servicio_seleccionado()
        if idServicio is None:
            return
        con = conectar()
        try:
            empleados = empleados_activos(con)
            tasa = obtener_iva(con)
            info = con.execute(
                """SELECT c.nombre AS cliente, s.descripcionProblema AS problema,
                          es.nombreEstado AS estado
                   FROM servicio s
                   JOIN cliente c ON c.idCliente = s.idCliente
                   JOIN estado_servicio es ON es.idEstado = s.idEstado
                   WHERE s.idServicio = ?;""", (idServicio,)).fetchone()
        finally:
            con.close()
        if info is None:
            aviso(self, MSG_REGISTRO)
            return
        cliente = info["cliente"]
        problema = info["problema"]
        re_cotizacion = info["estado"] == "COTIZACION_GENERADA"
        if not empleados:
            aviso(self, "No hay personal registrado. Pida al gerente que registre al personal.")
            return
        dlg = DialogoCotizacion(self, [(e[0], e[1]) for e in empleados], tasa,
                                idServicio, cliente, problema, re_cotizacion)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        idEmpleado, conceptos = dlg.valores()
        if not conceptos:
            aviso(self, "Agregue al menos un concepto a la cotización.")
            return
        subtotal, iva, total = calcular_totales(conceptos, tasa)
        lineas = "\n".join(
            f"{d} ({u}): {c:g} x ${p:,.2f} = ${i:,.2f}" for d, u, c, p, i in conceptos)
        resumen = (f"{lineas}\n\nSubtotal: ${subtotal:,.2f}\nIVA {tasa * 100:g}%: ${iva:,.2f}\n"
                   f"Total: ${total:,.2f}")
        titulo = "Confirmar nueva cotización" if re_cotizacion else "Confirmar cotización"
        if not confirmar(self, titulo, f"¿Desea generar la cotización?\n\n{resumen}"):
            return
        ok, res = self.ejecutar_db(cotizar_db, idServicio, idEmpleado, conceptos)
        if not ok:
            return
        idCot, ruta, total = res
        if ruta and os.path.exists(ruta):
            if preguntar(self, "Cotización generada",
                         f"Se generó la cotización COT-{int(idCot):05d} por ${total:,.2f}.\n\n"
                         "¿Desea abrir el archivo Excel ahora?",
                         "Sí, abrir", "No, cerrar", por_defecto_no=False):
                abrir_archivo(ruta)
        else:
            aviso(self, f"La cotización COT-{int(idCot):05d} se registró, pero no se encontró el archivo Excel.")
        self.actualizar_tabla_servicios()

    @requiere(PERMISO_VENTA)
    @protegido
    def accion_aceptar(self):
        idServicio = self.servicio_seleccionado()
        if idServicio is None:
            return
        if not confirmar(self, "Aceptar cotización",
                         f"¿El cliente aceptó la cotización del servicio {idServicio}?"):
            return
        ok, _ = self.ejecutar_db(aceptar_cotizacion_db, idServicio, exito="Cotización aceptada.")
        if ok:
            self.actualizar_tabla_servicios()

    @requiere(PERMISO_SUPERVISION)
    @protegido
    def accion_iniciar(self):
        idServicio = self.servicio_seleccionado()
        if idServicio is None:
            return
        if not confirmar(self, "Iniciar trabajo", f"¿Desea iniciar el trabajo del servicio {idServicio}?"):
            return
        ok, _ = self.ejecutar_db(iniciar_trabajo_db, idServicio, exito="El trabajo ya está en proceso.")
        if ok:
            self.actualizar_tabla_servicios()

    @requiere(PERMISO_SUPERVISION)
    @protegido
    def accion_enviar(self):
        idServicio = self.servicio_seleccionado()
        if idServicio is None:
            return
        if not confirmar(self, "Enviar a revisión", f"¿Enviar el servicio {idServicio} a revisión?"):
            return
        ok, _ = self.ejecutar_db(enviar_revision_db, idServicio, exito="El servicio fue enviado a revisión.")
        if ok:
            self.actualizar_tabla_servicios()

    @requiere(PERMISO_SUPERVISION)
    @protegido
    def accion_revisar(self):
        idServicio = self.servicio_seleccionado()
        if idServicio is None:
            return
        con = conectar()
        try:
            supervisores = empleados_activos(con, ROLES_REVISOR)
        finally:
            con.close()
        if not supervisores:
            aviso(self, "No hay supervisores registrados para revisar el servicio.")
            return
        dlg = DialogoRevision(self, [(e[0], e[1]) for e in supervisores])
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        idSup, aprobado, comentarios = dlg.valores()
        decision = "aprobar el servicio" if aprobado else "solicitar una corrección"
        if not confirmar(self, "Confirmar revisión", f"¿Desea {decision} del servicio {idServicio}?"):
            return
        ok, nuevo = self.ejecutar_db(revisar_servicio_db, idServicio, idSup, aprobado, comentarios)
        if ok:
            listo(self, "Revisión guardada. El servicio fue aprobado." if aprobado
                  else "Revisión guardada. Se solicitó corrección al personal.")
            self.actualizar_tabla_servicios()

    @requiere(PERMISO_FACTURACION)
    @protegido
    def accion_facturar(self):
        idServicio = self.servicio_seleccionado()
        if idServicio is None:
            return
        con = conectar()
        try:
            cot = con.execute(
                """SELECT idCotizacion, subtotal, iva, total, estado
                   FROM cotizacion WHERE idServicio = ? ORDER BY idCotizacion DESC LIMIT 1;""",
                (idServicio,)).fetchone()
        finally:
            con.close()
        if cot is None:
            aviso(self, "Este servicio no tiene cotización. Genere la cotización antes de facturar.")
            return
        estado_cot = str(cot["estado"] or "REGISTRADA").strip().upper()
        resumen = (f"Cotización COT-{int(cot['idCotizacion']):05d}\n"
                   f"Subtotal: ${float(cot['subtotal'] or 0):,.2f}\n"
                   f"IVA: ${float(cot['iva'] or 0):,.2f}\n"
                   f"Total a facturar: ${float(cot['total'] or 0):,.2f}\n"
                   f"Estado de la cotización: {estado_cot.title()}\n\n"
                   "Al facturar se valida la cotización y se usa su importe.\n"
                   "¿Desea continuar?")
        if not confirmar(self, "Facturar servicio", resumen):
            return
        ok, res = self.ejecutar_db(facturar_db, idServicio)
        if ok:
            numero, total, folio_cot = res
            listo(self, f"Factura {numero} generada por ${total:,.2f}.\n\n"
                        f"Se usó el importe de la cotización {folio_cot}, que quedó validada.\n"
                        "El servicio pasa a Facturado.")
            self.actualizar_tabla_servicios()

    def servicio_seleccionado(self):
        idServicio = id_de_fila(self.tabla_servicios)
        if idServicio is None:
            aviso(self, "Primero seleccione un servicio de la lista.")
        return idServicio

    @requiere(PERMISO_VENTA)
    @protegido
    def modificar_servicio(self):
        idServicio = self.servicio_seleccionado()
        if idServicio is None:
            return
        con = conectar()
        try:
            datos = con.execute("SELECT * FROM servicio WHERE idServicio = ?;", (idServicio,)).fetchone()
        finally:
            con.close()
        if datos is None:
            aviso(self, MSG_REGISTRO)
            self.actualizar_tabla_servicios()
            return
        dlg = DialogoServicio(self, datos)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        nuevos = dlg.valores()
        if not confirmar(self, "Confirmar cambios",
                         f"¿Desea guardar los cambios del servicio {idServicio}?\n\nProblema: {nuevos['problema']}"):
            return
        ok, _ = self.ejecutar_db(modificar_servicio_db, idServicio, nuevos["problema"],
                                 nuevos["descripcion"], nuevos["observaciones"], exito="Los cambios se guardaron.")
        if ok:
            self.actualizar_tabla_servicios()

    @requiere(PERMISO_FACTURACION)
    @protegido
    def finalizar_servicio(self):
        idServicio = self.servicio_seleccionado()
        if idServicio is None:
            return
        if not confirmar(self, "Finalizar servicio", f"¿Desea finalizar el servicio {idServicio}?"):
            return
        ok, _ = self.ejecutar_db(finalizar_servicio_db, idServicio, exito="El servicio fue finalizado.")
        if ok:
            self.actualizar_tabla_servicios()

    @requiere(PERMISO_SUPERVISION)
    @protegido
    def retroalimentar_servicio(self):
        idServicio = self.servicio_seleccionado()
        if idServicio is None:
            return
        con = conectar()
        try:
            participantes = participantes_servicio(con, idServicio)
        finally:
            con.close()
        if not participantes:
            aviso(self, "Este servicio todavía no tiene personal asignado.")
            return
        dlg = DialogoSeleccion(self, "Comentario al personal", "Elija al trabajador:",
                               [f"{p[1]} - {p[2]}" for p in participantes])
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        idEmpleado, nombre = participantes[dlg.indice()][0], participantes[dlg.indice()][1]
        dlg_texto = DialogoTexto(self, "Comentario al personal", "Comentario",
                                 "El trabajador verá este comentario en el reporte del servicio.",
                                 multilinea=True)
        if dlg_texto.exec() != QDialog.DialogCode.Accepted:
            return
        texto = dlg_texto.texto()
        if not texto_valido(texto):
            aviso(self, "Escriba un comentario con al menos 3 letras.")
            return
        if not confirmar(self, "Confirmar comentario", f"Trabajador: {nombre}\n\n{texto}"):
            return
        self.ejecutar_db(retroalimentar_db, idServicio, idEmpleado, texto,
                         exito="El comentario se guardó.")

    @requiere(PERMISO_GERENTE)
    @protegido
    def cancelar_servicio(self):
        idServicio = self.servicio_seleccionado()
        if idServicio is None:
            return
        dlg = DialogoTexto(self, "Cancelar servicio", "Motivo de la cancelación",
                           "El servicio se conserva en el historial como cancelado.", multilinea=True)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        motivo = dlg.texto()
        if not texto_valido(motivo):
            aviso(self, "Escriba un motivo con al menos 3 letras.")
            return
        if not confirmar(self, "Cancelar servicio", f"¿Desea cancelar el servicio {idServicio}?\n\nMotivo: {motivo}"):
            return
        ok, _ = self.ejecutar_db(cancelar_servicio_db, idServicio, motivo, exito="El servicio fue cancelado.")
        if ok:
            self.actualizar_tabla_servicios()


    @con_espera
    def actualizar_tabla_usuarios(self):
        seleccionado = id_de_fila(self.tabla_usuarios)
        texto = self.usu_buscar.text().strip()
        con = conectar()
        try:
            filas = con.execute(
                """SELECT idUsuario, nombreUsuario, nombreCompleto, rol, activo
                   FROM usuario
                   WHERE nombreUsuario LIKE ? OR COALESCE(nombreCompleto, '') LIKE ?
                   ORDER BY idUsuario;""",
                (f"%{texto}%", f"%{texto}%")).fetchall()
        finally:
            con.close()
        self.tabla_usuarios.setRowCount(0)
        for f in filas:
            row = self.tabla_usuarios.rowCount()
            self.tabla_usuarios.insertRow(row)
            item = poner_celda(self.tabla_usuarios, row, 0, f["idUsuario"])
            item.setData(Qt.ItemDataRole.UserRole, 1 if f["activo"] else 0)
            poner_celda(self.tabla_usuarios, row, 1, f["nombreUsuario"])
            poner_celda(self.tabla_usuarios, row, 2, f["nombreCompleto"] or "")
            rol = str(f["rol"] or "").strip().upper()
            poner_celda(self.tabla_usuarios, row, 3, ROLES_TEXTO.get(rol, rol.capitalize()))
            poner_celda(self.tabla_usuarios, row, 4, "Activo" if f["activo"] else "Inactivo")
        alinear_tabla(self.tabla_usuarios, (0, 4), izquierda=(1, 2, 3))
        restaurar_seleccion(self.tabla_usuarios, seleccionado)
        self.actualizar_boton_usuarios()
        self.statusBar().showMessage(f"Usuarios registrados: {len(filas)}")

    def actualizar_boton_usuarios(self):
        fila = self.tabla_usuarios.currentRow()
        if fila < 0:
            self.btn_usu_estado.setText("Desactivar")
            self.btn_usu_eliminar.setEnabled(False)
            self.btn_usu_eliminar.setToolTip("Seleccione una cuenta inactiva para eliminarla.")
            return
        item = self.tabla_usuarios.item(fila, 0)
        activo = True
        if item is not None:
            activo = bool(item.data(Qt.ItemDataRole.UserRole))
        self.btn_usu_estado.setText("Desactivar" if activo else "Activar")
        self.btn_usu_eliminar.setEnabled(not activo)
        self.btn_usu_eliminar.setToolTip(
            "Borra la cuenta inactiva. No se puede deshacer."
            if not activo else "Primero desactive la cuenta para poder eliminarla.")

    def usuario_seleccionado(self):
        idUsuario = id_de_fila(self.tabla_usuarios)
        if idUsuario is None:
            aviso(self, "Primero seleccione un usuario de la lista.")
            return None
        fila = self.tabla_usuarios.currentRow()
        item = self.tabla_usuarios.item(fila, 0)
        activo = bool(item.data(Qt.ItemDataRole.UserRole)) if item is not None else True
        return idUsuario, texto_de_fila(self.tabla_usuarios, 1), activo

    @requiere(PERMISO_GERENTE)
    @protegido
    def guardar_usuario(self):
        datos = {
            "nombre": self.usu_nombre.text().strip(),
            "usuario": self.usu_usuario.text().strip(),
            "rol": self.usu_rol.currentData(),
            "clave": self.usu_clave.text(),
            "clave2": self.usu_clave2.text(),
        }
        campos = {"nombre": self.usu_nombre, "usuario": self.usu_usuario, "rol": self.usu_rol,
                  "clave": self.usu_clave, "clave2": self.usu_clave2}
        for widget in campos.values():
            marcar_error(widget, False)
        errores = errores_usuario(datos)
        if errores:
            aviso(self, "Corrija los siguientes datos:\n\n" + texto_errores(errores))
            marcar_errores(errores, campos)
            return
        con = conectar()
        try:
            existente = con.execute("SELECT 1 FROM usuario WHERE nombreUsuario = ?;",
                                    (datos["usuario"],)).fetchone()
        finally:
            con.close()
        if existente:
            aviso(self, "Ese usuario ya existe. Elija otro nombre de usuario.")
            return
        resumen = (f"Nombre: {datos['nombre']}\nUsuario: {datos['usuario']}\n"
                   f"Puesto: {ROLES_TEXTO.get(datos['rol'], datos['rol'])}")
        if not confirmar(self, "Confirmar usuario", "¿Desea crear esta cuenta?\n\n" + resumen):
            return
        ok, idNuevo = self.ejecutar_db(registrar_usuario_db, datos)
        if not ok:
            return
        listo(self, f"El usuario {datos['usuario']} se creó correctamente.")
        for campo in (self.usu_nombre, self.usu_usuario, self.usu_clave, self.usu_clave2):
            campo.clear()
        self.usu_rol.setCurrentIndex(0)
        self.actualizar_tabla_usuarios()

    @requiere(PERMISO_GERENTE)
    @protegido
    def cambiar_estado_usuario(self):
        seleccion = self.usuario_seleccionado()
        if seleccion is None:
            return
        idUsuario, nombreUsuario, activo = seleccion
        if idUsuario == SESION.get("id"):
            aviso(self, "No puede desactivar la cuenta con la que está trabajando.")
            return
        if activo:
            if not confirmar(self, "Desactivar usuario",
                             f"¿Desea desactivar la cuenta {nombreUsuario}?\n\n"
                             "La cuenta dejará de poder entrar al sistema y conservará su historial y su número."):
                return
            ok, _ = self.ejecutar_db(cambiar_estado_usuario_db, idUsuario, False,
                                     exito=f"La cuenta {nombreUsuario} fue desactivada.")
        else:
            if not confirmar(self, "Activar usuario",
                             f"¿Desea activar la cuenta {nombreUsuario}?"):
                return
            ok, _ = self.ejecutar_db(cambiar_estado_usuario_db, idUsuario, True,
                                     exito=f"La cuenta {nombreUsuario} fue activada.")
        if ok:
            self.actualizar_tabla_usuarios()

    @requiere(PERMISO_GERENTE)
    @protegido
    def eliminar_usuario(self):
        seleccion = self.usuario_seleccionado()
        if seleccion is None:
            return
        idUsuario, nombreUsuario, activo = seleccion
        if activo:
            aviso(self, "Primero desactive la cuenta y después elimínela definitivamente.")
            return
        if idUsuario == SESION.get("id"):
            aviso(self, "No puede eliminar la cuenta con la que está trabajando.")
            return
        if not confirmar(self, "Eliminar cuenta definitivamente",
                         f"¿Desea borrar la cuenta {nombreUsuario}?\n\n"
                         "Esta acción no se puede deshacer. El historial de acciones se conserva."):
            return
        ok, _ = self.ejecutar_db(eliminar_usuario_db, idUsuario,
                                 exito=f"La cuenta {nombreUsuario} se eliminó definitivamente.")
        if ok:
            self.actualizar_tabla_usuarios()

    @requiere(PERMISO_GERENTE)
    @protegido
    def restablecer_clave(self):
        seleccion = self.usuario_seleccionado()
        if seleccion is None:
            return
        idUsuario, nombreUsuario, _activo = seleccion
        dlg = DialogoClave(self, nombreUsuario)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        clave, _clave2 = dlg.valores()
        if not confirmar(self, "Restablecer contraseña",
                         f"¿Desea asignar una contraseña nueva a la cuenta {nombreUsuario}?"):
            return
        ok, _ = self.ejecutar_db(cambiar_clave_db, idUsuario, clave,
                                 exito=f"La contraseña de {nombreUsuario} se actualizó.")
        if ok:
            self.actualizar_tabla_usuarios()


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setApplicationName("The Gonzalez")
    app.setApplicationDisplayName("The Gonzalez - Gestión de servicios")
    configurar_ruta_base()
    fuente = QFont()
    fuente.setFamilies(["Segoe UI", "Inter", "Helvetica Neue", "Arial", "Noto Sans"])
    fuente.setPointSize(10)
    app.setFont(fuente)
    app.setStyleSheet(QSS)
    try:
        con = conectar()
        try:
            asegurar_esquema(con)
            respaldo_diario(con)
        except (sqlite3.Error, OSError):
            pass
        hay_gerente = con.execute(
            "SELECT 1 FROM usuario WHERE UPPER(TRIM(rol)) = 'GERENTE' AND activo = 1;").fetchone() is not None
    except sqlite3.Error:
        fallo(None, MSG_BD)
        return
    finally:
        try:
            con.close()
        except (sqlite3.Error, UnboundLocalError):
            pass
    problemas = []
    ruta_bd = ruta_base_datos()
    if ruta_bd and not base_escribible(ruta_bd):
        problemas.append("La base de datos no permite guardar cambios. Cierre otras ventanas del "
                         "programa, quite la marca de solo lectura del archivo y evite carpetas "
                         "sincronizadas en la nube.")
    if not carpeta_escribible(carpeta_programa()):
        problemas.append("La carpeta del programa no permite guardar cambios. Mueva el programa "
                         "y la base de datos a Documentos o al Escritorio.")
    if problemas:
        aviso(None, "\n\n".join(problemas) +
              "\n\nPuede consultar la información, pero no guardar. "
              "Abra Ayuda > Diagnóstico para ver el detalle.")

    try:
        QSettings("TheGonzalez", "Servicios").setValue("sistema/ruta_bd", ruta_base_datos())
    except Exception:
        pass

    usuario_inicial = ""
    if not hay_gerente:
        primer_uso = DialogoPrimerUso()
        if primer_uso.exec() != QDialog.DialogCode.Accepted:
            return
        usuario_inicial = primer_uso.cuenta or ""
    try:
        con = conectar()
        try:
            asegurar_esquema(con)
        finally:
            con.close()
    except sqlite3.Error:
        pass

    acceso = DialogoAcceso(usuario_inicial)
    if acceso.exec() != QDialog.DialogCode.Accepted:
        return
    SESION.update(acceso.sesion)
    marcar_conectado(True)
    ventana = VentanaPrincipal()
    ajustes = QSettings("TheGonzalez", "Servicios")
    geometria = ajustes.value("ventana/geometria")
    pantalla = app.primaryScreen()
    if geometria:
        ventana.restoreGeometry(geometria)
        if pantalla is not None and app.screenAt(ventana.frameGeometry().center()) is None:
            ventana.resize(1180, 720)
            marco = ventana.frameGeometry()
            marco.moveCenter(pantalla.availableGeometry().center())
            ventana.move(marco.topLeft())
    elif pantalla is not None:
        marco = ventana.frameGeometry()
        marco.moveCenter(pantalla.availableGeometry().center())
        ventana.move(marco.topLeft())
    ventana.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
