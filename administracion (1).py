import re

from nucleo import (
    Cancelado,
    NOMBRES_ROL, ROLES_PUESTO, SESION, TASAS_IVA,
    bitacora, cambiar_mi_clave, captura, conectar, crear_usuario, ejecutar, elegir_de, iniciar_sesion,
    imprimir_menu, imprimir_tabla, leer, nombre_valido, obtener_iva, pedir_clave_nueva, pedir_rol,
    pedir_usuario_nuevo, respaldar, respaldo_diario, si_no, telefono_valido, correo_valido,
)


def hay_gerente(con):
    return con.execute("SELECT 1 FROM usuario WHERE rol = 'GERENTE' AND activo = 1;").fetchone() is not None











def configurar_gerente(con):
    print("\n=== CONFIGURACION INICIAL ===")
    print("Todavia no hay un gerente activo registrado.")
    print("El gerente administra usuarios, empleados, IVA y respaldos.")
    print("Responda las preguntas siguientes y guarde bien la clave.\n")
    nombre = leer("Nombre completo del gerente (ej. Juan Perez Lopez): ", nombre_valido,
                  "Escriba solo letras, sin numeros ni signos.")
    usuario = pedir_usuario_nuevo(con)
    clave = pedir_clave_nueva()
    crear_usuario(con, usuario, nombre, "GERENTE", clave)
    print(f"\nListo. Cuenta creada para {nombre}. Entre con el usuario {usuario}.")


def nuevo_usuario(con):
    nombre = leer("Nombre completo: ", nombre_valido, "Escriba solo letras, sin numeros ni signos.")
    usuario = pedir_usuario_nuevo(con)
    print("Rol del usuario:")
    rol = pedir_rol()
    clave = pedir_clave_nueva()
    idNuevo = crear_usuario(con, usuario, nombre, rol, clave)
    print(f"Usuario {usuario} creado (ID {idNuevo}, rol {NOMBRES_ROL[rol]}).")





def listar_usuarios(con):
    filas = con.execute(
        """SELECT idUsuario, nombreCompleto, nombreUsuario, rol,
                  CASE WHEN activo = 1 THEN 'SI' ELSE 'NO' END, creado
           FROM usuario ORDER BY idUsuario;""").fetchall()
    imprimir_tabla(["ID", "NOMBRE", "USUARIO", "ROL", "ACTIVO", "CREADO"],
                   [(f[0], f[1], f[2], NOMBRES_ROL.get(f[3], f[3]), f[4], f[5]) for f in filas])


def otros_gerentes_activos(con, idExcluir):
    return con.execute(
        "SELECT COUNT(*) FROM usuario WHERE rol = 'GERENTE' AND activo = 1 AND idUsuario != ?;",
        (idExcluir,)).fetchone()[0]


def usuarios_activos(con):
    return con.execute(
        "SELECT idUsuario, nombreCompleto, nombreUsuario, rol FROM usuario WHERE activo = 1 ORDER BY idUsuario;"
    ).fetchall()


def cambiar_rol(con):
    filas = [(f[0], f[1], f[2], NOMBRES_ROL.get(f[3], f[3])) for f in usuarios_activos(con)]
    idUsuario = elegir_de(filas, ["ID", "NOMBRE", "USUARIO", "ROL"], "ID del usuario: ")
    if idUsuario is None:
        return
    if idUsuario == SESION["id"]:
        print("No puede cambiar su propio rol.")
        return
    actual = con.execute("SELECT rol FROM usuario WHERE idUsuario = ?;", (idUsuario,)).fetchone()["rol"]
    print("Nuevo rol:")
    nuevo = pedir_rol()
    if actual == "GERENTE" and nuevo != "GERENTE" and otros_gerentes_activos(con, idUsuario) == 0:
        print("Debe quedar al menos un gerente activo.")
        return
    with con:
        con.execute("UPDATE usuario SET rol = ? WHERE idUsuario = ?;", (nuevo, idUsuario))
        bitacora(con, "CAMBIAR ROL", "usuario", idUsuario, f"{actual} -> {nuevo}")
    print(f"Rol actualizado a {NOMBRES_ROL[nuevo]}.")


def desactivar_usuario(con):
    filas = [(f[0], f[1], f[2], NOMBRES_ROL.get(f[3], f[3])) for f in usuarios_activos(con)]
    idUsuario = elegir_de(filas, ["ID", "NOMBRE", "USUARIO", "ROL"], "ID del usuario a desactivar: ")
    if idUsuario is None:
        return
    if idUsuario == SESION["id"]:
        print("No puede desactivarse a si mismo.")
        return
    rol = con.execute("SELECT rol FROM usuario WHERE idUsuario = ?;", (idUsuario,)).fetchone()["rol"]
    if rol == "GERENTE" and otros_gerentes_activos(con, idUsuario) == 0:
        print("Debe quedar al menos un gerente activo.")
        return
    if not si_no("Desactivar este usuario? Sus registros se conservan. (s/n): "):
        print("Operacion cancelada.")
        return
    with con:
        con.execute("UPDATE usuario SET activo = 0 WHERE idUsuario = ?;", (idUsuario,))
        bitacora(con, "DESACTIVAR USUARIO", "usuario", idUsuario)
    print("Usuario desactivado.")


def registrar_empleado(con):
    nombre = leer("Nombre: ", nombre_valido, "Nombre no valido.")
    ap = leer("Apellido paterno: ", nombre_valido, "Apellido no valido.")
    am = leer("Apellido materno (Enter para omitir): ", nombre_valido, "Apellido no valido.", opcional=True)
    telefono = leer("Telefono (10 digitos, Enter para omitir): ", telefono_valido, "Telefono invalido.", opcional=True)
    if telefono:
        telefono = re.sub(r"[\s\-()]", "", telefono)
    correo = leer("Correo (Enter para omitir): ", correo_valido, "Correo invalido.", opcional=True)
    print("Puestos del empleado:")
    for i, r in enumerate(ROLES_PUESTO, 1):
        print(f"  {i}. {r}")
    while True:
        partes = [p.strip() for p in captura("Numeros de puestos separados por coma: ").split(",")]
        if all(p.isascii() and p.isdigit() and 1 <= int(p) <= len(ROLES_PUESTO) for p in partes):
            break
        print("Seleccion no valida.")
    puestos = sorted({int(p) for p in partes})
    if not si_no("Guardar empleado? (s/n): "):
        print("Registro cancelado.")
        return
    with con:
        cur = con.execute(
            "INSERT INTO empleado (nombre, aPaterno, aMaterno, telefono, correo) VALUES (?, ?, ?, ?, ?);",
            (nombre, ap, am, telefono, correo))
        idEmpleado = cur.lastrowid
        for n in puestos:
            id_rol = con.execute("SELECT idRol FROM rol WHERE nombreRol = ?;", (ROLES_PUESTO[n - 1],)).fetchone()[0]
            con.execute("INSERT INTO puesto (idEmpleado, idRol) VALUES (?, ?);", (idEmpleado, id_rol))
        bitacora(con, "REGISTRAR EMPLEADO", "empleado", idEmpleado, f"{nombre} {ap}")
    print(f"Empleado guardado. ID: {idEmpleado}")


def listar_empleados(con):
    filas = con.execute(
        """SELECT e.idEmpleado, e.nombre || ' ' || e.aPaterno, COALESCE(GROUP_CONCAT(r.nombreRol, ', '), '-')
           FROM empleado e
           LEFT JOIN puesto p ON p.idEmpleado = e.idEmpleado
           LEFT JOIN rol r ON r.idRol = p.idRol
           WHERE e.activo = 1
           GROUP BY e.idEmpleado
           ORDER BY e.nombre;""").fetchall()
    imprimir_tabla(["ID", "NOMBRE", "PUESTOS"], [tuple(f) for f in filas])


def configurar_iva(con):
    actual = obtener_iva(con)
    print(f"IVA actual: {actual * 100:g}%. Tasas permitidas: 0, 8 o 16.")
    print("Las cotizaciones anteriores conservan la tasa con la que se hicieron.")
    while True:
        v = captura("Nuevo IVA (0, 8 o 16): ").strip()
        if v in TASAS_IVA:
            break
        print("Solo se permite 0, 8 o 16.")
    nuevo = int(v) / 100
    if not si_no(f"Cambiar el IVA a {int(v)}%? (s/n): "):
        print("Sin cambios.")
        return
    with con:
        con.execute("UPDATE configuracion SET valor = ? WHERE clave = 'IVA';", (str(nuevo),))
        bitacora(con, "CAMBIAR IVA", "configuracion", None, f"{actual} -> {nuevo}")
    print(f"IVA actualizado a {int(v)}%.")


def respaldo_manual(con):
    ruta = respaldar(con)
    with con:
        bitacora(con, "RESPALDO MANUAL", detalle=ruta)
    print(f"Respaldo creado y verificado: {ruta}")


def ver_bitacora(con):
    filas = con.execute(
        """SELECT idBitacora, fecha, usuario, accion, tabla, idRegistro, detalle
           FROM bitacora ORDER BY idBitacora DESC LIMIT 50;""").fetchall()
    imprimir_tabla(["ID", "FECHA", "USUARIO", "ACCION", "TABLA", "REG", "DETALLE"], [tuple(f) for f in filas])


OPCIONES = {



    "1": ("Crear usuario", nuevo_usuario, None, "dar acceso a una persona"),

    "2": ("Listar usuarios", listar_usuarios, None, "ver cuentas y roles"),

    "3": ("Cambiar rol", cambiar_rol, None, "cambiar el puesto en el sistema"),

    "4": ("Desactivar usuario", desactivar_usuario, None, "quitar acceso sin borrar"),

    "5": ("Registrar empleado", registrar_empleado, None, "alta de personal de campo"),


    "6": ("Listar empleados", listar_empleados, None, "ver personal activo"),


    "7": ("Configurar IVA", configurar_iva, None, "tasa de IVA: 0, 8 o 16"),


    "8": ("Respaldar base", respaldo_manual, None, "copia de seguridad "),

    "9": ("Ver bitacora", ver_bitacora, None, "historial de cambios"),

    "10": ("Cambiar mi clave", cambiar_mi_clave, None, "cambiar tu clave de acceso"),
}


def mostrar_menu():
    imprimir_menu("ADMINISTRACION - THE GONZALEZ", OPCIONES, lambda _: True)


def main():
    con = conectar()
    try:
        respaldo_diario(con)
        if not hay_gerente(con):
            configurar_gerente(con)
        if not iniciar_sesion(con, rol_requerido="GERENTE"):
            print("No se pudo iniciar sesion. Saliendo.")
            return
        while True:
            mostrar_menu()
            op = input("Opcion: ").strip()
            if op == "0":
                break
            if op not in OPCIONES:
                print("Opcion no valida. Escriba un numero del menu.")
                continue
            ejecutar(OPCIONES[op][1], con)
    except Cancelado:
        print("\nOperacion cancelada. Saliendo.")
    except (EOFError, KeyboardInterrupt):
        print("\nSaliendo.")
    finally:
        con.close()


if __name__ == "__main__":
    main()
