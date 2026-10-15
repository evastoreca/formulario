# ==============================================================================
#  REGISTRO DE APLICACIONES DE EVALUACIONES
#  Shiny + base de datos (SQLite local / PostgreSQL en la nube)
# ------------------------------------------------------------------------------
#  Qué hace
#   - Evaluadores externos ingresan con una clave, eligen colegio y registran
#     cada aplicación (colegio + grado + letra + evaluación [+ toma en Kínder]).
#   - Se guarda automáticamente un borrador; luego se envía. Un envío puede
#     corregirse después con "Actualizar envío" (queda historial de versiones),
#     pero solo si el administrador habilita la corrección (solo para quien envió).
#   - Administración: avance, revisión, generación de claves y descarga de un
#     Excel con la estructura de "Planilla de ejemplo de registro.xlsx".
#
#  Archivos necesarios en la misma carpeta
#   - app.R                   (este archivo)
#   - listado_colegios.xlsx   (instancias de aplicación)
#
#  Paquetes
#   Requiere R >= 4.1.
#   install.packages(c("shiny","bslib","DT","DBI","RSQLite","readxl",
#                      "openxlsx","digest","jsonlite"))
#   Solo para la nube (PostgreSQL): install.packages(c("pool","RPostgres"))
#
#  Uso local (pruebas): abrir el proyecto en RStudio y ejecutar shiny::runApp()
#     (o botón "Run App" con app.R abierto). Usa SQLite: datos_evaluaciones.sqlite
#     Clave de administración local: admin-local
#     Para empezar de cero: cerrar la app y borrar datos_evaluaciones.sqlite
#
#  Despliegue (nube): definir estas variables de entorno
#     DB_HOST, DB_PORT, DB_NAME, DB_USER, DB_PASSWORD   (PostgreSQL)
#     APP_PEPPER    texto largo y secreto (protege las claves guardadas)
#     ADMIN_CLAVE   clave del administrador
#   Opcionales (para reutilizar el sistema en varios procesos/semestres):
#     PROCESO          nombre visible, ej. "2026 - Semestre 2" (título y nombres de archivo)
#     ARCHIVO_LISTADO  listado de instancias de ese proceso (por defecto listado_colegios.xlsx)
#     DB_SCHEMA        esquema de PostgreSQL propio del proceso, ej. proceso_2026_s2
#     DB_SQLITE        (solo local) archivo SQLite propio del proceso
#   Con DB_HOST definido la app NO acepta valores por defecto: si faltan
#   APP_PEPPER o ADMIN_CLAVE, se niega a iniciar.
# ==============================================================================

library(shiny)
library(bslib)
library(DT)
library(DBI)
library(readxl)
library(openxlsx)
library(digest)
library(jsonlite)

`%||%` <- function(a, b) if (is.null(a) || length(a) == 0) b else a

# ------------------------------------------------------------------------------
# 1. CONFIGURACIÓN
# ------------------------------------------------------------------------------
TZ                <- "America/Santiago"
ARCHIVO_LISTADO   <- Sys.getenv("ARCHIVO_LISTADO", "listado_colegios.xlsx")
PROCESO           <- Sys.getenv("PROCESO", "")   # nombre visible del proceso, ej. "2026 - Semestre 2"
DB_SCHEMA         <- tolower(Sys.getenv("DB_SCHEMA", ""))  # solo PostgreSQL, ej. "proceso_2026_s2"
if (nzchar(DB_SCHEMA) && !grepl("^[a-z][a-z0-9_]{0,40}$", DB_SCHEMA))
  stop("DB_SCHEMA solo admite minúsculas, números y guion bajo (ej. proceso_2026_s2).")
SUFIJO_PROCESO    <- if (nzchar(PROCESO)) paste0(" · ", PROCESO) else ""
SLUG_PROCESO      <- if (nzchar(PROCESO)) paste0("_", gsub("[^A-Za-z0-9]+", "-", PROCESO)) else ""
LOCK_MINUTOS      <- 10      # un registro "en edición" se considera libre tras este tiempo
EVALUACIONES      <- c("MAT", "CL", "PAL", "KLEN", "KMAT", "CM")
ORDEN_GRADOS      <- c("K", as.character(1:8))
SITUACIONES_NEEP  <- c("Estudiante NEEP permanente rinde la evaluación",
                       "Estudiante NEEP que sí rendía, no rinde evaluación",
                       "Estudiante que originalmente no era NEEP, ahora sí lo es")
OPCIONES_OBS      <- c("Sin observaciones", "Con observaciones")
ESTADOS_REVISION  <- c("Pendiente", "Revisado", "Requiere corrección")

env <- function(x, d = "") {
  v <- Sys.getenv(x, unset = "")
  if (nzchar(v)) v else d
}
USA_POSTGRES <- nzchar(env("DB_HOST"))
ADMIN_CLAVE  <- env("ADMIN_CLAVE", if (!USA_POSTGRES) "admin-local" else "")
PEPPER       <- env("APP_PEPPER",  if (!USA_POSTGRES) "pepper-local-solo-pruebas" else "")
if (USA_POSTGRES && (!nzchar(ADMIN_CLAVE) || !nzchar(PEPPER))) {
  stop("Faltan las variables de entorno APP_PEPPER y/o ADMIN_CLAVE.")
}

ahora <- function() format(Sys.time(), "%Y-%m-%d %H:%M:%S", tz = TZ)
hoy   <- function() as.Date(format(Sys.time(), "%Y-%m-%d", tz = TZ))
fmt_fecha_hora <- function(x) {
  ifelse(is.na(x) | !nzchar(x), "", format(as.POSIXct(x, tz = TZ), "%d-%m-%Y %H:%M"))
}

# ------------------------------------------------------------------------------
# 2. LISTADO DE INSTANCIAS ("slots"): una fila por registro a completar
#    Kínder (K) tiene dos tomas -> dos registros por instancia
# ------------------------------------------------------------------------------
cargar_slots <- function() {
  x <- readxl::read_excel(ARCHIVO_LISTADO, col_types = "text")
  x <- as.data.frame(x[, 1:9], stringsAsFactors = FALSE)
  names(x) <- c("sigla", "nombre", "letra", "grado", "evaluacion",
                "modalidad", "gestora", "id_colegio", "rbd")
  x <- x[!is.na(x$sigla) & nzchar(trimws(x$sigla)), , drop = FALSE]
  x[] <- lapply(x, function(v) trimws(as.character(v)))
  x$grado <- toupper(x$grado)
  x$letra <- toupper(x$letra)
  s <- x
  s$toma <- 1L                                          # todas las instancias parten con la Toma 1
  base <- paste(s$sigla, s$grado, s$letra, s$evaluacion, sep = "-")
  s$id_registro <- ifelse(s$grado == "K", paste0(base, "-T1"), base)   # Kínder conserva el sufijo -T1
  stopifnot("IDs de registro duplicados en el listado" = !anyDuplicated(s$id_registro))
  s <- s[order(s$sigla, match(s$grado, ORDEN_GRADOS), match(s$evaluacion, EVALUACIONES),
               s$letra, s$toma), ]
  rownames(s) <- NULL
  s
}
slots <- cargar_slots()
OPCIONES_COLEGIO <- local({
  u <- unique(slots[, c("sigla", "nombre")])
  setNames(u$sigla, paste0(u$sigla, " — ", u$nombre))
})

# La Toma 2 (de cualquier curso) no está en el listado: existe solo si un evaluador la agrega.
# Su ID es el de la Toma 1 con sufijo -T2 (en Kínder -T1 pasa a -T2; en los demás se agrega -T2).
slot_por_id <- function(id) {
  s <- slots[slots$id_registro == id, , drop = FALSE]
  if (nrow(s) > 0) return(s)
  if (grepl("-T2$", id)) {
    base <- sub("-T2$", "", id)
    b <- slots[slots$id_registro %in% c(paste0(base, "-T1"), base), , drop = FALSE]
    if (nrow(b) > 0) { b <- b[1, , drop = FALSE]; b$toma <- 2L; b$id_registro <- id; return(b) }
  }
  s
}
id_toma2 <- function(id) if (grepl("-T1$", id)) sub("-T1$", "-T2", id) else paste0(id, "-T2")
id_toma1 <- function(id) {
  c1 <- sub("-T2$", "-T1", id)
  if (c1 %in% slots$id_registro) c1 else sub("-T2$", "", id)
}
# Texto " · Toma n": en Kínder siempre; en los demás cursos solo desde la Toma 2
txt_toma <- function(grado, toma) if (!is.na(toma) && (grado == "K" || toma > 1)) sprintf(" · Toma %d", toma) else ""
# Listado + las Tomas 2 que existen en la base
slots_ext <- function() {
  ids <- q("SELECT id_registro FROM registros WHERE id_registro LIKE '%-T2'")$id_registro
  if (!length(ids)) return(slots)
  t2 <- do.call(rbind, lapply(ids, slot_por_id))
  if (is.null(t2) || nrow(t2) == 0) return(slots)
  s <- rbind(slots, t2)
  s <- s[order(s$sigla, match(s$grado, ORDEN_GRADOS), match(s$evaluacion, EVALUACIONES), s$letra, s$toma), ]
  rownames(s) <- NULL
  s
}

etiqueta_curso <- function(grado, letra) ifelse(grado == "K", paste0("Kínder ", letra), paste0(grado, "° ", letra))

# ------------------------------------------------------------------------------
# 3. BASE DE DATOS
# ------------------------------------------------------------------------------
conectar <- function() {
  if (USA_POSTGRES) {
    args <- list(host = env("DB_HOST"), port = as.integer(env("DB_PORT", "5432")),
                 dbname = env("DB_NAME", "postgres"), user = env("DB_USER"),
                 password = env("DB_PASSWORD"))
    # Esquema propio por proceso: se crea y se fija como search_path en cada conexión
    fijar_esquema <- function(con) {
      if (nzchar(DB_SCHEMA)) DBI::dbExecute(con, paste0("SET search_path TO ", DB_SCHEMA))
      invisible(con)
    }
    if (nzchar(DB_SCHEMA)) {
      c0 <- do.call(DBI::dbConnect, c(list(drv = RPostgres::Postgres()), args))
      try(DBI::dbExecute(c0, paste0("CREATE SCHEMA IF NOT EXISTS ", DB_SCHEMA)), silent = FALSE)
      DBI::dbDisconnect(c0)
    }
    if (requireNamespace("pool", quietly = TRUE)) {
      do.call(pool::dbPool, c(list(drv = RPostgres::Postgres(), minSize = 1, maxSize = 8,
                                   onCreate = fijar_esquema), args))
    } else {
      fijar_esquema(do.call(DBI::dbConnect, c(list(drv = RPostgres::Postgres()), args)))
    }
  } else {
    f <- env("DB_SQLITE", "datos_evaluaciones.sqlite")
    if (requireNamespace("pool", quietly = TRUE)) {
      pool::dbPool(RSQLite::SQLite(), dbname = f)
    } else {
      DBI::dbConnect(RSQLite::SQLite(), dbname = f)
    }
  }
}
db <- conectar()
onStop(function() {
  try(if (inherits(db, "Pool")) pool::poolClose(db) else DBI::dbDisconnect(db), silent = TRUE)
})

con_run <- function(f) {
  if (inherits(db, "Pool")) {
    con <- pool::poolCheckout(db)
    on.exit(pool::poolReturn(con))
    f(con)
  } else f(db)
}
with_tx <- function(f) {
  if (inherits(db, "Pool")) pool::poolWithTransaction(db, f) else DBI::dbWithTransaction(db, f(db))
}
qc  <- function(con, sql, ...) DBI::dbGetQuery(con, DBI::sqlInterpolate(con, sql, ...))
exc <- function(con, sql, ...) DBI::dbExecute(con, DBI::sqlInterpolate(con, sql, ...))
q   <- function(sql, ...) con_run(function(con) qc(con, sql, ...))
ex  <- function(sql, ...) con_run(function(con) exc(con, sql, ...))

crear_esquema <- function() {
  sentencias <- c(
    "CREATE TABLE IF NOT EXISTS claves (
       clave_id TEXT PRIMARY KEY, evaluador TEXT, hash TEXT NOT NULL,
       activa INTEGER NOT NULL, creada TEXT, ultimo_uso TEXT)",
    "CREATE TABLE IF NOT EXISTS registros (
       id_registro TEXT PRIMARY KEY, sigla TEXT NOT NULL, grado TEXT NOT NULL,
       letra TEXT NOT NULL, evaluacion TEXT NOT NULL, toma INTEGER,
       estado TEXT NOT NULL, fecha_aplicacion TEXT, presentes INTEGER,
       ausentes_decl INTEGER, retirados_decl INTEGER,
       observaciones TEXT, detalle_obs TEXT,
       estado_revision TEXT NOT NULL, comentario_revision TEXT,
       rev INTEGER NOT NULL, creado_por TEXT, enviado_por TEXT, fecha_envio TEXT,
       actualizado_por TEXT, actualizado TEXT, habilitado_para TEXT)",
    "CREATE TABLE IF NOT EXISTS ausentes (
       id_registro TEXT NOT NULL, orden INTEGER NOT NULL, apellidos TEXT, nombre TEXT)",
    "CREATE TABLE IF NOT EXISTS retirados (
       id_registro TEXT NOT NULL, orden INTEGER NOT NULL, rut TEXT, apellidos TEXT, nombre TEXT)",
    "CREATE TABLE IF NOT EXISTS nuevos (
       id_registro TEXT NOT NULL, orden INTEGER NOT NULL, rut TEXT, apellidos TEXT, nombre TEXT, rut_pendiente TEXT)",
    "CREATE TABLE IF NOT EXISTS neep (
       id_registro TEXT NOT NULL, orden INTEGER NOT NULL, rut TEXT, apellidos TEXT, nombre TEXT, situacion TEXT)",
    "CREATE TABLE IF NOT EXISTS excepcionales (
       id_registro TEXT NOT NULL, orden INTEGER NOT NULL, apellidos TEXT, nombre TEXT, situacion TEXT)",
    "CREATE TABLE IF NOT EXISTS bloqueos (
       id_registro TEXT PRIMARY KEY, clave_id TEXT, ts TEXT)",
    "CREATE TABLE IF NOT EXISTS historial (
       id_registro TEXT NOT NULL, rev INTEGER NOT NULL, accion TEXT, clave_id TEXT,
       ts TEXT, snapshot TEXT, PRIMARY KEY (id_registro, rev))",
    "CREATE INDEX IF NOT EXISTS idx_registros_sigla ON registros (sigla)",
    "CREATE INDEX IF NOT EXISTS idx_ausentes ON ausentes (id_registro)",
    "CREATE INDEX IF NOT EXISTS idx_retirados ON retirados (id_registro)",
    "CREATE INDEX IF NOT EXISTS idx_nuevos ON nuevos (id_registro)",
    "CREATE INDEX IF NOT EXISTS idx_neep ON neep (id_registro)",
    "CREATE INDEX IF NOT EXISTS idx_excepcionales ON excepcionales (id_registro)"
  )
  con_run(function(con) {
    for (s in sentencias) DBI::dbExecute(con, s)
    # Bases creadas antes de agregar el RUT de retirados
    if (!"rut" %in% DBI::dbListFields(con, "retirados"))
      DBI::dbExecute(con, "ALTER TABLE retirados ADD COLUMN rut TEXT")
    # Bases creadas antes del RUT en NEEP y del "RUT pendiente" en alumnos nuevos
    if (!"rut" %in% DBI::dbListFields(con, "neep"))
      DBI::dbExecute(con, "ALTER TABLE neep ADD COLUMN rut TEXT")
    if (!"rut_pendiente" %in% DBI::dbListFields(con, "nuevos"))
      DBI::dbExecute(con, "ALTER TABLE nuevos ADD COLUMN rut_pendiente TEXT")
    # Bases creadas antes de la habilitación de correcciones
    if (!"habilitado_para" %in% DBI::dbListFields(con, "registros"))
      DBI::dbExecute(con, "ALTER TABLE registros ADD COLUMN habilitado_para TEXT")
  })
}
crear_esquema()

# ------------------------------------------------------------------------------
# 4. CLAVES
# ------------------------------------------------------------------------------
normaliza_clave <- function(x) toupper(gsub("[^A-Za-z0-9]", "", x %||% ""))
hash_clave <- function(x) digest::hmac(PEPPER, normaliza_clave(x), "sha256")

# Claves de 6 dígitos (100000-999999), al azar y únicas entre sí (incluye claves desactivadas)
generar_claves <- function(evaluadores) {
  n <- length(evaluadores)
  usados <- q("SELECT hash FROM claves")$hash
  elegidas <- character(); hs_el <- character()
  while (length(elegidas) < n) {
    c1 <- as.character(sample(100000:999999, 1))
    h1 <- hash_clave(c1)
    if (h1 %in% usados || h1 %in% hs_el) next
    elegidas <- c(elegidas, c1); hs_el <- c(hs_el, h1)
  }
  actuales <- q("SELECT clave_id FROM claves")$clave_id
  n0 <- if (length(actuales)) max(as.integer(sub("^E", "", actuales)), na.rm = TRUE) else 0L
  out <- data.frame(clave_id = sprintf("E%03d", n0 + seq_len(n)), evaluador = evaluadores,
                    clave = as.character(elegidas), stringsAsFactors = FALSE)
  reg <- data.frame(clave_id = out$clave_id, evaluador = out$evaluador,
                    hash = unname(vapply(out$clave, hash_clave, "")), activa = 1L,
                    creada = ahora(), ultimo_uso = NA_character_, stringsAsFactors = FALSE)
  con_run(function(con) DBI::dbAppendTable(con, "claves", reg))
  out
}

# ------------------------------------------------------------------------------
# 5. UTILIDADES (RUT, validación, normalización)
# ------------------------------------------------------------------------------
vacio <- function(x) is.null(x) || length(x) == 0 || is.na(x[1]) || !nzchar(trimws(x[1]))
na_chr <- function(x) if (vacio(x)) NA_character_ else trimws(x)
a_entero <- function(x) {
  x <- trimws(as.character(x %||% ""))
  if (!grepl("^[0-9]{1,6}$", x)) return(NA_integer_)
  as.integer(x)
}

# RUT: solo dígitos (y K), sin puntos ni guion, con dígito verificador. Ej: 123456785
rut_valido <- function(x) {
  vapply(x, function(z) {
    r <- toupper(trimws(z))
    if (!grepl("^[0-9]{6,8}[0-9K]$", r)) return(FALSE)
    cuerpo <- substr(r, 1, nchar(r) - 1)
    dv <- substr(r, nchar(r), nchar(r))
    d <- rev(as.integer(strsplit(cuerpo, "")[[1]]))
    s <- sum(d * rep(2:7, length.out = length(d)))
    e <- 11 - (s %% 11)
    dvc <- if (e == 11) "0" else if (e == 10) "K" else as.character(e)
    identical(dv, dvc)
  }, logical(1), USE.NAMES = FALSE)
}
formatear_rut <- function(x) toupper(trimws(x))

# f = lista con el contenido del formulario:
#   fecha, presentes, aus_n, ret_n, obs, detalle_obs  y data.frames ausentes, retirados,
#   nuevos, neep, excepcionales
filas_incompletas <- function(d, cols) {
  if (is.null(d) || nrow(d) == 0) return(0L)
  sum(apply(d[, cols, drop = FALSE], 1, function(r) any(!nzchar(trimws(r)))))
}

validar_form <- function(f) {
  sec <- function(msgs, resumen) list(ok = length(msgs) == 0, msg = msgs, resumen = resumen)

  m <- character()
  fe <- suppressWarnings(as.Date(f$fecha, format = "%Y-%m-%d"))
  if (vacio(f$fecha) || is.na(fe)) m <- c(m, "Datos generales: indica la fecha de aplicación.")
  else if (fe > hoy()) m <- c(m, "Datos generales: la fecha de aplicación no puede ser futura.")
  if (is.na(f$presentes)) m <- c(m, "Datos generales: indica la cantidad de estudiantes presentes (solo números).")
  general <- sec(m, "Completo")

  lista_con_conteo <- function(n, d, nombre, con_rut = FALSE) {
    m <- character()
    cols <- if (con_rut) c("rut", "apellidos", "nombre") else c("apellidos", "nombre")
    inc <- filas_incompletas(d, cols)
    k <- nrow(d) - inc
    if (is.na(n)) {
      m <- c(m, sprintf("%s: indica la cantidad (0 si no hubo; solo números).", nombre))
    } else {
      if (inc > 0) m <- c(m, sprintf("%s: completa %s en todas las filas.", nombre,
                                     if (con_rut) "RUT, apellidos y nombre" else "apellidos y nombre"))
      else if (k != n) m <- c(m, sprintf("%s: declaraste %d pero registraste %d.", nombre, n, k))
    }
    if (con_rut && inc == 0 && nrow(d) > 0 && any(!rut_valido(d$rut)))
      m <- c(m, sprintf("%s: hay un RUT inválido (escríbelo sin puntos ni guion y con dígito verificador, ej. 123456785).", nombre))
    sec(m, sprintf("%d registrados", nrow(d)))
  }
  ausentes  <- lista_con_conteo(f$aus_n, f$ausentes, "Ausentes")
  retirados <- lista_con_conteo(f$ret_n, f$retirados, "Retirados", con_rut = TRUE)

  msg_rut <- "escríbelo sin puntos ni guion y con dígito verificador, ej. 123456785"
  m <- character()
  if (nrow(f$nuevos) > 0) {
    n <- f$nuevos
    pend <- n$rut_pendiente == "1"
    if (any(!nzchar(n$apellidos) | !nzchar(n$nombre)) || any(!pend & !nzchar(n$rut)))
      m <- c(m, "Alumnos nuevos: completa RUT, apellidos y nombre en todas las filas (si aún no tienes el RUT de un estudiante, marca \"RUT pendiente\").")
    else if (any(nzchar(n$rut) & !rut_valido(n$rut)))
      m <- c(m, paste0("Alumnos nuevos: hay un RUT inválido (", msg_rut, ")."))
  }
  nuevos <- sec(m, if (nrow(f$nuevos)) sprintf("%d registrados", nrow(f$nuevos)) else "Sin registros")

  m <- character()
  if (filas_incompletas(f$neep, c("rut", "apellidos", "nombre", "situacion")) > 0)
    m <- c(m, "Situaciones NEEP: completa RUT, apellidos, nombre y situación en todas las filas.")
  else if (nrow(f$neep) > 0 && any(!rut_valido(f$neep$rut)))
    m <- c(m, paste0("Situaciones NEEP: hay un RUT inválido (", msg_rut, ")."))
  neep <- sec(m, if (nrow(f$neep)) sprintf("%d registrados", nrow(f$neep)) else "Sin registros")

  m <- character()
  if (vacio(f$obs)) m <- c(m, "Observaciones: elige \"Sin observaciones\" o \"Con observaciones\".")
  else if (identical(f$obs, "Con observaciones") && vacio(f$detalle_obs))
    m <- c(m, "Observaciones: escribe el detalle de las observaciones.")
  if (filas_incompletas(f$excepcionales, c("apellidos", "nombre", "situacion")) > 0)
    m <- c(m, "Situaciones excepcionales: completa apellidos, nombre y situación en todas las filas.")
  obs <- sec(m, "Completo")

  list(general = general, ausentes = ausentes, retirados = retirados,
       nuevos = nuevos, neep = neep, obs = obs)
}

# ------------------------------------------------------------------------------
# 6. ACCESO A DATOS DE REGISTROS
# ------------------------------------------------------------------------------
COLS_DETALLE <- list(
  ausentes      = c("apellidos", "nombre"),
  retirados     = c("rut", "apellidos", "nombre"),
  nuevos        = c("rut", "apellidos", "nombre", "rut_pendiente"),
  neep          = c("rut", "apellidos", "nombre", "situacion"),
  excepcionales = c("apellidos", "nombre", "situacion")
)

# Nombres y apellidos siempre en MAYÚSCULAS, sin espacios sobrantes (incluye tildes y ñ)
a_mayus <- function(x) {
  x <- gsub("\\s+", " ", trimws(as.character(x)))
  chartr("áéíóúüñàèìòùâêîôû", "ÁÉÍÓÚÜÑÀÈÌÒÙÂÊÎÔÛ", toupper(x))
}
mayus_nombres <- function(d) {
  if (is.null(d) || nrow(d) == 0) return(d)
  for (cn in intersect(c("apellidos", "nombre"), names(d))) d[[cn]] <- a_mayus(d[[cn]])
  d
}

cargar_registro <- function(id) {
  reg <- q("SELECT * FROM registros WHERE id_registro = ?id", id = id)
  out <- list(reg = if (nrow(reg) == 1) reg else NULL)
  for (t in names(COLS_DETALLE)) {
    out[[t]] <- q(sprintf("SELECT %s FROM %s WHERE id_registro = ?id ORDER BY orden",
                          paste(COLS_DETALLE[[t]], collapse = ", "), t), id = id)
    out[[t]] <- mayus_nombres(out[[t]])
  }
  out
}

detalle_df <- function(id, d, cols) {
  if (is.null(d) || nrow(d) == 0) return(NULL)
  data.frame(id_registro = id, orden = seq_len(nrow(d)), d[, cols, drop = FALSE],
             stringsAsFactors = FALSE, row.names = NULL)
}

# Guarda un borrador o un envío. Devuelve lista(ok, conflicto, rev, estado, ts)
guardar_registro <- function(slot, f, accion, clave_id, rev_esperada, forzar = FALSE) {
  id <- slot$id_registro
  with_tx(function(con) {
    cur <- qc(con, "SELECT * FROM registros WHERE id_registro = ?id", id = id)
    existe <- nrow(cur) == 1
    if (existe && cur$rev != rev_esperada && !forzar) {
      return(list(ok = FALSE, conflicto = TRUE, por = cur$actualizado_por, cuando = cur$actualizado))
    }
    # Un registro enviado solo lo modifica el administrador o el evaluador habilitado
    if (existe && identical(cur$estado, "enviada") && clave_id != "ADMIN" &&
        !identical(cur$habilitado_para, clave_id)) {
      return(list(ok = FALSE, conflicto = FALSE, bloqueado = TRUE,
                  error = "El registro está bloqueado. Pide al administrador que lo habilite para corregirlo."))
    }
    rev_nueva <- if (existe) as.integer(cur$rev + 1L) else 1L
    ts <- ahora()
    enviar <- identical(accion, "enviar")
    con_obs <- identical(f$obs, "Con observaciones")

    row <- data.frame(
      id_registro = id, sigla = slot$sigla, grado = slot$grado, letra = slot$letra,
      evaluacion = slot$evaluacion, toma = as.integer(slot$toma),
      estado = if (enviar) "enviada" else "borrador",
      fecha_aplicacion = na_chr(f$fecha),
      presentes = as.integer(f$presentes),
      ausentes_decl = as.integer(f$aus_n), retirados_decl = as.integer(f$ret_n),
      observaciones = na_chr(f$obs),
      detalle_obs = if (con_obs) na_chr(f$detalle_obs) else NA_character_,
      estado_revision = if ((enviar && clave_id != "ADMIN") || !existe) "Pendiente" else cur$estado_revision,
      comentario_revision = if (existe) cur$comentario_revision else NA_character_,
      rev = rev_nueva,
      creado_por = if (existe) cur$creado_por else clave_id,
      enviado_por = if (enviar && !(existe && !is.na(cur$enviado_por))) clave_id else if (existe) cur$enviado_por else NA_character_,
      fecha_envio = if (enviar && !(existe && !is.na(cur$fecha_envio))) ts else if (existe) cur$fecha_envio else NA_character_,
      actualizado_por = clave_id, actualizado = ts,
      habilitado_para = if (existe && !(enviar && clave_id != "ADMIN")) cur$habilitado_para else NA_character_,
      stringsAsFactors = FALSE)

    exc(con, "DELETE FROM registros WHERE id_registro = ?id", id = id)
    DBI::dbAppendTable(con, "registros", row)

    det <- list()
    for (t in names(COLS_DETALLE)) {
      exc(con, sprintf("DELETE FROM %s WHERE id_registro = ?id", t), id = id)
      d <- mayus_nombres(f[[t]])
      if (t %in% c("nuevos", "retirados", "neep") && !is.null(d) && nrow(d) > 0) {
        d$rut <- formatear_rut(d$rut)
        d$rut[!nzchar(d$rut)] <- NA_character_
      }
      if (t == "nuevos" && !is.null(d) && nrow(d) > 0) d$rut_pendiente[!nzchar(d$rut_pendiente)] <- NA_character_
      dd <- detalle_df(id, d, COLS_DETALLE[[t]])
      if (!is.null(dd)) DBI::dbAppendTable(con, t, dd)
      det[[t]] <- d
    }

    if (enviar) {
      snap <- jsonlite::toJSON(list(registro = as.list(row), detalle = det),
                               auto_unbox = TRUE, na = "null", dataframe = "rows")
      DBI::dbAppendTable(con, "historial",
        data.frame(id_registro = id, rev = rev_nueva,
                   accion = if (existe && cur$estado == "enviada") "actualizar" else "enviar",
                   clave_id = clave_id, ts = ts, snapshot = as.character(snap),
                   stringsAsFactors = FALSE))
    }
    list(ok = TRUE, conflicto = FALSE, rev = rev_nueva, estado = row$estado, ts = ts)
  })
}

# Elimina por completo un registro (se usa para descartar una Toma 2 en borrador)
descartar_registro <- function(slot) {
  id <- slot$id_registro
  with_tx(function(con) {
    for (t in c("registros", names(COLS_DETALLE))) exc(con, sprintf("DELETE FROM %s WHERE id_registro = ?id", t), id = id)
  })
  invisible(TRUE)
}

# Elimina definitivamente registros (solo administración): datos, detalle, historial y bloqueos
eliminar_registros <- function(ids) {
  with_tx(function(con) {
    for (id in ids)
      for (t in c("registros", names(COLS_DETALLE), "historial", "bloqueos"))
        exc(con, sprintf("DELETE FROM %s WHERE id_registro = ?id", t), id = id)
  })
  invisible(length(ids))
}

# --- Bloqueos suaves (aviso de edición simultánea) -----------------------------
lock_set <- function(id, clave_id) {
  with_tx(function(con) {
    exc(con, "DELETE FROM bloqueos WHERE id_registro = ?id", id = id)
    DBI::dbAppendTable(con, "bloqueos", data.frame(id_registro = id, clave_id = clave_id, ts = ahora(), stringsAsFactors = FALSE))
  })
}
lock_liberar <- function(id, clave_id) {
  ex("DELETE FROM bloqueos WHERE id_registro = ?id AND clave_id = ?cl", id = id, cl = clave_id)
}
lock_otro <- function(id, clave_id) {
  b <- q("SELECT b.clave_id, b.ts, c.evaluador FROM bloqueos b LEFT JOIN claves c ON c.clave_id = b.clave_id WHERE b.id_registro = ?id", id = id)
  if (nrow(b) == 0 || identical(b$clave_id, clave_id)) return(NULL)
  mins <- as.numeric(difftime(Sys.time(), as.POSIXct(b$ts, tz = TZ), units = "mins"))
  if (is.na(mins) || mins > LOCK_MINUTOS) return(NULL)
  list(quien = if (identical(b$clave_id, "ADMIN")) "Administración" else if (!vacio(b$evaluador)) b$evaluador else b$clave_id, mins = max(0, round(mins)))
}

# ------------------------------------------------------------------------------
# 7. EXPORTACIÓN A EXCEL (misma estructura que la planilla de ejemplo)
# ------------------------------------------------------------------------------
exportar_excel <- function(file, incluir_borradores = FALSE) {
  reg <- q("SELECT * FROM registros")
  cl  <- rbind(q("SELECT clave_id, evaluador FROM claves"), data.frame(clave_id = "ADMIN", evaluador = "Administración"))
  persona <- function(id) {
    ev <- cl$evaluador[match(id, cl$clave_id)]
    ifelse(is.na(id), NA_character_, ifelse(!is.na(ev) & nzchar(ev), paste0(id, " – ", ev), id))
  }
  etiqueta_estado <- c(enviada = "Enviada", borrador = "Borrador")
  S <- slots_ext()

  # --- Seguimiento (todos los registros esperados) ---
  seg <- S
  m <- match(seg$id_registro, reg$id_registro)
  est <- reg$estado[m]; est[is.na(est)] <- "pendiente"
  seguimiento <- data.frame(
    "ID aplicación" = seg$id_registro, "Sigla colegio" = seg$sigla, "Nombre colegio" = seg$nombre,
    "Evaluación" = seg$evaluacion, "Curso" = seg$grado, "Letra" = seg$letra, "Toma" = seg$toma,
    "Modalidad" = seg$modalidad, "Gestora a cargo" = seg$gestora,
    "Estado" = c(pendiente = "Pendiente", etiqueta_estado)[est],
    "Estado revisión" = reg$estado_revision[m],
    "Última actualización" = reg$actualizado[m], "Actualizado por" = persona(reg$actualizado_por[m]),
    check.names = FALSE, stringsAsFactors = FALSE)

  # --- Aplicaciones ---
  validos <- c("enviada", if (incluir_borradores) "borrador")
  reg <- reg[reg$estado %in% validos, , drop = FALSE]
  reg <- reg[order(match(reg$id_registro, S$id_registro)), , drop = FALSE]
  sl  <- S[match(reg$id_registro, S$id_registro), , drop = FALSE]
  n <- nrow(reg)

  ap <- data.frame(
    "ID aplicación" = reg$id_registro,
    "Fecha aplicación" = as.Date(reg$fecha_aplicacion),
    "Sigla colegio" = reg$sigla, "Nombre colegio" = sl$nombre,
    "Evaluación" = reg$evaluacion, "Curso" = reg$grado, "Letra" = reg$letra, "Toma" = reg$toma,
    "Modalidad" = sl$modalidad, "Gestora a cargo" = sl$gestora,
    "Presentes" = reg$presentes, "Ausentes declarados" = reg$ausentes_decl,
    "Retirados declarados" = reg$retirados_decl,
    "Observaciones" = reg$observaciones, "Detalle observaciones" = reg$detalle_obs,
    "Estado envío" = unname(etiqueta_estado[reg$estado]),
    "Estado revisión" = reg$estado_revision, "Comentario revisión" = reg$comentario_revision,
    "Ausentes registrados" = NA_integer_, "Retirados registrados" = NA_integer_,
    "Control cantidades" = NA_character_,
    "Enviado por" = persona(reg$enviado_por), "Fecha de envío" = reg$fecha_envio,
    "Última actualización" = reg$actualizado, "Actualizado por" = persona(reg$actualizado_por),
    "Versión" = reg$rev,
    check.names = FALSE, stringsAsFactors = FALSE)

  detalle <- function(tabla, cols, nombres) {
    d <- mayus_nombres(q(sprintf("SELECT * FROM %s", tabla)))
    d <- d[d$id_registro %in% reg$id_registro, , drop = FALSE]
    d <- d[order(match(d$id_registro, reg$id_registro), d$orden), , drop = FALSE]
    s <- S[match(d$id_registro, S$id_registro), , drop = FALSE]
    out <- data.frame(d$id_registro, d[, cols, drop = FALSE], s$sigla, s$evaluacion, s$grado, s$letra, s$toma,
                      check.names = FALSE, stringsAsFactors = FALSE)
    names(out) <- c("ID aplicación", nombres, "Sigla colegio", "Evaluación", "Curso", "Letra", "Toma")
    out
  }
  h_aus <- detalle("ausentes", c("apellidos", "nombre"), c("Apellidos", "Nombre"))
  h_ret <- detalle("retirados", c("rut", "apellidos", "nombre"), c("RUT", "Apellidos", "Nombre"))
  h_nue <- detalle("nuevos", c("rut", "apellidos", "nombre", "rut_pendiente"), c("RUT", "Apellidos", "Nombre", "RUT pendiente"))
  h_nue[["RUT pendiente"]] <- ifelse(!is.na(h_nue[["RUT pendiente"]]) & h_nue[["RUT pendiente"]] == "1", "Sí", "")
  h_nee <- detalle("neep", c("rut", "apellidos", "nombre", "situacion"), c("RUT", "Apellidos", "Nombre", "Situación"))
  h_exc <- detalle("excepcionales", c("apellidos", "nombre", "situacion"), c("Apellidos", "Nombre", "Situación ocurrida"))

  hist <- q("SELECT id_registro, rev, accion, clave_id, ts FROM historial ORDER BY ts")
  historial <- data.frame("ID aplicación" = hist$id_registro, "Versión" = hist$rev, "Acción" = hist$accion,
                          "Realizado por" = persona(hist$clave_id), "Fecha y hora" = hist$ts,
                          check.names = FALSE, stringsAsFactors = FALSE)

  # --- Libro ---
  wb <- createWorkbook()
  est_enc <- createStyle(textDecoration = "bold", fgFill = "#1F4E78", fontColour = "#FFFFFF",
                         halign = "center", valign = "center", wrapText = TRUE, border = "TopBottomLeftRight")
  est_fecha <- createStyle(numFmt = "dd-mm-yyyy")
  hoja <- function(nombre, df) {
    addWorksheet(wb, nombre)
    writeData(wb, nombre, df, headerStyle = est_enc, withFilter = TRUE)
    freezePane(wb, nombre, firstRow = TRUE)
    setColWidths(wb, nombre, cols = seq_len(ncol(df)), widths = "auto")
    setRowHeights(wb, nombre, rows = 1, heights = 32)
  }
  hoja("Aplicaciones", ap)
  hoja("Ausentes", h_aus); hoja("Retirados", h_ret); hoja("Nuevos", h_nue)
  hoja("NEEP", h_nee); hoja("Excepcionales", h_exc)
  hoja("Seguimiento", seguimiento)
  addWorksheet(wb, "Resumen")
  hoja("Historial", historial)

  col <- function(nombre) openxlsx::int2col(match(nombre, names(ap)))
  if (n > 0) {
    addStyle(wb, "Aplicaciones", est_fecha, rows = 2:(n + 1), cols = match("Fecha aplicación", names(ap)), stack = TRUE)
    r <- 2:(n + 1)
    writeFormula(wb, "Aplicaciones", startCol = match("Ausentes registrados", names(ap)), startRow = 2,
                 x = sprintf('IF(A%d="","",COUNTIF(Ausentes!$A$2:$A$30000,A%d))', r, r))
    writeFormula(wb, "Aplicaciones", startCol = match("Retirados registrados", names(ap)), startRow = 2,
                 x = sprintf('IF(A%d="","",COUNTIF(Retirados!$A$2:$A$30000,A%d))', r, r))
    writeFormula(wb, "Aplicaciones", startCol = match("Control cantidades", names(ap)), startRow = 2,
                 x = sprintf('IF(A%d="","",IF(AND(%s%d=%s%d,%s%d=%s%d),"OK","REVISAR"))', r,
                             col("Ausentes declarados"), r, col("Ausentes registrados"), r,
                             col("Retirados declarados"), r, col("Retirados registrados"), r))
    conditionalFormatting(wb, "Aplicaciones", cols = match("Control cantidades", names(ap)), rows = 2:(n + 1),
                          rule = '="REVISAR"', style = createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE"))
  }

  rng <- function(nombre, hoja = "Aplicaciones") sprintf("%s!%s$2:%s$30000", hoja, paste0("$", col(nombre)), paste0("$", col(nombre)))
  rs <- data.frame(
    Indicador = c("Aplicaciones registradas (en esta descarga)", "Aplicaciones enviadas", "Borradores incluidos",
                  "Aplicaciones revisadas", "Pendientes de revisión", "Requieren corrección",
                  "Inconsistencias de cantidades", "Estudiantes presentes", "Ausentes declarados",
                  "Retirados declarados", "Alumnos nuevos registrados", "Alumnos nuevos con RUT pendiente", "Situaciones NEEP registradas",
                  "Situaciones excepcionales", "Registros esperados (listado + segundas tomas agregadas)",
                  "Registros aún pendientes de envío (sin enviar ni borrador)"),
    Valor = c(sprintf('COUNTA(%s)', rng("ID aplicación")),
              sprintf('COUNTIF(%s,"Enviada")', rng("Estado envío")),
              sprintf('COUNTIF(%s,"Borrador")', rng("Estado envío")),
              sprintf('COUNTIF(%s,"Revisado")', rng("Estado revisión")),
              sprintf('COUNTIF(%s,"Pendiente")', rng("Estado revisión")),
              sprintf('COUNTIF(%s,"Requiere corrección")', rng("Estado revisión")),
              sprintf('COUNTIF(%s,"REVISAR")', rng("Control cantidades")),
              sprintf('SUM(%s)', rng("Presentes")),
              sprintf('SUM(%s)', rng("Ausentes declarados")),
              sprintf('SUM(%s)', rng("Retirados declarados")),
              'COUNTA(Nuevos!$A$2:$A$30000)', 'COUNTIF(Nuevos!$E$2:$E$30000,"Sí")', 'COUNTA(NEEP!$A$2:$A$30000)',
              'COUNTA(Excepcionales!$A$2:$A$30000)',
              'COUNTA(Seguimiento!$A$2:$A$30000)',
              'COUNTIF(Seguimiento!$J$2:$J$30000,"Pendiente")'),
    stringsAsFactors = FALSE)
  writeData(wb, "Resumen", rs["Indicador"], headerStyle = est_enc)
  writeData(wb, "Resumen", data.frame(Valor = ""), startCol = 2, headerStyle = est_enc)
  writeFormula(wb, "Resumen", x = rs$Valor, startCol = 2, startRow = 2)
  setColWidths(wb, "Resumen", cols = 1:2, widths = c(62, 16))

  # El orden de las hojas sigue la planilla de ejemplo
  worksheetOrder(wb) <- match(c("Aplicaciones", "Ausentes", "Retirados", "Nuevos", "NEEP",
                                "Excepcionales", "Seguimiento", "Resumen", "Historial"), names(wb))
  saveWorkbook(wb, file, overwrite = TRUE)
  invisible(file)
}

# ------------------------------------------------------------------------------
# 8. MÓDULO: LISTA DINÁMICA DE ESTUDIANTES (filas que se agregan / quitan)
# ------------------------------------------------------------------------------
lista_ui <- function(id) {
  ns <- NS(id)
  tagList(
    uiOutput(ns("filas")),
    div(actionButton(ns("agregar"), "+ Agregar estudiante", class = "btn btn-outline-primary btn-sm"))
  )
}

lista_server <- function(id, campos) {
  moduleServer(id, function(input, output, session) {
    ns <- session$ns
    nombres <- names(campos)
    contador <- 0
    nuevo_uid <- function() { contador <<- contador + 1; paste0("u", contador) }
    vacia <- function() {
      d <- as.data.frame(setNames(rep(list(character(0)), length(nombres)), nombres), stringsAsFactors = FALSE)
      d$uid <- character(0)
      d
    }
    filas <- reactiveVal(vacia())

    # Valores actuales: lo escrito en pantalla; si el input aún no llegó, lo cargado
    valores <- function() {
      f <- filas()
      if (nrow(f) == 0) return(f)
      for (cn in nombres) {
        f[[cn]] <- vapply(seq_len(nrow(f)), function(i) {
          v <- input[[paste0(cn, "_", f$uid[i])]]
          if (is.null(v)) f[[cn]][i]
          else if (identical(campos[[cn]]$tipo, "check")) { if (isTRUE(v)) "1" else "" }
          else as.character(v)
        }, character(1))
      }
      f
    }

    observeEvent(input$agregar, {
      f <- valores()
      nueva <- as.data.frame(c(setNames(as.list(rep("", length(nombres))), nombres), list(uid = nuevo_uid())),
                             stringsAsFactors = FALSE)
      filas(rbind(f, nueva))
    })
    observeEvent(input$quitar, {
      f <- valores()
      filas(f[f$uid != input$quitar, , drop = FALSE])
    })

    output$filas <- renderUI({
      f <- filas()
      if (nrow(f) == 0) return(div(class = "text-muted small mb-2", "Sin registros."))
      cols <- paste(c(vapply(campos, function(c) c$ancho, ""), "44px"), collapse = " ")
      estilo <- sprintf("grid-template-columns:%s;", cols)
      cab <- div(class = "cab-lista", style = estilo,
                 lapply(campos, function(c) span(c$label)), span(""))
      cuerpo <- lapply(seq_len(nrow(f)), function(i) {
        uid <- f$uid[i]
        celdas <- lapply(nombres, function(cn) {
          cp <- campos[[cn]]
          idc <- ns(paste0(cn, "_", uid)); val <- f[[cn]][i]
          switch(cp$tipo %||% "text",
            text   = textInput(idc, NULL, value = val, placeholder = cp$label, width = "100%"),
            select = selectInput(idc, NULL, choices = c("Seleccione…" = "", cp$choices),
                                 selected = val, width = "100%", selectize = FALSE),
            area   = textAreaInput(idc, NULL, value = val, placeholder = cp$label, rows = 2,
                                   width = "100%", resize = "vertical"),
            check  = div(class = "pt-2", checkboxInput(idc, cp$texto %||% "Sí", value = identical(val, "1"), width = "100%")))
        })
        div(class = "fila-lista", style = estilo, celdas,
            tags$button(type = "button", class = "btn btn-outline-danger btn-sm", title = "Quitar",
                        onclick = sprintf("Shiny.setInputValue('%s','%s',{priority:'event'})", ns("quitar"), uid),
                        HTML("&times;<span class='d-md-none ms-1'>Quitar</span>")))
      })
      tagList(cab, cuerpo)
    })
    outputOptions(output, "filas", suspendWhenHidden = FALSE)

    list(
      get = function() {
        f <- valores()
        f <- f[, nombres, drop = FALSE]
        if (nrow(f) == 0) return(f)
        f[] <- lapply(f, trimws)
        keep <- Reduce(`|`, lapply(f, nzchar))
        f <- f[keep, , drop = FALSE]
        rownames(f) <- NULL
        f
      },
      set = function(d) {
        if (is.null(d) || nrow(d) == 0) { filas(vacia()); return(invisible()) }
        f <- d[, nombres, drop = FALSE]
        f[] <- lapply(f, function(x) ifelse(is.na(x), "", as.character(x)))
        f$uid <- vapply(seq_len(nrow(f)), function(i) nuevo_uid(), "")
        rownames(f) <- NULL
        filas(f)
      }
    )
  })
}

CAMPOS_PERSONA <- list(
  apellidos = list(label = "Apellidos", ancho = "1fr"),
  nombre    = list(label = "Nombres",   ancho = "1fr"))
CAMPOS_RET <- list(
  rut       = list(label = "RUT sin puntos ni guion (ej. 123456785)", ancho = "250px"),
  apellidos = list(label = "Apellidos", ancho = "1fr"),
  nombre    = list(label = "Nombres",   ancho = "1fr"))
CAMPOS_NUEVOS <- c(CAMPOS_RET, list(
  rut_pendiente = list(label = "¿RUT pendiente?", ancho = "130px", tipo = "check", texto = "RUT pendiente")))
CAMPOS_NEEP <- list(
  rut       = list(label = "RUT sin puntos ni guion (ej. 123456785)", ancho = "250px"),
  apellidos = list(label = "Apellidos", ancho = "1fr"),
  nombre    = list(label = "Nombres",   ancho = "1fr"),
  situacion = list(label = "Situación", ancho = "1.4fr", tipo = "select", choices = SITUACIONES_NEEP))
CAMPOS_EXC <- list(
  apellidos = list(label = "Apellidos", ancho = "1fr"),
  nombre    = list(label = "Nombres",   ancho = "1fr"),
  situacion = list(label = "Situación ocurrida", ancho = "1.6fr", tipo = "area"))

# Formulario vacío (para crear de inmediato un borrador, p. ej. una Toma 2 recién agregada)
f_vacio <- function() {
  vdf <- function(campos) as.data.frame(setNames(rep(list(character(0)), length(campos)), names(campos)),
                                        stringsAsFactors = FALSE)
  list(fecha = "", presentes = NA_integer_, aus_n = NA_integer_, ret_n = NA_integer_,
       obs = "", detalle_obs = "",
       ausentes = vdf(CAMPOS_PERSONA), retirados = vdf(CAMPOS_RET), nuevos = vdf(CAMPOS_NUEVOS),
       neep = vdf(CAMPOS_NEEP), excepcionales = vdf(CAMPOS_EXC))
}

# ------------------------------------------------------------------------------
# 9. INTERFAZ
# ------------------------------------------------------------------------------
CSS <- "
.fila-lista{display:grid;gap:8px;align-items:start;margin-bottom:6px}
.fila-lista .form-group,.fila-lista .shiny-input-container{margin-bottom:0!important;width:100%!important}
.cab-lista{display:grid;gap:8px;font-size:.8rem;color:#6c757d;margin-bottom:4px}
.barra-acciones{position:sticky;bottom:0;background:var(--bs-body-bg);border-top:1px solid var(--bs-border-color);padding:.75rem .5rem;z-index:5}
.pill{display:inline-block;padding:.2rem .6rem;border-radius:1rem;font-size:.8rem;margin:0 .25rem .25rem 0}
.pill-ok{background:#d1e7dd;color:#0f5132}.pill-pend{background:#fff3cd;color:#664d03}
.card-header{font-weight:600}
#tabla_instancias{overflow-x:auto}
input[id*='-apellidos_'],input[id*='-nombre_']{text-transform:uppercase}
.login-fondo{position:fixed;inset:0;overflow:auto;z-index:10;display:flex;align-items:center;justify-content:center;padding:1.5rem 1rem;background:linear-gradient(160deg,#f4f7fa 0%,#e6edf4 100%)}
.login-caja{width:100%;max-width:400px;text-align:center}
.login-logo{width:210px;max-width:70%;height:auto;margin-bottom:1.5rem;mix-blend-mode:multiply}
.login-card{background:#fff;border:1px solid #e1e7ee;border-radius:14px;padding:1.75rem 1.5rem 1.5rem;box-shadow:0 6px 24px rgba(20,40,70,.08);text-align:left}
.login-titulo{font-size:1.15rem;font-weight:600;color:#1f2d3d;margin:0 0 .25rem;line-height:1.3}
.login-proceso{display:inline-block;font-size:.72rem;font-weight:600;letter-spacing:.06em;text-transform:uppercase;color:#1b6ca8;background:#e3effa;border-radius:1rem;padding:.1rem .6rem;margin-bottom:.4rem}
.login-sub{font-size:.88rem;color:#6c757d;margin:.25rem 0 1.1rem}
.login-card .form-control{height:46px;font-size:1.05rem;letter-spacing:.15em}
.login-btn{height:46px;font-weight:600;margin-top:.25rem}
.login-pie{font-size:.78rem;color:#8a97a6;margin-top:1.1rem}
@media (max-width:768px){.fila-lista{grid-template-columns:1fr!important}.cab-lista{display:none}}
"

JS <- "
(function reg(){
  if (!(window.Shiny && Shiny.inputBindings && window.jQuery)) { setTimeout(reg, 20); return; }
  var b = new Shiny.InputBinding();
  $.extend(b, {
    find: function(scope){ return $(scope).find('input.fecha-nativa'); },
    getValue: function(el){ return el.value || ''; },
    subscribe: function(el, cb){ $(el).on('change.fechaNativa input.fechaNativa', function(){ cb(); }); },
    unsubscribe: function(el){ $(el).off('.fechaNativa'); },
    receiveMessage: function(el, data){
      if (data.hasOwnProperty('value')) { el.value = data.value || ''; $(el).trigger('change'); }
    }
  });
  Shiny.inputBindings.register(b, 'evaluaciones.fechaNativa');
})();
$(document).on('keyup', '#clave', function(e){ if (e.key === 'Enter') $('#entrar').click(); });
"

fecha_nativa <- function(id, label, value = "") {
  div(class = "form-group shiny-input-container",
      tags$label(`for` = id, label),
      tags$input(id = id, type = "date", class = "form-control fecha-nativa", value = value))
}
entero_input <- function(id, label, value = "", ...) {
  x <- textInput(id, label, value = value, ...)
  htmltools::tagQuery(x)$find("input")$addAttrs(inputmode = "numeric", autocomplete = "off")$allTags()
}

ui <- page_fluid(
  theme = bs_theme(version = 5, primary = "#1b6ca8"),
  title = paste0("Registro de aplicaciones", SUFIJO_PROCESO),
  tags$head(tags$style(HTML(CSS)), tags$script(HTML(JS))),
  div(class = "container", style = "max-width:1100px;", uiOutput("pagina"))
)

# Logo embebido (no depende de carpetas www ni de rutas al desplegar)
LOGO_URI <- local({
  f <- "logo.png"
  if (!file.exists(f)) return("")
  paste0("data:image/png;base64,", jsonlite::base64_enc(readBin(f, "raw", file.info(f)$size)))
})

login_ui <- div(class = "login-fondo",
  div(class = "login-caja",
    if (nzchar(LOGO_URI)) img(src = LOGO_URI, alt = "Fundación Astoreca", class = "login-logo"),
    div(class = "login-card",
      h1(class = "login-titulo", "Registro de aplicaciones de evaluaciones"),
      if (nzchar(PROCESO)) div(class = "login-proceso", PROCESO),
      p(class = "login-sub", "Ingresa con la clave que te entregó la Fundación."),
      passwordInput("clave", "Clave de identificación", width = "100%"),
      actionButton("entrar", "Ingresar", class = "btn-primary w-100 login-btn"),
      uiOutput("msg_login")
    ),
    div(class = "login-pie", "Fundación Astoreca · Educar bien")
  )
)

vista_lista_ui <- tagList(
  selectizeInput("colegio", "Colegio", choices = c("Seleccione un colegio…" = "", OPCIONES_COLEGIO), width = "100%"),
  uiOutput("resumen_colegio"),
  fluidRow(
    column(4, selectInput("flt_eval", "Evaluación", c("Todas", EVALUACIONES))),
    column(4, selectInput("flt_estado", "Estado", c("Todos", "Pendiente", "Borrador", "Enviada", "Habilitada para corregir", "Requiere corrección"))),
    column(4, div(style = "margin-top:32px;", actionButton("refrescar", "↻ Actualizar lista", class = "btn-outline-secondary")))
  ),
  uiOutput("agregar_toma2"),
  DTOutput("tabla_instancias"),
  helpText("Toca una fila (o haz clic) para abrir el registro. Los registros enviados quedan bloqueados: si necesitas corregir uno, pide al administrador que lo habilite. ",
           "Si hubo una segunda toma de algún curso, agrégala con el recuadro de arriba: aparecerá como una fila nueva con Toma 2.")
)

seccion <- function(titulo, ...) card(class = "mb-3", card_header(titulo), card_body(...))

crear_form <- function(k, reg) {
  v <- function(col) if (is.null(reg) || is.na(reg[[col]])) "" else as.character(reg[[col]])
  idk <- function(x) paste0(x, "_", k)
  tagList(
    seccion("1. Datos generales",
      fluidRow(
        column(6, fecha_nativa(idk("f_fecha"), "Fecha de aplicación", v("fecha_aplicacion"))),
        column(6, entero_input(idk("f_presentes"), "Cantidad total de estudiantes presentes", v("presentes"), width = "100%")))),
    seccion("2. Estudiantes ausentes",
      entero_input(idk("f_aus_n"), "Cantidad total de estudiantes ausentes (0 si no hubo)", v("ausentes_decl"), width = "260px"),
      lista_ui("l_aus")),
    seccion("3. Estudiantes retirados",
      entero_input(idk("f_ret_n"), "Cantidad total de estudiantes retirados (0 si no hubo)", v("retirados_decl"), width = "260px"),
      lista_ui("l_ret")),
    seccion("4. Estudiantes nuevos",
      helpText("Estudiantes que se incorporaron al curso. Déjalo vacío si no hubo. ",
               "Si aún no tienes el RUT de un estudiante, marca ", strong("RUT pendiente"),
               ": no será obligatorio por ahora, pero debes conseguirlo con el colegio."),
      lista_ui("l_nue")),
    seccion("5. Situaciones de estudiantes NEEP",
      helpText("Registra a los estudiantes NEEP que tuvieron alguna de las situaciones indicadas en la aplicación. Déjalo vacío si no hubo."),
      lista_ui("l_nee")),
    seccion("6. Observaciones generales y situaciones excepcionales",
      selectInput(idk("f_obs"), "Observaciones", c("Seleccione…" = "", OPCIONES_OBS),
                  selected = v("observaciones"), width = "260px", selectize = FALSE),
      conditionalPanel(sprintf("input.%s == 'Con observaciones'", idk("f_obs")),
        textAreaInput(idk("f_detobs"), "Detalle de las observaciones", value = v("detalle_obs"),
                      width = "100%", rows = 4, resize = "vertical")),
      tags$hr(),
      tags$strong("Situaciones excepcionales con estudiantes"),
      helpText("Déjalo vacío si no hubo."),
      lista_ui("l_exc"))
  )
}

vista_form_ui <- tagList(
  div(class = "d-flex flex-wrap justify-content-between align-items-center mb-2",
      actionButton("volver", "← Volver", class = "btn-outline-secondary btn-sm"),
      uiOutput("indicador_guardado", inline = TRUE)),
  uiOutput("titulo_form"),
  uiOutput("aviso_form"),
  uiOutput("checklist"),
  uiOutput("form_ui"),
  div(class = "barra-acciones", uiOutput("barra_acciones"))
)

eval_ui <- function() tagList(
  div(class = "d-flex justify-content-between align-items-center py-3 border-bottom mb-3",
      div(strong(paste0("Registro de aplicaciones", SUFIJO_PROCESO)), span(class = "text-muted ms-2", textOutput("quien", inline = TRUE))),
      actionButton("salir", "Cerrar sesión", class = "btn-outline-secondary btn-sm")),
  conditionalPanel("!output.en_form", vista_lista_ui),
  conditionalPanel("output.en_form", vista_form_ui)
)

admin_ui <- function() tagList(
  div(class = "d-flex justify-content-between align-items-center py-3 border-bottom mb-3",
      div(strong("Administración"), span(class = "text-muted ms-2", paste0("Registro de aplicaciones", SUFIJO_PROCESO))),
      actionButton("salir", "Cerrar sesión", class = "btn-outline-secondary btn-sm")),
  conditionalPanel("output.en_form", vista_form_ui),
  conditionalPanel("!output.en_form", navset_tab(
    nav_panel("Avance",
      div(class = "my-3", actionButton("adm_refrescar", "↻ Actualizar", class = "btn-outline-secondary btn-sm")),
      uiOutput("adm_totales"),
      h5("Por colegio"), DTOutput("adm_tabla_colegio"),
      h5(class = "mt-4", "Por gestora"), DTOutput("adm_tabla_gestora")),
    nav_panel("Revisión",
      p(class = "mt-3 text-muted",
        strong("Doble clic en una fila"), " (o selecciónala y presiona ", em("Ver contenido"),
        ") para ver lo que registró el evaluador. Para cambiar el estado de revisión de varios registros a la vez, selecciónalos y usa los botones de abajo."),
      DTOutput("adm_tabla_rev"),
      div(class = "mt-2 d-flex gap-2",
          actionButton("adm_ver", "Ver contenido del registro seleccionado", class = "btn-primary btn-sm"),
          actionButton("adm_editar_sel", "Editar registro seleccionado", class = "btn-outline-primary btn-sm"),
          actionButton("adm_eliminar_sel", "Eliminar registro(s) seleccionado(s)", class = "btn-outline-danger btn-sm ms-auto")),
      div(class = "mt-3 d-flex flex-wrap gap-2 align-items-end",
          textInput("adm_comentario", "Comentario para el evaluador (opcional)", width = "420px"),
          actionButton("adm_rev_ok", "Marcar Revisado", class = "btn-success"),
          actionButton("adm_rev_corr", "Requiere corrección", class = "btn-warning"),
          actionButton("adm_rev_pend", "Volver a Pendiente", class = "btn-outline-secondary")),
      div(class = "mt-2 d-flex flex-wrap gap-2 align-items-center",
          actionButton("adm_hab", "Habilitar corrección (evaluador que envió)", class = "btn-info btn-sm"),
          actionButton("adm_deshab", "Bloquear de nuevo", class = "btn-outline-secondary btn-sm"),
          span(class = "text-muted small", "Solo el evaluador que envió el registro podrá abrirlo; al actualizarlo vuelve a quedar bloqueado. ",
               "\"Requiere corrección\" también lo habilita; \"Marcar Revisado\" lo bloquea."))),
    nav_panel("Claves",
      div(class = "my-3",
        p("Genera claves nuevas. Las claves se muestran y descargan ", strong("una sola vez"),
          ": en la base solo queda una versión cifrada. Son números de 6 dígitos, distintos entre sí."),
        fluidRow(
          column(3, numericInput("adm_n_claves", "Cantidad de claves", 10, min = 1, max = 900)),
          column(6, textAreaInput("adm_nombres", "o pega nombres de evaluadores (uno por línea)", rows = 3, width = "100%")),
          column(3, div(style = "margin-top:32px;", actionButton("adm_generar", "Generar claves", class = "btn-primary")))),
        uiOutput("adm_nuevas")),
      h5("Claves existentes"),
      DTOutput("adm_tabla_claves"),
      div(class = "mt-2 d-flex gap-2",
          actionButton("adm_clave_off", "Desactivar seleccionadas", class = "btn-outline-danger btn-sm"),
          actionButton("adm_clave_on", "Activar seleccionadas", class = "btn-outline-success btn-sm"))),
    nav_panel("Descarga",
      div(class = "my-3",
        p("Excel con las hojas Aplicaciones, Ausentes, Retirados, Nuevos, NEEP, Excepcionales, Seguimiento, Resumen e Historial."),
        checkboxInput("adm_borradores", "Incluir también los borradores (no enviados)", FALSE),
        downloadButton("adm_descargar", "Descargar Excel", class = "btn-primary")))
  ))
)

# ------------------------------------------------------------------------------
# 10. SERVIDOR
# ------------------------------------------------------------------------------
server <- function(input, output, session) {
  session$allowReconnect(TRUE)

  ses <- reactiveValues(rol = NULL, clave_id = NULL, nombre = NULL, fallos = 0L, espera_hasta = 0)
  msg_login <- reactiveVal("")

  # Estado del registro abierto
  rv <- reactiveValues(
    slot = NULL, k = 0L, reg = NULL, rev = 0L, estado = "pendiente",
    hash = NULL, ult_guardado = NULL, error_guardado = NULL, conflicto = FALSE,
    pendiente_abrir = NULL, tick = 0L)

  # --------------------------- Acceso -----------------------------------------
  output$pagina <- renderUI({
    if (is.null(ses$rol)) login_ui else if (ses$rol == "admin") admin_ui() else eval_ui()
  })
  output$msg_login <- renderUI(if (nzchar(msg_login())) div(class = "text-danger small mt-2", msg_login()))

  observeEvent(input$entrar, {
    x <- trimws(input$clave %||% "")
    if (!nzchar(x)) { msg_login("Ingresa tu clave."); return() }
    if (nzchar(ADMIN_CLAVE) && identical(x, ADMIN_CLAVE)) {
      ses$rol <- "admin"; ses$clave_id <- "ADMIN"; ses$nombre <- "Administrador"
      msg_login(""); return()
    }
    # Freno por sesión (no global): tras 5 fallos seguidos, espera creciente solo para esa sesión
    if (ses$fallos >= 5 && as.numeric(Sys.time()) < ses$espera_hasta) {
      msg_login(sprintf("Demasiados intentos fallidos. Espera %d segundos.",
                        ceiling(ses$espera_hasta - as.numeric(Sys.time())))); return()
    }
    fila <- q("SELECT clave_id, evaluador FROM claves WHERE hash = ?h AND activa = 1", h = hash_clave(x))
    if (nrow(fila) == 1) {
      ses$rol <- "evaluador"; ses$clave_id <- fila$clave_id
      ses$nombre <- if (vacio(fila$evaluador)) fila$clave_id else fila$evaluador
      ex("UPDATE claves SET ultimo_uso = ?t WHERE clave_id = ?cl", t = ahora(), cl = fila$clave_id)
      msg_login("")
    } else {
      ses$fallos <- ses$fallos + 1L
      if (ses$fallos >= 5) ses$espera_hasta <- as.numeric(Sys.time()) + min(300, 15 * 2^(ses$fallos - 5))
      msg_login("Clave incorrecta o desactivada.")
    }
  })

  liberar_actual <- function() {
    s <- isolate(rv$slot)
    if (!is.null(s)) try(lock_liberar(s$id_registro, isolate(ses$clave_id)), silent = TRUE)
  }
  observeEvent(input$salir, {
    liberar_actual()
    session$reload()
  })
  session$onSessionEnded(function() liberar_actual())

  output$quien <- renderText(paste0("· ", ses$nombre))

  # --------------------------- Lista de instancias ----------------------------
  nombres_claves <- reactive({
    rv$tick; tick_adm()
    rbind(q("SELECT clave_id, evaluador FROM claves"), data.frame(clave_id = "ADMIN", evaluador = "Administración"))
  })

  lista_df <- reactive({
    req(input$colegio)
    rv$tick; input$refrescar
    se <- slots_ext()
    s <- se[se$sigla == input$colegio, , drop = FALSE]
    est <- q("SELECT id_registro, estado, estado_revision, habilitado_para, actualizado, actualizado_por FROM registros WHERE sigla = ?sg",
             sg = input$colegio)
    m <- match(s$id_registro, est$id_registro)
    estado <- est$estado[m]; estado[is.na(estado)] <- "pendiente"
    rev_ <- est$estado_revision[m]
    etiqueta <- c(pendiente = "Pendiente", borrador = "Borrador", enviada = "Enviada")[estado]
    # Solo quien fue habilitado ve que puede corregir; para el resto la enviada está bloqueada
    hab <- est$habilitado_para[m]
    mio <- estado == "enviada" & !is.na(hab) & hab == ses$clave_id
    etiqueta[mio] <- "Habilitada para corregir"
    etiqueta[mio & !is.na(rev_) & rev_ == "Requiere corrección"] <- "Requiere corrección"
    nc <- nombres_claves()
    por <- nc$evaluador[match(est$actualizado_por[m], nc$clave_id)]
    por[is.na(por) | !nzchar(por)] <- est$actualizado_por[m][is.na(por) | !nzchar(por)]
    data.frame(id = s$id_registro, Curso = ifelse(s$grado == "K", "Kínder", paste0(s$grado, "°")),
               Letra = s$letra, `Evaluación` = s$evaluacion,
               Toma = ifelse(is.na(s$toma), "—", as.character(s$toma)), Modalidad = s$modalidad,
               Estado = unname(etiqueta), `Última actualización` = fmt_fecha_hora(est$actualizado[m]),
               Por = ifelse(is.na(por), "", por), check.names = FALSE, stringsAsFactors = FALSE)
  })
  lista_filtrada <- reactive({
    d <- lista_df()
    if (input$flt_eval != "Todas") d <- d[d$`Evaluación` == input$flt_eval, , drop = FALSE]
    if (input$flt_estado != "Todos") d <- d[d$Estado == input$flt_estado, , drop = FALSE]
    d
  })

  output$resumen_colegio <- renderUI({
    req(input$colegio)
    d <- lista_df()
    n <- function(e) sum(d$Estado %in% e)
    div(class = "mb-3",
        span(class = "pill pill-ok", sprintf("Enviadas: %d", n(c("Enviada", "Habilitada para corregir", "Requiere corrección")))),
        span(class = "pill pill-pend", sprintf("Borradores: %d", n("Borrador"))),
        span(class = "pill", style = "background:#e9ecef;", sprintf("Pendientes: %d", n("Pendiente"))),
        if (n("Requiere corrección")) span(class = "pill", style = "background:#f8d7da;color:#842029;",
                                            sprintf("Requieren corrección: %d", n("Requiere corrección"))))
  })

  output$tabla_instancias <- renderDT({
    req(input$colegio)
    d <- lista_filtrada()
    datatable(d[, setdiff(names(d), "id"), drop = FALSE], rownames = FALSE, selection = "single",
              options = list(pageLength = 25, dom = "ftip",
                             # En pantallas chicas se ocultan Modalidad, Última actualización y Por
                             columnDefs = list(list(targets = c(4, 6, 7), className = "d-none d-md-table-cell")),
                             language = list(
                search = "Buscar:", info = "_START_–_END_ de _TOTAL_", infoEmpty = "Sin registros",
                zeroRecords = "Sin resultados", paginate = list(previous = "Anterior", `next` = "Siguiente")))) |>
      formatStyle("Estado",
        backgroundColor = styleEqual(c("Pendiente", "Borrador", "Enviada", "Habilitada para corregir", "Requiere corrección"),
                                     c("#ffffff", "#fff3cd", "#d1e7dd", "#cff4fc", "#f8d7da")))
  })
  proxy_tabla <- dataTableProxy("tabla_instancias")

  observeEvent(input$tabla_instancias_rows_selected, {
    i <- input$tabla_instancias_rows_selected
    selectRows(proxy_tabla, NULL)
    d <- lista_filtrada()
    if (length(i) == 1 && i <= nrow(d)) intentar_abrir(d$id[i])
  })

  # Agregar una segunda toma de cualquier curso, aunque la Toma 1 ya esté enviada (y bloqueada)
  output$agregar_toma2 <- renderUI({
    req(input$colegio)
    rv$tick; input$refrescar
    c1 <- slots[slots$sigla == input$colegio, , drop = FALSE]
    ya <- q("SELECT id_registro FROM registros WHERE sigla = ?sg AND id_registro LIKE '%-T2'", sg = input$colegio)$id_registro
    c1 <- c1[!vapply(c1$id_registro, id_toma2, "") %in% ya, , drop = FALSE]
    if (nrow(c1) == 0) return(NULL)
    etq <- paste0(etiqueta_curso(c1$grado, c1$letra), " · ", c1$evaluacion)
    div(class = "d-flex flex-wrap gap-2 align-items-end mb-3 p-2 border rounded",
        div(style = "min-width:220px;flex:1;",
            selectInput("sel_toma2", "¿Hubo una segunda toma de algún curso? Agrégala aquí",
                        choices = c("Seleccione el curso…" = "", setNames(c1$id_registro, etq)),
                        width = "100%", selectize = FALSE)),
        div(style = "margin-bottom:1rem;", actionButton("btn_toma2", "+ Agregar Toma 2", class = "btn-outline-primary")))
  })
  # Crea de inmediato el borrador de una Toma 2 (así aparece en la lista como una fila nueva)
  crear_toma2 <- function(id2) {
    if (nrow(q("SELECT 1 AS x FROM registros WHERE id_registro = ?id", id = id2)) > 0) return(TRUE)
    sl <- slot_por_id(id2)
    if (nrow(sl) == 0) return(FALSE)
    r <- tryCatch(guardar_registro(sl, f_vacio(), "borrador", ses$clave_id, 0L),
                  error = function(e) list(ok = FALSE))
    isTRUE(r$ok)
  }
  observeEvent(input$btn_toma2, {
    id1 <- input$sel_toma2
    if (!nzchar(id1 %||% "")) { showNotification("Selecciona el curso.", type = "warning"); return() }
    id2 <- id_toma2(id1)
    if (!crear_toma2(id2)) { showNotification("No se pudo crear la Toma 2. Inténtalo de nuevo.", type = "error"); return() }
    rv$tick <- rv$tick + 1L
    intentar_abrir(id2)
  })

  output$en_form <- reactive(!is.null(rv$slot))
  outputOptions(output, "en_form", suspendWhenHidden = FALSE)

  # --------------------------- Listas dinámicas -------------------------------
  L <- list(
    aus = lista_server("l_aus", CAMPOS_PERSONA),
    ret = lista_server("l_ret", CAMPOS_RET),
    nue = lista_server("l_nue", CAMPOS_NUEVOS),
    nee = lista_server("l_nee", CAMPOS_NEEP),
    exc = lista_server("l_exc", CAMPOS_EXC))

  # --------------------------- Abrir un registro ------------------------------
  # Un evaluador no puede abrir un registro ya enviado, salvo que el administrador lo haya habilitado para él
  bloqueado_para_mi <- function(id) {
    if (identical(ses$rol, "admin")) return(FALSE)
    r <- q("SELECT estado, habilitado_para FROM registros WHERE id_registro = ?id", id = id)
    nrow(r) == 1 && identical(r$estado, "enviada") && !identical(r$habilitado_para, ses$clave_id)
  }
  aviso_bloqueado <- function() showModal(modalDialog(
    title = "Registro enviado y bloqueado",
    "Este registro ya fue enviado y no se puede abrir. Si necesitas corregirlo, pide al administrador que lo habilite.",
    easyClose = TRUE, footer = modalButton("Entendido")))

  intentar_abrir <- function(id) {
    if (bloqueado_para_mi(id)) { aviso_bloqueado(); return(invisible(FALSE)) }
    otro <- lock_otro(id, ses$clave_id)
    if (!is.null(otro)) {
      rv$pendiente_abrir <- id
      showModal(modalDialog(
        title = "Registro en uso",
        sprintf("%s tiene abierto este registro (actividad hace %d min). Si lo abres también, los cambios de uno pueden pisar los del otro.",
                otro$quien, otro$mins),
        footer = tagList(modalButton("Cancelar"), actionButton("abrir_igual", "Abrir de todos modos", class = "btn-warning"))))
    } else abrir(id)
  }
  observeEvent(input$abrir_igual, { removeModal(); abrir(rv$pendiente_abrir) })

  abrir <- function(id) {
    if (bloqueado_para_mi(id)) { aviso_bloqueado(); return(invisible(FALSE)) }
    slot <- slot_por_id(id)
    d <- cargar_registro(id)
    reg <- d$reg
    lock_set(id, ses$clave_id)
    L$aus$set(d$ausentes); L$ret$set(d$retirados); L$nue$set(d$nuevos)
    L$nee$set(d$neep);     L$exc$set(d$excepcionales)
    rv$hash <- NULL; rv$ult_guardado <- NULL; rv$error_guardado <- NULL; rv$conflicto <- FALSE
    rv$reg <- reg
    rv$rev <- if (is.null(reg)) 0L else as.integer(reg$rev)
    rv$estado <- if (is.null(reg)) "pendiente" else reg$estado
    rv$k <- rv$k + 1L
    rv$slot <- slot
  }

  output$form_ui <- renderUI({
    req(rv$k > 0)
    crear_form(rv$k, isolate(rv$reg))
  })
  outputOptions(output, "form_ui", suspendWhenHidden = FALSE)

  output$titulo_form <- renderUI({
    s <- rv$slot; req(s)
    div(class = "mb-3",
        h4(class = "mb-0", sprintf("%s · %s · %s%s", s$sigla, etiqueta_curso(s$grado, s$letra), s$evaluacion,
                                   txt_toma(s$grado, s$toma))),
        span(class = "text-muted", sprintf("%s · Modalidad: %s", s$nombre, s$modalidad)))
  })
  output$aviso_form <- renderUI({
    r <- rv$reg; s <- rv$slot; req(s)
    tagList(
      if (identical(ses$rol, "admin"))
        div(class = "alert alert-warning py-2", strong("Editando como administrador. "),
            "Los cambios quedan registrados a nombre de Administración y no modifican el estado de revisión."),
      if (!is.null(r) && identical(r$estado_revision, "Requiere corrección") && identical(r$estado, "enviada"))
        div(class = "alert alert-danger", strong("Requiere corrección. "),
            if (!vacio(r$comentario_revision)) r$comentario_revision else "Revisa los datos y vuelve a enviarlos."),
      if (identical(rv$estado, "enviada"))
        div(class = "alert alert-success py-2",
            "Este registro ya fue enviado. Puedes corregirlo y presionar ", strong("Actualizar envío"), ". Al actualizarlo, volverá a quedar bloqueado."),
      if (!is.na(s$toma) && s$toma == 1L) {
        existe2 <- nrow(q("SELECT 1 AS x FROM registros WHERE id_registro = ?id", id = id_toma2(s$id_registro))) > 0
        div(class = "alert alert-info py-2",
            if (s$grado == "K") "Kínder puede tener dos tomas, con estudiantes distintos. Esta es la Toma 1. "
            else "Esta es la Toma 1. Si hubo una segunda toma, con estudiantes distintos, puedes agregarla. ",
            actionLink("ir_toma2", if (existe2) "Ir a la Toma 2" else "+ Agregar Toma 2"))
      },
      if (!is.na(s$toma) && s$toma == 2L)
        div(class = "alert alert-info py-2",
            "Toma 2: estudiantes distintos de la Toma 1. Si no hubo segunda toma, usa ",
            strong("Descartar Toma 2"), ". ",
            actionLink("ir_toma1", "Ir a la Toma 1"),
            if (!identical(rv$estado, "enviada")) tagList(" · ", actionLink("descartar_toma", "Descartar Toma 2")))
    )
  })

  # --------------------------- Lectura del formulario -------------------------
  leer_form <- reactive({
    k <- rv$k
    if (k == 0) return(NULL)
    ids <- paste0(c("f_fecha_", "f_presentes_", "f_aus_n_", "f_ret_n_", "f_obs_"), k)
    vals <- lapply(ids, function(i) input[[i]])
    if (any(vapply(vals, is.null, TRUE))) return(NULL)          # aún no sincronizado
    list(
      fecha = vals[[1]], presentes = a_entero(vals[[2]]),
      aus_n = a_entero(vals[[3]]), ret_n = a_entero(vals[[4]]),
      obs = vals[[5]], detalle_obs = trimws(input[[paste0("f_detobs_", k)]] %||% ""),
      ausentes = L$aus$get(), retirados = L$ret$get(), nuevos = L$nue$get(),
      neep = L$nee$get(), excepcionales = L$exc$get())
  })
  validacion <- reactive({ f <- leer_form(); req(f); validar_form(f) })

  # Línea base: contenido recién cargado (para no guardar sin cambios reales)
  observe({
    f <- leer_form()
    if (!is.null(f) && !is.null(isolate(rv$slot)) && is.null(isolate(rv$hash))) rv$hash <- digest::digest(f)
  })
  hay_cambios <- reactive({
    f <- leer_form()
    !is.null(f) && !is.null(rv$hash) && !identical(digest::digest(f), rv$hash)
  })

  # --------------------------- Guardado ---------------------------------------
  ejecutar_guardado <- function(accion, f, forzar = FALSE) {
    r <- tryCatch(
      guardar_registro(rv$slot, f, accion, ses$clave_id, rv$rev, forzar),
      error = function(e) list(ok = FALSE, conflicto = FALSE, error = conditionMessage(e)))
    if (isTRUE(r$ok)) {
      rv$rev <- r$rev; rv$estado <- r$estado; rv$hash <- digest::digest(f)
      rv$ult_guardado <- r$ts; rv$error_guardado <- NULL; rv$conflicto <- FALSE
      rv$reg <- cargar_registro(rv$slot$id_registro)$reg
      lock_set(rv$slot$id_registro, ses$clave_id)
    } else if (isTRUE(r$conflicto)) {
      rv$conflicto <- TRUE
      nc <- nombres_claves()
      quien <- nc$evaluador[match(r$por, nc$clave_id)]
      if (is.na(quien) || !nzchar(quien)) quien <- r$por
      showModal(modalDialog(
        title = "Este registro cambió mientras lo editabas",
        sprintf("%s guardó cambios el %s. Puedes recargar lo guardado (pierdes lo que no alcanzó a guardarse) o sobrescribir con lo que tienes en pantalla.",
                quien, fmt_fecha_hora(r$cuando)),
        footer = tagList(actionButton("conflicto_recargar", "Recargar lo guardado", class = "btn-outline-primary"),
                         actionButton("conflicto_sobrescribir", "Sobrescribir con mi versión", class = "btn-danger")),
        easyClose = FALSE))
    } else {
      rv$error_guardado <- r$error %||% "Error desconocido"
    }
    r
  }

  observeEvent(input$conflicto_recargar, { removeModal(); abrir(rv$slot$id_registro) })
  observeEvent(input$conflicto_sobrescribir, {
    removeModal()
    f <- leer_form(); req(f)
    accion <- if (identical(rv$estado, "enviada")) "enviar" else "borrador"
    ejecutar_guardado(accion, f, forzar = TRUE)
  })

  leer_debounced <- debounce(leer_form, 1500)
  observeEvent(leer_debounced(), {
    f <- leer_debounced()
    if (is.null(f) || is.null(rv$slot) || is.null(rv$hash) || isTRUE(rv$conflicto)) return()
    if (!rv$estado %in% c("pendiente", "borrador")) return()      # los envíos no se autoguardan
    if (identical(digest::digest(f), rv$hash)) return()
    ejecutar_guardado("borrador", f)
  }, ignoreInit = TRUE)

  # Reintento si el último guardado falló (p. ej. conexión inestable)
  observe({
    invalidateLater(15000)
    isolate({
      if (!is.null(rv$slot) && !is.null(rv$error_guardado) && rv$estado %in% c("pendiente", "borrador")) {
        f <- leer_form()
        if (!is.null(f) && !identical(digest::digest(f), rv$hash)) ejecutar_guardado("borrador", f)
      }
    })
  })
  # Mantiene vigente el aviso de "registro en uso"
  observe({
    invalidateLater(60000)
    isolate(if (!is.null(rv$slot)) try(lock_set(rv$slot$id_registro, ses$clave_id), silent = TRUE))
  })

  output$indicador_guardado <- renderUI({
    req(rv$slot)
    if (!is.null(rv$error_guardado))
      return(span(class = "text-danger small", "⚠ No se pudo guardar (¿sin conexión?). Reintentando…"))
    if (identical(rv$estado, "enviada")) {
      r <- rv$reg
      txt <- if (isTRUE(hay_cambios())) "Cambios sin actualizar"
             else if (!is.null(r)) sprintf("Enviado · última actualización %s", fmt_fecha_hora(r$actualizado)) else ""
      return(span(class = if (isTRUE(hay_cambios())) "text-warning small" else "text-success small", txt))
    }
    if (isTRUE(hay_cambios())) span(class = "text-muted small", "Guardando cambios…")
    else if (!is.null(rv$ult_guardado)) span(class = "text-success small", paste0("✓ Borrador guardado ", substr(rv$ult_guardado, 12, 19)))
    else span(class = "text-muted small", "Se guarda automáticamente mientras escribes")
  })

  output$checklist <- renderUI({
    v <- validacion()
    nombres <- c(general = "Datos generales", ausentes = "Ausentes", retirados = "Retirados",
                 nuevos = "Nuevos", neep = "Situaciones NEEP", obs = "Observaciones")
    div(class = "mb-3", lapply(names(nombres), function(n) {
      s <- v[[n]]
      span(class = paste("pill", if (s$ok) "pill-ok" else "pill-pend"),
           title = paste(s$msg, collapse = "\n"),
           paste0(if (s$ok) "✓ " else "• ", nombres[[n]], if (s$ok && n %in% c("ausentes", "retirados", "nuevos", "neep")) paste0(" (", s$resumen, ")") else ""))
    }))
  })

  output$barra_acciones <- renderUI({
    v <- validacion()
    msgs <- unlist(lapply(v, `[[`, "msg"))
    ok <- length(msgs) == 0
    enviada <- identical(rv$estado, "enviada")
    puede <- ok && (!enviada || isTRUE(hay_cambios()))
    etiqueta <- if (enviada) "Actualizar envío" else "Enviar"
    tagList(
      if (!ok) div(class = "small text-danger mb-2", style = "max-height:28vh;overflow-y:auto;", strong("Para enviar falta:"),
                   tags$ul(class = "mb-0", lapply(head(msgs, 8), tags$li))),
      div(class = "d-flex gap-2 flex-wrap",
          if (!enviada) actionButton("guardar_borrador", "Guardar borrador", class = "btn-outline-primary"),
          if (puede) actionButton("enviar", etiqueta, class = "btn-success")
          else actionButton("enviar_off", etiqueta, class = "btn-success", disabled = "disabled")))
  })

  observeEvent(input$guardar_borrador, {
    f <- leer_form(); req(f)
    r <- ejecutar_guardado("borrador", f)
    if (isTRUE(r$ok)) showNotification("Borrador guardado.", type = "message", duration = 3)
  })

  observeEvent(input$enviar, {
    f <- leer_form(); req(f)
    if (!all(vapply(validar_form(f), function(s) s$ok, TRUE))) return()
    enviada <- identical(rv$estado, "enviada")
    showModal(modalDialog(
      title = if (enviada) "Actualizar envío" else "Enviar registro",
      if (enviada) "Se guardarán tus correcciones y quedará registro de la versión anterior."
      else "Una vez enviado podrás corregirlo después con \"Actualizar envío\".",
      footer = tagList(modalButton("Cancelar"),
                       actionButton("confirmar_envio", if (enviada) "Actualizar" else "Enviar", class = "btn-success"))))
  })
  observeEvent(input$confirmar_envio, {
    removeModal()
    f <- leer_form(); req(f)
    r <- ejecutar_guardado("enviar", f)
    if (isTRUE(r$ok)) {
      showNotification("Registro enviado ✓", type = "message", duration = 5)
      salir_del_registro()
    } else if (!isTRUE(r$conflicto)) {
      showNotification(paste("No se pudo enviar:", r$error %||% ""), type = "error", duration = 8)
    }
  })

  salir_del_registro <- function() {
    liberar_actual()
    rv$slot <- NULL; rv$reg <- NULL; rv$hash <- NULL; rv$estado <- "pendiente"
    rv$tick <- rv$tick + 1L
    if (identical(isolate(ses$rol), "admin")) tick_adm(isolate(tick_adm()) + 1L)
  }
  observeEvent(input$volver, {
    f <- leer_form()
    if (identical(rv$estado, "enviada") && isTRUE(hay_cambios())) {
      showModal(modalDialog(title = "Cambios sin actualizar",
        "Hiciste cambios en un registro ya enviado que aún no se guardan. ¿Salir sin guardarlos?",
        footer = tagList(modalButton("Seguir editando"), actionButton("salir_sin_guardar", "Salir sin guardar", class = "btn-danger"))))
      return()
    }
    if (rv$estado %in% c("pendiente", "borrador") && !is.null(f) && isTRUE(hay_cambios()) && !isTRUE(rv$conflicto))
      ejecutar_guardado("borrador", f)
    salir_del_registro()
  })
  observeEvent(input$salir_sin_guardar, { removeModal(); salir_del_registro() })
  observeEvent(input$refrescar, rv$tick <- rv$tick + 1L)

  # --- Toma 2 (se crea solo si el evaluador la agrega) ---
  ir_a_toma <- function(id) {
    f <- leer_form()
    if (identical(rv$estado, "enviada") && isTRUE(hay_cambios())) {
      showNotification("Primero actualiza el envío, o vuelve a la lista para descartar los cambios.", type = "warning")
      return()
    }
    if (rv$estado %in% c("pendiente", "borrador") && !is.null(f) && isTRUE(hay_cambios()) && !isTRUE(rv$conflicto))
      ejecutar_guardado("borrador", f)
    liberar_actual()
    abrir(id)
  }
  observeEvent(input$ir_toma2, {
    id2 <- id_toma2(rv$slot$id_registro)
    if (!crear_toma2(id2)) { showNotification("No se pudo crear la Toma 2. Inténtalo de nuevo.", type = "error"); return() }
    ir_a_toma(id2)
  })
  observeEvent(input$ir_toma1, ir_a_toma(id_toma1(rv$slot$id_registro)))
  observeEvent(input$descartar_toma, {
    showModal(modalDialog(title = "Descartar Toma 2",
      "Se borrará lo que escribiste en la Toma 2. La Toma 1 no se modifica.",
      footer = tagList(modalButton("Cancelar"), actionButton("confirmar_descartar", "Descartar", class = "btn-danger"))))
  })
  observeEvent(input$confirmar_descartar, {
    removeModal()
    descartar_registro(rv$slot)
    salir_del_registro()
  })

  # --------------------------- ADMINISTRACIÓN ---------------------------------
  tick_adm <- reactiveVal(0L)
  observeEvent(input$adm_refrescar, tick_adm(tick_adm() + 1L))

  avance_base <- reactive({
    tick_adm()
    reg <- q("SELECT id_registro, estado FROM registros")
    s <- slots_ext()
    s$estado <- reg$estado[match(s$id_registro, reg$id_registro)]
    s$estado[is.na(s$estado)] <- "pendiente"
    s
  })
  resumir <- function(s, por) {
    t <- table(factor(s[[por]]), factor(s$estado, levels = c("enviada", "borrador", "pendiente")))
    d <- data.frame(Grupo = rownames(t), Total = as.integer(rowSums(t)), Enviadas = as.integer(t[, "enviada"]),
                    Borradores = as.integer(t[, "borrador"]),
                    Pendientes = as.integer(t[, "pendiente"]), check.names = FALSE, stringsAsFactors = FALSE)
    d$`% completado` <- round(100 * d$Enviadas / d$Total, 0)
    d[order(d$`% completado`, d$Grupo), ]
  }
  output$adm_totales <- renderUI({
    s <- avance_base()
    n <- function(e) sum(s$estado == e)
    div(class = "mb-3",
        span(class = "pill pill-ok", sprintf("Enviadas: %d", n("enviada"))),
        span(class = "pill pill-pend", sprintf("Borradores: %d", n("borrador"))),
        span(class = "pill", style = "background:#e9ecef;", sprintf("Pendientes: %d", n("pendiente"))),
        span(class = "pill", style = "background:#cfe2ff;", sprintf("Total registros: %d", nrow(s))))
  })
  dt_opts <- list(pageLength = 15, dom = "ftip", language = list(search = "Buscar:", zeroRecords = "Sin resultados",
                  info = "_START_–_END_ de _TOTAL_", paginate = list(previous = "Anterior", `next` = "Siguiente")))
  output$adm_tabla_colegio <- renderDT({
    s <- avance_base()
    d <- resumir(s, "sigla")
    d$Nombre <- slots$nombre[match(d$Grupo, slots$sigla)]
    d$Gestora <- slots$gestora[match(d$Grupo, slots$sigla)]
    names(d)[1] <- "Colegio"
    datatable(d[, c("Colegio", "Nombre", "Gestora", "Total", "Enviadas", "Borradores", "Pendientes", "% completado")],
              rownames = FALSE, options = dt_opts) |>
      formatStyle("% completado", background = styleColorBar(c(0, 100), "#b6e0c4"), backgroundSize = "100% 70%",
                  backgroundRepeat = "no-repeat", backgroundPosition = "center")
  })
  output$adm_tabla_gestora <- renderDT({
    d <- resumir(avance_base(), "gestora"); names(d)[1] <- "Gestora"
    datatable(d, rownames = FALSE, options = c(dt_opts, list(dom = "t")))
  })

  rev_df <- reactive({
    tick_adm()
    r <- q("SELECT id_registro, estado, estado_revision, comentario_revision, habilitado_para, actualizado, actualizado_por FROM registros WHERE estado IN ('enviada','borrador')")
    S <- slots_ext()
    r <- r[order(match(r$id_registro, S$id_registro)), , drop = FALSE]
    s <- S[match(r$id_registro, S$id_registro), ]
    nc <- nombres_claves()
    data.frame(id = r$id_registro, Colegio = s$sigla, Curso = ifelse(s$grado == "K", "K", s$grado), Letra = s$letra,
               `Evaluación` = s$evaluacion, Toma = ifelse(is.na(s$toma), "—", as.character(s$toma)), Estado = r$estado,
               `Revisión` = r$estado_revision, Comentario = r$comentario_revision,
               `Corrección habilitada para` = local({
                 h <- r$habilitado_para; n <- nc$evaluador[match(h, nc$clave_id)]
                 ifelse(is.na(h), "", ifelse(is.na(n) | !nzchar(n), h, n)) }),
               Actualizado = fmt_fecha_hora(r$actualizado),
               Por = ifelse(is.na(match(r$actualizado_por, nc$clave_id)), r$actualizado_por,
                            nc$evaluador[match(r$actualizado_por, nc$clave_id)]),
               check.names = FALSE, stringsAsFactors = FALSE)
  })
  output$adm_tabla_rev <- renderDT({
    d <- rev_df()
    datatable(d[, -1, drop = FALSE], rownames = FALSE, selection = "multiple", filter = "top", options = dt_opts,
              callback = JS("$(table.table().node()).on('dblclick', 'tbody tr', function(){",
                            "  Shiny.setInputValue('adm_dbl', table.row(this).index() + 1, {priority: 'event'}); });"))
  })

  aplicar_revision <- function(ids, estado, comentario) {
    com <- na_chr(comentario)
    con_run(function(con) for (id in ids)
      exc(con, "UPDATE registros SET estado_revision = ?e, comentario_revision = ?cm WHERE id_registro = ?id",
          e = estado, cm = com, id = id))
    # "Requiere corrección" habilita al evaluador que envió; "Revisado" vuelve a bloquear
    if (identical(estado, "Requiere corrección")) habilitar_correccion(ids, TRUE, avisar = FALSE)
    if (identical(estado, "Revisado")) habilitar_correccion(ids, FALSE, avisar = FALSE)
    tick_adm(tick_adm() + 1L)
    showNotification(sprintf("%d registro(s) actualizados.", length(ids)), type = "message")
  }
  # Habilita (o bloquea de nuevo) la corrección de registros enviados, solo para quien los envió
  habilitar_correccion <- function(ids, habilitar = TRUE, avisar = TRUE) {
    n <- 0L
    con_run(function(con) for (id in ids) {
      n <<- n + if (habilitar)
        exc(con, "UPDATE registros SET habilitado_para = enviado_por WHERE id_registro = ?id AND estado = 'enviada' AND enviado_por IS NOT NULL AND enviado_por <> 'ADMIN'", id = id)
      else
        exc(con, "UPDATE registros SET habilitado_para = NULL WHERE id_registro = ?id", id = id)
    })
    tick_adm(tick_adm() + 1L)
    if (avisar) {
      if (habilitar && n < length(ids))
        showNotification(sprintf("%d habilitado(s). %d no se pudieron habilitar: no están enviados o los envió Administración.",
                                 n, length(ids) - n), type = "warning", duration = 8)
      else showNotification(sprintf("%d registro(s) %s.", n, if (habilitar) "habilitados para corrección" else "bloqueados de nuevo"),
                            type = "message")
    }
    invisible(n)
  }
  poner_revision <- function(estado) {
    sel <- input$adm_tabla_rev_rows_selected
    if (!length(sel)) { showNotification("Selecciona al menos un registro.", type = "warning"); return() }
    aplicar_revision(rev_df()$id[sel], estado, input$adm_comentario)
  }
  observeEvent(input$adm_rev_ok,   poner_revision("Revisado"))
  observeEvent(input$adm_rev_corr, poner_revision("Requiere corrección"))
  observeEvent(input$adm_rev_pend, poner_revision("Pendiente"))
  observeEvent(input$adm_hab, {
    sel <- input$adm_tabla_rev_rows_selected
    if (!length(sel)) { showNotification("Selecciona al menos un registro.", type = "warning"); return() }
    habilitar_correccion(rev_df()$id[sel], TRUE)
  })
  observeEvent(input$adm_deshab, {
    sel <- input$adm_tabla_rev_rows_selected
    if (!length(sel)) { showNotification("Selecciona al menos un registro.", type = "warning"); return() }
    habilitar_correccion(rev_df()$id[sel], FALSE)
  })

  # --- Ver el contenido de un registro (solo lectura) ---
  adm_id_abierto <- reactiveVal(NULL)
  mostrar_registro <- function(id) {
    d <- cargar_registro(id)
    reg <- d$reg
    if (is.null(reg)) { showNotification("Ese registro ya no existe.", type = "warning"); return() }
    S <- slots_ext(); s <- S[S$id_registro == id, , drop = FALSE]
    nc <- nombres_claves()
    quien <- function(cid) { n <- nc$evaluador[match(cid, nc$clave_id)]; ifelse(is.na(cid), "", ifelse(!is.na(n) & nzchar(n), n, cid)) }
    tabla <- function(df, titulo, nombres, declarados = NULL) {
      aviso <- if (!is.null(declarados) && !is.na(declarados) && declarados != nrow(df))
        span(class = "badge text-bg-danger ms-2", sprintf("No cuadra: declarados %d, registrados %d", declarados, nrow(df)))
      cab <- tags$h6(class = "mt-3 fw-bold", titulo, if (!is.null(declarados)) span(class = "text-muted fw-normal", sprintf(" · declarados: %s", ifelse(is.na(declarados), "—", declarados))), aviso)
      if (nrow(df) == 0) return(tagList(cab, p(class = "text-muted small", "Sin registros.")))
      tagList(cab, tags$table(class = "table table-sm table-striped",
        tags$thead(tags$tr(lapply(nombres, tags$th))),
        tags$tbody(lapply(seq_len(nrow(df)), function(i) tags$tr(lapply(df[i, , drop = FALSE], function(v) tags$td(if (is.na(v) || !nzchar(as.character(v))) "—" else as.character(v))))))))
    }
    hist <- q("SELECT rev, accion, clave_id, ts FROM historial WHERE id_registro = ?id ORDER BY rev", id = id)
    fecha <- if (vacio(reg$fecha_aplicacion)) "—" else format(as.Date(reg$fecha_aplicacion), "%d-%m-%Y")
    adm_id_abierto(id)
    showModal(modalDialog(
      size = "l", easyClose = TRUE, title = sprintf("%s · %s · %s%s", s$sigla, etiqueta_curso(s$grado, s$letra), s$evaluacion,
                                                     txt_toma(s$grado, s$toma)),
      div(class = "text-muted small mb-2",
          sprintf("%s · Modalidad: %s · Gestora: %s", s$nombre, s$modalidad, s$gestora)),
      div(class = "mb-2",
          span(class = paste("badge", if (reg$estado == "enviada") "text-bg-success" else "text-bg-warning"),
               if (reg$estado == "enviada") "Enviada" else "Borrador"),
          span(class = "badge text-bg-secondary ms-1", paste("Revisión:", reg$estado_revision)),
          span(class = "text-muted small ms-2",
               sprintf("Última actualización: %s por %s", fmt_fecha_hora(reg$actualizado), quien(reg$actualizado_por)))),
      if (!vacio(reg$comentario_revision)) div(class = "alert alert-warning py-2 small", strong("Comentario de revisión: "), reg$comentario_revision),
      tags$h6(class = "mt-3 fw-bold", "Datos generales"),
      tags$table(class = "table table-sm w-auto", tags$tbody(
        tags$tr(tags$th("Fecha de aplicación"), tags$td(fecha)),
        tags$tr(tags$th("Estudiantes presentes"), tags$td(ifelse(is.na(reg$presentes), "—", reg$presentes))))),
      tabla(d$ausentes,  "Estudiantes ausentes",  c("Apellidos", "Nombres"), reg$ausentes_decl),
      tabla(d$retirados, "Estudiantes retirados", c("RUT", "Apellidos", "Nombres"), reg$retirados_decl),
      tabla(transform(d$nuevos, rut_pendiente = ifelse(!is.na(rut_pendiente) & rut_pendiente == "1", "Sí", "")),
            "Estudiantes nuevos", c("RUT", "Apellidos", "Nombres", "RUT pendiente")),
      tabla(d$neep,      "Situaciones de estudiantes NEEP", c("RUT", "Apellidos", "Nombres", "Situación")),
      tags$h6(class = "mt-3 fw-bold", "Observaciones generales"),
      p(if (vacio(reg$observaciones)) "—" else reg$observaciones,
        if (!vacio(reg$detalle_obs)) tagList(tags$br(), span(style = "white-space:pre-wrap;", reg$detalle_obs))),
      tabla(d$excepcionales, "Situaciones excepcionales con estudiantes", c("Apellidos", "Nombres", "Situación ocurrida")),
      if (nrow(hist)) tagList(tags$h6(class = "mt-3 fw-bold", "Historial de envíos"),
        tags$ul(class = "small", lapply(seq_len(nrow(hist)), function(i)
          tags$li(sprintf("Versión %d · %s · %s · %s", hist$rev[i], hist$accion[i], quien(hist$clave_id[i]), fmt_fecha_hora(hist$ts[i])))))),
      tags$hr(),
      textInput("adm_m_comentario", "Comentario para el evaluador (opcional)", value = if (vacio(reg$comentario_revision)) "" else reg$comentario_revision, width = "100%"),
      footer = tagList(
        actionButton("adm_m_editar", "Editar registro", class = "btn-primary me-auto"),
        actionButton("adm_m_eliminar", "Eliminar", class = "btn-outline-danger"),
        actionButton("adm_m_hab",   "Habilitar corrección", class = "btn-info"),
        actionButton("adm_m_deshab", "Bloquear", class = "btn-outline-secondary"),
        actionButton("adm_m_ok",   "Marcar Revisado",      class = "btn-success"),
        actionButton("adm_m_corr", "Requiere corrección",  class = "btn-warning"),
        actionButton("adm_m_pend", "Volver a Pendiente",   class = "btn-outline-secondary"),
        modalButton("Cerrar"))))
  }
  observeEvent(input$adm_ver, {
    sel <- input$adm_tabla_rev_rows_selected
    if (length(sel) != 1) { showNotification("Selecciona un solo registro para ver su contenido.", type = "warning"); return() }
    mostrar_registro(rev_df()$id[sel])
  })
  observeEvent(input$adm_dbl, {
    d <- rev_df(); i <- input$adm_dbl
    if (!is.null(i) && i >= 1 && i <= nrow(d)) mostrar_registro(d$id[i])
  })
  editar_registro <- function(id) {
    removeModal()
    intentar_abrir(id)
  }
  observeEvent(input$adm_m_editar, { id <- adm_id_abierto(); req(id); editar_registro(id) })
  observeEvent(input$adm_editar_sel, {
    sel <- input$adm_tabla_rev_rows_selected
    if (length(sel) != 1) { showNotification("Selecciona un solo registro para editarlo.", type = "warning"); return() }
    editar_registro(rev_df()$id[sel])
  })
  # --- Eliminar registros (solo administración, con confirmación escrita) ---
  ids_eliminar <- reactiveVal(character())
  pedir_eliminar <- function(ids) {
    req(identical(ses$rol, "admin"))
    ids_eliminar(ids)
    S <- slots_ext()
    desc <- vapply(ids, function(id) {
      sl <- S[S$id_registro == id, , drop = FALSE]
      if (nrow(sl)) sprintf("%s · %s · %s%s", sl$sigla, etiqueta_curso(sl$grado, sl$letra), sl$evaluacion,
                            txt_toma(sl$grado, sl$toma)) else id
    }, "")
    showModal(modalDialog(
      title = "Eliminar registros",
      p(class = "text-danger", strong("Esta acción no se puede deshacer."),
        sprintf(" Se borrarán %d registro(s), con todo su detalle e historial de envíos:", length(ids))),
      tags$ul(class = "small", lapply(head(desc, 12), tags$li),
              if (length(desc) > 12) tags$li(sprintf("… y %d más", length(desc) - 12))),
      textInput("adm_texto_eliminar", "Para confirmar, escribe ELIMINAR", width = "100%"),
      footer = tagList(modalButton("Cancelar"),
                       actionButton("adm_conf_eliminar", "Eliminar definitivamente", class = "btn-danger"))))
  }
  observeEvent(input$adm_eliminar_sel, {
    sel <- input$adm_tabla_rev_rows_selected
    if (!length(sel)) { showNotification("Selecciona al menos un registro.", type = "warning"); return() }
    pedir_eliminar(rev_df()$id[sel])
  })
  observeEvent(input$adm_m_eliminar, { id <- adm_id_abierto(); req(id); pedir_eliminar(id) })
  observeEvent(input$adm_conf_eliminar, {
    req(identical(ses$rol, "admin"))
    if (!identical(trimws(input$adm_texto_eliminar %||% ""), "ELIMINAR")) {
      showNotification("Escribe ELIMINAR (en mayúsculas) para confirmar.", type = "warning"); return()
    }
    ids <- ids_eliminar(); req(length(ids) > 0)
    n <- eliminar_registros(ids)
    ids_eliminar(character())
    removeModal()
    tick_adm(tick_adm() + 1L); rv$tick <- rv$tick + 1L
    showNotification(sprintf("%d registro(s) eliminados.", n), type = "message")
  })

  revision_desde_modal <- function(estado) {
    id <- adm_id_abierto(); req(id)
    aplicar_revision(id, estado, input$adm_m_comentario)
    removeModal()
  }
  observeEvent(input$adm_m_hab,   { id <- adm_id_abierto(); req(id); habilitar_correccion(id, TRUE);  removeModal() })
  observeEvent(input$adm_m_deshab, { id <- adm_id_abierto(); req(id); habilitar_correccion(id, FALSE); removeModal() })
  observeEvent(input$adm_m_ok,   revision_desde_modal("Revisado"))
  observeEvent(input$adm_m_corr, revision_desde_modal("Requiere corrección"))
  observeEvent(input$adm_m_pend, revision_desde_modal("Pendiente"))

  # --- Claves ---
  nuevas_claves <- reactiveVal(NULL)
  observeEvent(input$adm_generar, {
    nombres <- trimws(strsplit(input$adm_nombres %||% "", "\n")[[1]])
    nombres <- nombres[nzchar(nombres)]
    if (!length(nombres)) {
      n <- suppressWarnings(as.integer(input$adm_n_claves))
      if (is.na(n) || n < 1 || n > 900) { showNotification("Cantidad inválida (1–900).", type = "error"); return() }
      nombres <- rep("", n)
    }
    res <- tryCatch(generar_claves(nombres), error = function(e) { showNotification(conditionMessage(e), type = "error"); NULL })
    req(res)
    nuevas_claves(res)
    tick_adm(tick_adm() + 1L)
    updateTextAreaInput(session, "adm_nombres", value = "")
  })
  output$adm_nuevas <- renderUI({
    d <- nuevas_claves(); req(d)
    div(class = "alert alert-warning",
        strong(sprintf("Se generaron %d claves. ", nrow(d))),
        "Descárgalas ahora: no se podrán volver a ver.",
        div(class = "mt-2", downloadButton("adm_desc_claves", "Descargar claves (Excel)", class = "btn-sm btn-warning")))
  })
  output$adm_desc_claves <- downloadHandler(
    filename = function() sprintf("claves%s_%s.xlsx", SLUG_PROCESO, format(Sys.time(), "%Y%m%d_%H%M", tz = TZ)),
    content = function(file) {
      d <- nuevas_claves()
      out <- data.frame(ID = d$clave_id, Evaluador = d$evaluador, Clave = d$clave)
      write.xlsx(out, file, headerStyle = createStyle(textDecoration = "bold"), colWidths = c(8, 36, 16))
    })
  claves_df <- reactive({
    tick_adm()
    q("SELECT clave_id, evaluador, activa, creada, ultimo_uso FROM claves ORDER BY clave_id")
  })
  output$adm_tabla_claves <- renderDT({
    d <- claves_df()
    d$activa <- ifelse(d$activa == 1, "Sí", "No")
    d$ultimo_uso <- ifelse(is.na(d$ultimo_uso), "", fmt_fecha_hora(d$ultimo_uso))
    d$creada <- fmt_fecha_hora(d$creada)
    names(d) <- c("ID", "Evaluador", "Activa", "Creada", "Último uso")
    datatable(d, rownames = FALSE, selection = "multiple", options = c(dt_opts, list(pageLength = 20)))
  })
  cambiar_activa <- function(valor) {
    sel <- input$adm_tabla_claves_rows_selected
    if (!length(sel)) { showNotification("Selecciona al menos una clave.", type = "warning"); return() }
    ids <- claves_df()$clave_id[sel]
    con_run(function(con) for (id in ids) exc(con, "UPDATE claves SET activa = ?a WHERE clave_id = ?id", a = valor, id = id))
    tick_adm(tick_adm() + 1L)
  }
  observeEvent(input$adm_clave_off, cambiar_activa(0L))
  observeEvent(input$adm_clave_on,  cambiar_activa(1L))

  # --- Descarga ---
  output$adm_descargar <- downloadHandler(
    filename = function() sprintf("registro_evaluaciones%s_%s.xlsx", SLUG_PROCESO, format(Sys.time(), "%Y%m%d_%H%M", tz = TZ)),
    content = function(file) exportar_excel(file, isTRUE(input$adm_borradores)))
}

shinyApp(ui, server)
