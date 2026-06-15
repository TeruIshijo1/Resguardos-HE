import sqlite3, os, uuid

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH  = os.path.join(BASE_DIR, 'database.db')

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn

def init_db():
    conn = get_db()
    conn.executescript('''
        -- ── EMPLEADOS ──────────────────────────────────────────
        CREATE TABLE IF NOT EXISTS empleados (
            id            TEXT PRIMARY KEY,
            nombre        TEXT NOT NULL,
            fecha_alta    TEXT,
            direccion     TEXT,
            area          TEXT,
            ubicacion     TEXT,
            telefono      TEXT,
            observaciones TEXT,
            cuentas       TEXT DEFAULT '{}',
            activo        INTEGER DEFAULT 1,
            archivo_alta  TEXT,
            fecha_nacimiento TEXT,
            created_at    TEXT DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS bajas (
            id            TEXT PRIMARY KEY,
            emp_id        TEXT,
            nombre        TEXT NOT NULL,
            fecha_alta    TEXT,
            fecha_baja    TEXT,
            motivo        TEXT,
            direccion     TEXT,
            area          TEXT,
            ubicacion     TEXT,
            telefono      TEXT,
            cuentas       TEXT DEFAULT '{}',
            equipos_snap  TEXT DEFAULT '[]',
            archivo_alta  TEXT,
            archivo_baja  TEXT,
            created_at    TEXT DEFAULT CURRENT_TIMESTAMP
        );

        -- ── INVENTARIO GENERAL ─────────────────────────────────
        CREATE TABLE IF NOT EXISTS inventario (
            id               TEXT PRIMARY KEY,
            tipo             TEXT NOT NULL,
            categoria        TEXT,
            marca            TEXT,
            modelo           TEXT,
            num_serie        TEXT,
            ip_id            TEXT,
            procesador       TEXT,
            ram              TEXT,
            almacenamiento   TEXT,
            licencia_windows TEXT,
            version_office   TEXT,
            cantidad_total   INTEGER DEFAULT 1,
            cantidad_libre   INTEGER DEFAULT 1,
            estado           TEXT DEFAULT 'libre',
            notas            TEXT,
            created_at       TEXT DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS asignaciones (
            id          TEXT PRIMARY KEY,
            emp_id      TEXT NOT NULL,
            inv_id      TEXT NOT NULL,
            cantidad    INTEGER DEFAULT 1,
            fecha_asig  TEXT,
            fecha_devol TEXT,
            activa      INTEGER DEFAULT 1,
            FOREIGN KEY(emp_id) REFERENCES empleados(id),
            FOREIGN KEY(inv_id) REFERENCES inventario(id)
        );

        -- ── IPs ────────────────────────────────────────────────
        CREATE TABLE IF NOT EXISTS ips (
            id          TEXT PRIMARY KEY,
            direccion   TEXT NOT NULL UNIQUE,
            subred      TEXT,
            estado      TEXT DEFAULT 'libre',
            asignada_a  TEXT,
            tipo_disp   TEXT,
            notas       TEXT,
            inv_id      TEXT,
            created_at  TEXT DEFAULT CURRENT_TIMESTAMP
        );

        -- ── IMPRESORAS ─────────────────────────────────────────
        CREATE TABLE IF NOT EXISTS impresoras (
            id                TEXT PRIMARY KEY,
            nombre            TEXT,
            marca             TEXT,
            modelo            TEXT,
            tipo_consumible   TEXT,
            modelo_consumible TEXT,
            ubicacion         TEXT,
            ip_id             TEXT,
            conexion          TEXT DEFAULT 'red',
            estado            TEXT DEFAULT 'activa',
            notas             TEXT,
            created_at        TEXT DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS impresora_equipos (
            id           TEXT PRIMARY KEY,
            imp_id       TEXT NOT NULL,
            inv_id       TEXT NOT NULL,
            tipo_conn    TEXT DEFAULT 'red',
            FOREIGN KEY(imp_id) REFERENCES impresoras(id),
            FOREIGN KEY(inv_id) REFERENCES inventario(id)
        );

        -- ── CÁMARAS ────────────────────────────────────────────
        CREATE TABLE IF NOT EXISTS camaras (
            id         TEXT PRIMARY KEY,
            area       TEXT NOT NULL,
            ubicacion  TEXT NOT NULL,
            marca      TEXT,
            modelo     TEXT,
            ip_id      TEXT,
            estado     TEXT DEFAULT 'activa',
            notas      TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );

        -- ── SERVIDORES / RED ───────────────────────────────────
        CREATE TABLE IF NOT EXISTS red_infra (
            id          TEXT PRIMARY KEY,
            tipo        TEXT NOT NULL,
            nombre      TEXT,
            marca       TEXT,
            modelo      TEXT,
            num_serie   TEXT,
            ip_id       TEXT,
            ubicacion   TEXT,
            puertos     INTEGER,
            notas       TEXT,
            estado      TEXT DEFAULT 'activo',
            created_at  TEXT DEFAULT CURRENT_TIMESTAMP
        );

        -- ── HOJAS DE SERVICIO ─────────────────────────────────
        CREATE TABLE IF NOT EXISTS hojas_servicio (
            id             TEXT PRIMARY KEY,
            impresora_id   TEXT NOT NULL,
            nombre_archivo TEXT NOT NULL,
            nombre_visible TEXT NOT NULL,
            fecha          TEXT NOT NULL,
            categoria      TEXT NOT NULL,
            notas          TEXT,
            created_at     TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(impresora_id) REFERENCES impresoras(id) ON DELETE CASCADE
        );
    ''')
    conn.commit()

    # 3. Pre-cargar IPs 192.168.254.x (Datos)
    existing_192 = conn.execute("SELECT COUNT(*) FROM ips WHERE subred='192.168.254.0/24'").fetchone()[0]
    if existing_192 == 0:
        rows = []
        for i in range(1, 255):
            rows.append((str(uuid.uuid4())[:8], f'192.168.254.{i}', '192.168.254.0/24', 'libre'))
        conn.executemany('INSERT INTO ips (id, direccion, subred, estado) VALUES (?,?,?,?)', rows)
        conn.commit()

    # 4. Pre-cargar IPs 172.16.90.x (Voz / Conmutadores)
    existing_172 = conn.execute("SELECT COUNT(*) FROM ips WHERE subred='172.16.90.0/24'").fetchone()[0]
    if existing_172 == 0:
        rows = []
        for i in range(1, 255):
            rows.append((str(uuid.uuid4())[:8], f'172.16.90.{i}', '172.16.90.0/24', 'libre'))
        conn.executemany('INSERT INTO ips (id, direccion, subred, estado) VALUES (?,?,?,?)', rows)
        conn.commit()

    conn.close()