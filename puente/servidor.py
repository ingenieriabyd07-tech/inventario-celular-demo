# -*- coding: utf-8 -*-
"""
Puente de ejemplo de Asepsia Total.

Deja controlar desde el celular, en cualquier parte, un inventario guardado en
un Excel de este PC y ver la cámara del PC:

  - el inventario vive en datos/inventario.xlsx (se crea con datos de ejemplo),
  - la página del celular está publicada gratis en GitHub Pages (carpeta docs),
  - un túnel gratuito de Cloudflare deja entrar al celular sin abrir puertos
    del router; todo lo que llega del celular necesita la clave del puente.

Uso:
  python servidor.py                 puente + túnel
  python servidor.py --sin-tunel     solo en este PC, para probar
  python servidor.py --nueva-clave   cambia la clave (los celulares viejos quedan fuera)
"""

# ---------------- Parámetros ----------------
PUERTO = 8765
INDICE_CAMARA = 0                 # 0 = cámara principal del PC
SEGUNDOS_APAGAR_CAMARA = 20       # se apaga si nadie la mira
PAGINA_CELULAR = "https://ingenieriabyd07-tech.github.io/inventario-celular-demo/"
PUBLICAR_DIRECCION_EN_GITHUB = True   # sube docs/servidor.json con la dirección del túnel
MOVIMIENTOS_EN_RESPUESTA = 15

import argparse
import datetime as dt
import base64
import hmac
import io
import json
import os
import re
import secrets
import shutil
import subprocess
import sys
import threading
import time
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import quote, urlparse

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("Falta openpyxl. Instálelo con: python -m pip install openpyxl")
try:
    import cv2
except ImportError:
    cv2 = None
try:
    import qrcode
except ImportError:
    qrcode = None

RAIZ = Path(__file__).resolve().parent.parent
DATOS = RAIZ / "datos"
EXCEL = DATOS / "inventario.xlsx"
CONFIG = DATOS / "config_local.json"
RESPALDOS = DATOS / "respaldos"
DOCS = RAIZ / "docs"

_u = urlparse(PAGINA_CELULAR)
ORIGENES = {f"{_u.scheme}://{_u.netloc}", f"http://localhost:{PUERTO}", f"http://127.0.0.1:{PUERTO}"}

ESTADO = {"tunel": "", "tunel_error": "", "inicio": dt.datetime.now()}
CLAVE = ""
LOCK_EXCEL = threading.Lock()


class ErrorUsuario(Exception):
    """Error que se muestra tal cual en el celular."""


def log(msg):
    print(dt.datetime.now().strftime("%H:%M:%S"), msg, flush=True)


# ---------------- Excel ----------------
ENC_INV = ["Código", "Producto", "Unidad", "Existencia", "Mínimo", "Sale de", "Contenido"]
ENC_MOV = ["Fecha", "Hora", "Tipo", "Código", "Producto", "Cantidad", "Unidad", "Nota", "Desde"]
# Código, producto, unidad, existencia, mínimo, granel del que sale, contenido por unidad
EJEMPLO = [
    ("DET", "Detergente lavadora (granel)", "L", 69, 20, None, None),
    ("DET-GAL", "Detergente lavadora, galón", "und", 14, 5, "DET", 3.785),
    ("DET-LIT", "Detergente lavadora, litro", "und", 22, 8, "DET", 1),
    ("CLO", "Cloro 5,25 % (granel)", "L", 46, 20, None, None),
    ("CLO-GAL", "Cloro 5,25 %, galón", "und", 3, 5, "CLO", 3.785),
    ("EMO", "Esencia morada", "ml", 900, 1000, None, None),
    ("PIS", "Cloro granulado piscina (granel)", "kg", 72, 45, None, None),
    ("PIS-LB", "Cloro granulado piscina, libra", "und", 14, 5, "PIS", 0.5),
    ("ESC", "Escoba", "und", 15, 4, None, None),
    ("TRA", "Trapero", "und", 3, 4, None, None),
]
TIPOS = {"producido": "Producido", "empacado": "Empacado", "descontado": "Descontado"}


def crear_excel():
    DATOS.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    inv = wb.active
    inv.title = "Inventario"
    inv.append(ENC_INV)
    for fila in EJEMPLO:
        inv.append(list(fila))
    mov = wb.create_sheet("Movimientos")
    mov.append(ENC_MOV)
    for hoja, anchos in ((inv, [11, 34, 9, 12, 10, 10, 11]), (mov, [12, 8, 12, 10, 34, 11, 9, 30, 10])):
        for i, ancho in enumerate(anchos, 1):
            hoja.column_dimensions[get_column_letter(i)].width = ancho
        for celda in hoja[1]:
            celda.font = Font(bold=True)
            celda.fill = PatternFill("solid", fgColor="E7ECEF")
        hoja.freeze_panes = "A2"
    for fila in inv.iter_rows(min_row=2):
        fila[3].number_format = "#,##0.00"
        fila[4].number_format = "#,##0.00"
        fila[6].number_format = "0.000"
    wb.save(EXCEL)


def numero(v, defecto=0.0):
    if v is None or v == "":
        return defecto
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip().replace(".", "").replace(",", ".")) if "," in str(v) else float(v)
    except ValueError:
        return defecto


def leer_productos(wb):
    lista = []
    for fila in wb["Inventario"].iter_rows(min_row=2, values_only=True):
        fila = (list(fila) + [None] * 7)[:7]
        if not fila[0]:
            continue
        lista.append({
            "codigo": str(fila[0]).strip(),
            "producto": str(fila[1] or ""),
            "unidad": str(fila[2] or ""),
            "existencia": round(numero(fila[3]), 4),
            "minimo": round(numero(fila[4]), 4),
            "origen": str(fila[5]).strip() if fila[5] else "",
            "contenido": numero(fila[6], None) if fila[6] not in (None, "") else None,
        })
    return lista


def leer_movimientos(wb, n=MOVIMIENTOS_EN_RESPUESTA):
    filas = [f for f in wb["Movimientos"].iter_rows(min_row=2, values_only=True) if f and f[0]]
    salida = []
    for f in reversed(filas[-n:]):
        f = (list(f) + [None] * 9)[:9]
        fecha = f[0].strftime("%Y-%m-%d") if hasattr(f[0], "strftime") else str(f[0])
        salida.append({"fecha": fecha, "hora": str(f[1] or ""), "tipo": str(f[2] or ""), "codigo": str(f[3] or ""),
                       "producto": str(f[4] or ""), "cantidad": numero(f[5]), "unidad": str(f[6] or ""),
                       "nota": str(f[7] or ""), "desde": str(f[8] or "")})
    return salida


def foto_inventario():
    with LOCK_EXCEL:
        wb = openpyxl.load_workbook(EXCEL)
        try:
            return {"productos": leer_productos(wb), "movimientos": leer_movimientos(wb),
                    "actualizado": dt.datetime.now().isoformat(timespec="seconds")}
        finally:
            wb.close()


def fmt(x):
    r = round(x, 2)
    txt = f"{r:,.2f}".rstrip("0").rstrip(".") if r != int(r) else f"{int(r):,}"
    return txt.replace(",", "§").replace(".", ",").replace("§", ".")


def respaldo_diario():
    RESPALDOS.mkdir(parents=True, exist_ok=True)
    destino = RESPALDOS / f"inventario_{dt.date.today().isoformat()}.xlsx"
    if not destino.exists():
        shutil.copy2(EXCEL, destino)


def guardar(wb):
    temporal = EXCEL.with_name("inventario.tmp.xlsx")
    wb.save(temporal)
    try:
        os.replace(temporal, EXCEL)
    except PermissionError:
        temporal.unlink(missing_ok=True)
        raise ErrorUsuario("El Excel está abierto en el PC. Ciérrelo y vuelva a intentar.")


def registrar(tipo, codigo, cantidad, nota, desde):
    if tipo not in TIPOS:
        raise ErrorUsuario("Tipo de movimiento desconocido.")
    if not isinstance(cantidad, (int, float)) or cantidad <= 0 or cantidad > 1e7:
        raise ErrorUsuario("La cantidad debe ser un número mayor que cero.")
    nota = str(nota or "")[:120]
    with LOCK_EXCEL:
        wb = openpyxl.load_workbook(EXCEL)
        try:
            inv = wb["Inventario"]
            filas = {str(r[0].value).strip(): r for r in inv.iter_rows(min_row=2) if r[0].value}
            if codigo not in filas:
                raise ErrorUsuario("Ese producto no está en el Excel.")
            fila = filas[codigo]
            unidad = str(fila[2].value or "")
            if unidad == "und" and cantidad != int(cantidad):
                raise ErrorUsuario("Este producto se cuenta en unidades enteras.")
            existencia = numero(fila[3].value)
            ahora = dt.datetime.now()
            renglones = []
            if tipo == "producido":
                fila[3].value = round(existencia + cantidad, 4)
                renglones.append([TIPOS[tipo], codigo, fila[1].value, cantidad, unidad, nota])
            elif tipo == "descontado":
                if cantidad > existencia + 1e-9:
                    raise ErrorUsuario(f"No alcanza: hay {fmt(existencia)} {unidad}.")
                fila[3].value = round(existencia - cantidad, 4)
                renglones.append([TIPOS[tipo], codigo, fila[1].value, -cantidad, unidad, nota])
            else:
                origen = str(fila[5].value or "").strip()
                contenido = numero(fila[6].value, 0)
                if not origen or origen not in filas or contenido <= 0:
                    raise ErrorUsuario("Este producto no se empaca desde un granel.")
                fo = filas[origen]
                usa = round(cantidad * contenido, 4)
                exi_o = numero(fo[3].value)
                if usa > exi_o + 1e-9:
                    raise ErrorUsuario(f"No alcanza el granel: hay {fmt(exi_o)} {fo[2].value} y se necesitan {fmt(usa)}.")
                fo[3].value = round(exi_o - usa, 4)
                fila[3].value = round(existencia + cantidad, 4)
                renglones.append([TIPOS[tipo], codigo, fila[1].value, cantidad, unidad, nota])
                renglones.append(["Empacado (sale)", origen, fo[1].value, -usa, fo[2].value, f"Para {fmt(cantidad)} {unidad} de {codigo}"])
            mov = wb["Movimientos"]
            for r in renglones:
                mov.append([ahora.date(), ahora.strftime("%H:%M")] + r + [desde])
                mov.cell(row=mov.max_row, column=1).number_format = "yyyy-mm-dd"
            respaldo_diario()
            guardar(wb)
            return {"productos": leer_productos(wb), "movimientos": leer_movimientos(wb),
                    "actualizado": ahora.isoformat(timespec="seconds")}
        finally:
            wb.close()


# ---------------- Cámara ----------------
class Camara:
    def __init__(self, indice):
        self.indice = indice
        self.cap = None
        self.cuadro = None
        self.uso = 0.0
        self.lock = threading.Lock()
        self.disponible = cv2 is not None

    def _abrir(self):
        cap = cv2.VideoCapture(self.indice, cv2.CAP_DSHOW) if os.name == "nt" else cv2.VideoCapture(self.indice)
        if not cap.isOpened():
            cap.release()
            return False
        cap.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
        cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)
        self.cap = cap
        threading.Thread(target=self._leer, daemon=True).start()
        log("Cámara encendida")
        return True

    def _leer(self):
        while True:
            with self.lock:
                cap = self.cap
                if cap is None:
                    return
                if time.time() - self.uso > SEGUNDOS_APAGAR_CAMARA:
                    cap.release()
                    self.cap = None
                    self.cuadro = None
                    log("Cámara apagada (nadie la está mirando)")
                    return
            ok, cuadro = cap.read()
            if ok:
                self.cuadro = cuadro
            else:
                time.sleep(0.05)

    def foto(self):
        if not self.disponible:
            return None
        self.uso = time.time()
        with self.lock:
            if self.cap is None and not self._abrir():
                return None
        t0 = time.time()
        while self.cuadro is None and time.time() - t0 < 5:
            time.sleep(0.05)
        cuadro = self.cuadro
        if cuadro is None:
            return None
        cuadro = cuadro.copy()
        sello = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        y = cuadro.shape[0] - 14
        cv2.putText(cuadro, sello, (12, y), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 0, 0), 3, cv2.LINE_AA)
        cv2.putText(cuadro, sello, (12, y), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (255, 255, 255), 1, cv2.LINE_AA)
        ok, buf = cv2.imencode(".jpg", cuadro, [cv2.IMWRITE_JPEG_QUALITY, 70])
        return buf.tobytes() if ok else None

    def cerrar(self):
        with self.lock:
            if self.cap is not None:
                self.cap.release()
                self.cap = None


CAMARA = Camara(INDICE_CAMARA)


# ---------------- Clave ----------------
def cargar_clave(nueva=False):
    DATOS.mkdir(parents=True, exist_ok=True)
    cfg = {}
    if CONFIG.exists():
        try:
            cfg = json.loads(CONFIG.read_text(encoding="utf-8"))
        except ValueError:
            cfg = {}
    if nueva or not cfg.get("clave"):
        cfg["clave"] = secrets.token_urlsafe(12)
        CONFIG.write_text(json.dumps(cfg, indent=2), encoding="utf-8")
        log("Clave nueva creada. Los celulares deben volver a escanear el código.")
    return cfg["clave"]


def enlace_celular():
    if not ESTADO["tunel"]:
        return ""
    return f"{PAGINA_CELULAR}#s={quote(ESTADO['tunel'], safe='')}&k={quote(CLAVE, safe='')}"


def enlace_local():
    return f"http://localhost:{PUERTO}/app/#s={quote(f'http://localhost:{PUERTO}', safe='')}&k={quote(CLAVE, safe='')}"


def qr_png_base64(texto):
    if not qrcode or not texto:
        return ""
    img = qrcode.make(texto, box_size=8, border=2)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return base64.b64encode(buf.getvalue()).decode("ascii")


def pagina_conectar():
    tunel = ESTADO["tunel"]
    qr = qr_png_base64(enlace_celular())
    if tunel:
        bloque = (f'<img src="data:image/png;base64,{qr}" alt="Código QR para el celular" width="300" height="300">'
                  f'<p>Escanee este código con la cámara del celular. Queda guardado y no hay que repetirlo, '
                  f'salvo que cambie la clave.</p><p class="mut">Túnel activo: {tunel}</p>')
    elif ESTADO["tunel_error"]:
        bloque = f'<p class="err">{ESTADO["tunel_error"]}</p>'
    else:
        bloque = '<p>Abriendo el túnel seguro... esta página se actualiza sola.</p>'
    recarga = "" if tunel or ESTADO["tunel_error"] else '<meta http-equiv="refresh" content="3">'
    return f"""<!doctype html><html lang="es"><head><meta charset="utf-8">{recarga}
<meta name="viewport" content="width=device-width, initial-scale=1"><title>Conectar el celular</title>
<style>body{{font-family:system-ui,Segoe UI,Arial,sans-serif;background:#f3f5f6;color:#13191d;margin:0;display:grid;place-items:center;min-height:100vh}}
main{{background:#fff;border:1px solid #d6dee3;border-radius:14px;padding:28px;max-width:460px;display:grid;gap:12px;justify-items:center;text-align:center}}
h1{{font-size:1.35rem;margin:0}}.mut{{color:#62707a;font-size:.88rem;word-break:break-all}}.err{{color:#ad2f1b;font-weight:600}}
a{{color:#1b5ca4}}</style></head><body><main>
<h1>Conectar el celular al puente</h1>{bloque}
<p class="mut">Para probar en este PC: <a href="{enlace_local()}">abrir la página aquí</a>.</p>
<p class="mut">No muestre este código a nadie más: quien lo escanee puede ver la cámara y mover el inventario.</p>
</main></body></html>"""


# ---------------- Servidor HTTP ----------------
TIPOS_ARCHIVO = {".html": "text/html; charset=utf-8", ".json": "application/json; charset=utf-8",
                 ".webmanifest": "application/manifest+json", ".png": "image/png", ".svg": "image/svg+xml",
                 ".js": "text/javascript; charset=utf-8", ".css": "text/css; charset=utf-8", ".ico": "image/x-icon"}


class Manejador(BaseHTTPRequestHandler):
    server_version = "PuenteAsepsia/1.0"

    def log_message(self, formato, *args):
        pass

    def _cors(self):
        origen = self.headers.get("Origin")
        if origen and origen in ORIGENES:
            self.send_header("Access-Control-Allow-Origin", origen)
            self.send_header("Vary", "Origin")

    def _enviar(self, codigo, cuerpo, tipo="application/json; charset=utf-8"):
        if isinstance(cuerpo, (dict, list)):
            cuerpo = json.dumps(cuerpo, ensure_ascii=False, default=str).encode("utf-8")
        elif isinstance(cuerpo, str):
            cuerpo = cuerpo.encode("utf-8")
        self.send_response(codigo)
        self.send_header("Content-Type", tipo)
        self.send_header("Content-Length", str(len(cuerpo)))
        self.send_header("Cache-Control", "no-store")
        self._cors()
        self.end_headers()
        self.wfile.write(cuerpo)

    def _es_local(self):
        # Lo que entra por el túnel trae cabeceras de Cloudflare.
        if self.headers.get("Cf-Connecting-Ip") or self.headers.get("Cf-Ray"):
            return False
        host = (self.headers.get("Host") or "").rsplit(":", 1)[0]
        return self.client_address[0] in ("127.0.0.1", "::1") and host in ("localhost", "127.0.0.1")

    def _autorizado(self):
        dada = (self.headers.get("X-Clave") or "").encode("utf-8", "ignore")
        if hmac.compare_digest(dada, CLAVE.encode("utf-8")):
            return True
        time.sleep(0.8)
        return False

    def _desde(self):
        return "Celular" if self.headers.get("Cf-Connecting-Ip") else "Este PC"

    def do_OPTIONS(self):
        self.send_response(204)
        self._cors()
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "X-Clave, Content-Type")
        self.send_header("Access-Control-Max-Age", "600")
        self.end_headers()

    def _estatico(self, ruta):
        relativa = ruta[len("/app"):].lstrip("/") or "index.html"
        archivo = (DOCS / relativa).resolve()
        if DOCS.resolve() not in archivo.parents or not archivo.is_file():
            return self._enviar(404, {"error": "No existe"})
        self._enviar(200, archivo.read_bytes(), TIPOS_ARCHIVO.get(archivo.suffix.lower(), "application/octet-stream"))

    def do_GET(self):
        ruta = urlparse(self.path).path
        try:
            if ruta in ("/", "/conectar"):
                if not self._es_local():
                    return self._enviar(403, {"error": "Solo desde este PC"})
                return self._enviar(200, pagina_conectar(), "text/html; charset=utf-8")
            if ruta == "/app" or ruta.startswith("/app/"):
                if ruta == "/app":
                    self.send_response(302)
                    self.send_header("Location", "/app/")
                    self.end_headers()
                    return
                return self._estatico(ruta)
            if not ruta.startswith("/api/"):
                return self._enviar(404, {"error": "No existe"})
            if not self._autorizado():
                return self._enviar(401, {"error": "Clave incorrecta"})
            if ruta == "/api/estado":
                return self._enviar(200, {"ok": True, "hora": dt.datetime.now().isoformat(timespec="seconds"),
                                          "excel": EXCEL.name, "camara": CAMARA.disponible})
            if ruta == "/api/inventario":
                return self._enviar(200, foto_inventario())
            if ruta == "/api/camara":
                imagen = CAMARA.foto()
                if not imagen:
                    return self._enviar(503, {"error": "No se pudo leer la cámara del PC."})
                return self._enviar(200, imagen, "image/jpeg")
            return self._enviar(404, {"error": "No existe"})
        except ErrorUsuario as e:
            return self._enviar(409, {"error": str(e)})
        except Exception as e:
            log(f"Error en {ruta}: {e!r}")
            return self._enviar(500, {"error": "Error en el puente del PC."})

    def do_POST(self):
        ruta = urlparse(self.path).path
        if ruta != "/api/movimiento":
            return self._enviar(404, {"error": "No existe"})
        if not self._autorizado():
            return self._enviar(401, {"error": "Clave incorrecta"})
        try:
            largo = int(self.headers.get("Content-Length") or 0)
            if largo <= 0 or largo > 10000:
                raise ErrorUsuario("Pedido vacío o demasiado grande.")
            datos = json.loads(self.rfile.read(largo).decode("utf-8"))
            cantidad = datos.get("cantidad")
            if isinstance(cantidad, str):
                cantidad = numero(cantidad, -1)
            resultado = registrar(str(datos.get("tipo", "")), str(datos.get("codigo", "")).strip(),
                                  cantidad, datos.get("nota", ""), self._desde())
            log(f"{self._desde()}: {datos.get('tipo')} {cantidad} de {datos.get('codigo')}")
            return self._enviar(200, resultado)
        except ErrorUsuario as e:
            return self._enviar(409, {"error": str(e)})
        except ValueError:
            return self._enviar(400, {"error": "Datos mal formados."})
        except Exception as e:
            log(f"Error al registrar: {e!r}")
            return self._enviar(500, {"error": "Error en el puente del PC."})


# ---------------- Túnel ----------------
def buscar_cloudflared():
    candidatos = [shutil.which("cloudflared"), str(RAIZ / "herramientas" / "cloudflared.exe")]
    for c in candidatos:
        if c and Path(c).is_file():
            return c
    return None


def iniciar_tunel(exe):
    patron = re.compile(r"https://[a-z0-9]+(?:-[a-z0-9]+)+\.trycloudflare\.com")
    banderas = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
    proc = subprocess.Popen([exe, "tunnel", "--no-autoupdate", "--url", f"http://127.0.0.1:{PUERTO}"],
                            stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True,
                            encoding="utf-8", errors="replace", creationflags=banderas)

    def leer():
        for linea in proc.stdout:
            m = patron.search(linea)
            if m and not ESTADO["tunel"]:
                ESTADO["tunel"] = m.group(0)
                log(f"Túnel activo: {ESTADO['tunel']}")
                threading.Thread(target=publicar_direccion, args=(ESTADO["tunel"],), daemon=True).start()
        if not ESTADO["tunel"]:
            ESTADO["tunel_error"] = "El túnel se cerró sin dar una dirección. Revise la conexión a internet."
            log(ESTADO["tunel_error"])

    threading.Thread(target=leer, daemon=True).start()
    return proc


def publicar_direccion(url):
    archivo = DOCS / "servidor.json"
    archivo.write_text(json.dumps({"url": url, "actualizado": dt.datetime.now().isoformat(timespec="seconds")},
                                  ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    if not PUBLICAR_DIRECCION_EN_GITHUB or not (RAIZ / ".git").exists():
        return

    def git(*args):
        return subprocess.run(["git", *args], cwd=RAIZ, capture_output=True, text=True, timeout=90,
                              creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0)

    git("add", "docs/servidor.json")
    if git("diff", "--cached", "--quiet").returncode == 0:
        return
    git("commit", "-m", "Actualiza la dirección del puente")
    r = git("push", "--quiet")
    if r.returncode == 0:
        log("Dirección publicada en GitHub. La página del celular la toma en uno o dos minutos.")
    else:
        log("No se pudo publicar la dirección en GitHub: " + (r.stderr or r.stdout).strip()[:200])


# ---------------- Arranque ----------------
def main():
    global CLAVE
    ap = argparse.ArgumentParser(description="Puente de ejemplo: Excel y cámara del PC para el celular.")
    ap.add_argument("--sin-tunel", action="store_true", help="solo en este PC, sin túnel")
    ap.add_argument("--nueva-clave", action="store_true", help="crea una clave nueva")
    ap.add_argument("--no-abrir", action="store_true", help="no abre el navegador al arrancar")
    args = ap.parse_args()

    if not EXCEL.exists():
        crear_excel()
        log(f"Creé el Excel de ejemplo: {EXCEL}")
    CLAVE = cargar_clave(args.nueva_clave)

    servidor = ThreadingHTTPServer(("127.0.0.1", PUERTO), Manejador)
    threading.Thread(target=servidor.serve_forever, daemon=True).start()
    log(f"Puente listo en http://localhost:{PUERTO}  (Excel: {EXCEL.name}, cámara: {'sí' if CAMARA.disponible else 'no'})")

    proc = None
    if not args.sin_tunel:
        exe = buscar_cloudflared()
        if exe:
            proc = iniciar_tunel(exe)
        else:
            ESTADO["tunel_error"] = ("No encontré cloudflared. Póngalo en la carpeta herramientas "
                                     "o instálelo con: winget install Cloudflare.cloudflared")
            log(ESTADO["tunel_error"])
    if not args.no_abrir:
        webbrowser.open(f"http://localhost:{PUERTO}/conectar")
    log("Deje esta ventana abierta. Para apagar el puente: Ctrl + C.")
    try:
        while True:
            time.sleep(1)
            if proc is not None and proc.poll() is not None and ESTADO["tunel"]:
                log("El túnel se cerró. Reinicie el puente.")
                ESTADO["tunel"] = ""
                ESTADO["tunel_error"] = "El túnel se cerró. Reinicie el puente."
    except KeyboardInterrupt:
        log("Apagando...")
    finally:
        if proc is not None:
            proc.terminate()
        CAMARA.cerrar()
        servidor.shutdown()


if __name__ == "__main__":
    main()
