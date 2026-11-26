# Inventario desde el celular (ejemplo)

Ejemplo básico para controlar desde el celular, en cualquier parte, un inventario que vive en un Excel del PC, y ver la cámara del PC.

- **Página del celular:** carpeta `docs`, publicada gratis con GitHub Pages.
- **Puente:** `puente/servidor.py` corre en el PC. Lee y escribe el Excel, toma la foto de la cámara y pide una clave para todo.
- **Túnel:** `cloudflared` crea una dirección `https://...trycloudflare.com` gratuita que llega al puente sin abrir puertos del router.

El Excel, la clave y los respaldos quedan en la carpeta `datos`, que no se sube a GitHub.

## Uso

1. Doble clic en `Iniciar puente.bat`. La primera vez crea `datos/inventario.xlsx` con productos de ejemplo.
2. Se abre una página con un código QR. Escanéelo con la cámara del celular.
3. En el iPhone, en Safari: Compartir > Agregar a pantalla de inicio. Queda como una app.

Mientras la ventana del puente esté abierta, desde el celular se puede:

- registrar lo producido, lo empacado (descuenta el granel y suma las unidades) y lo descontado;
- ver el inventario y los últimos movimientos, que se actualizan cada 5 segundos;
- ver la cámara del PC (se apaga sola si nadie la mira).

## Requisitos del PC

- Python 3 con `openpyxl`, `opencv-python` y `qrcode[pil]`.
- `cloudflared` en el PATH o en la carpeta `herramientas`.

## Cuidados

- Si el Excel está abierto en el PC, el registro avisa que hay que cerrarlo.
- Cada día, antes del primer cambio, se guarda una copia en `datos/respaldos`.
- `python puente/servidor.py --nueva-clave` cambia la clave. Los celulares conectados tienen que volver a escanear.
- La dirección del túnel cambia cada vez que se reinicia el puente. El puente la publica en `docs/servidor.json` y la página la busca sola.
