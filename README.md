# Automatización de Reportes de Ventas por Email

Sistema que genera reportes de ventas personalizados por vendedor y los envía automáticamente por correo electrónico con un archivo Excel adjunto.

## Características

- Lee datos de ventas desde un archivo Excel (hojas `Ventas` y `Mails`)
- Genera un email HTML con métricas clave por vendedor (total vendido, meta, avance %)
- Incluye versión en texto plano para evitar clasificación como spam
- Crea un Excel con el detalle de transacciones, con estilos profesionales
- Envía los reportes vía SMTP de Gmail con STARTTLS
- Valida variables de entorno al inicio y falla con mensaje claro si faltan
- Timeout configurable en la conexión SMTP para evitar colgados

## Vista previa del email

El email incluye:
- Nombre del vendedor y período del reporte
- Total de ventas vs. meta del equipo
- Porcentaje de avance con color dinámico (verde / naranja / rojo)
- Excel adjunto con el detalle de ventas ordenado por fecha

## Estructura del proyecto

```
mails-automaticos/
├── envio_reporte_ventas.py       # Script principal
├── reporte_ventas_vendedor.html  # Template del email HTML
├── Prueba_envio_de_mails.xlsx    # Datos de ejemplo (ventas + emails ficticios)
├── ejecutar_reporte.bat          # Script para programar con el Programador de tareas
├── requirements.txt              # Dependencias Python
├── .env.example                  # Plantilla de variables de entorno
├── .gitignore                    # Excluye .env, venv/, logs/
└── logs/                         # Logs de ejecución (generado automáticamente)
```

## Formato del Excel de datos

El archivo Excel debe tener dos hojas:

**Hoja `Ventas`**

| Fecha | Folio | Vendedor | Cliente | Monto |
|-------|-------|----------|---------|-------|
| fecha | nro   | id       | nombre  | CLP   |

**Hoja `Mails`**

| Vendedor | Nombre Vendedor | Mails |
|----------|-----------------|-------|
| id       | nombre          | email |

La columna `Vendedor` es la clave de unión entre ambas hojas.

## Instalación

```bash
# 1. Clonar el repositorio
git clone https://github.com/fabianmendezs/mails-automaticos-ventas.git
cd mails-automaticos-ventas

# 2. Crear entorno virtual e instalar dependencias
python -m venv venv
venv\Scripts\activate        # Windows
pip install -r requirements.txt

# 3. Configurar credenciales
copy .env.example .env
# Editar .env con tu correo y App Password de Gmail
```

## Configuración de Gmail

1. Activa la **verificación en dos pasos** en tu cuenta de Google
2. Ve a [Contraseñas de aplicación](https://myaccount.google.com/apppasswords)
3. Crea una nueva contraseña de aplicación (tipo "Correo")
4. Copia esa contraseña de 16 caracteres en `.env` como `SMTP_PASS`

> **Importante:** Usa siempre una App Password, nunca tu contraseña real de Gmail. La App Password se puede revocar independientemente sin afectar el acceso a la cuenta.

## Variables de entorno (`.env`)

```env
# Requeridas
SMTP_USER=tu_correo@gmail.com
SMTP_PASS=xxxx xxxx xxxx xxxx   # App Password de Gmail (16 caracteres)
REMITENTE=Reportes de Venta <tu_correo@gmail.com>

# Opcionales (valores por defecto indicados)
SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_TIMEOUT=30                 # Timeout de conexión SMTP en segundos
EXCEL_FILE=Prueba_envio_de_mails.xlsx
HTML_FILE=reporte_ventas_vendedor.html
META=100000
```

> El script valida `SMTP_USER` y `SMTP_PASS` al iniciar y termina con mensaje de error claro si no están definidas.

## Uso

```bash
# Activar entorno virtual
venv\Scripts\activate

# Ejecutar
python envio_reporte_ventas.py
```

Salida esperada:

```
───────────────────────────────────────────────────────
  Inicio envio: 26/04/2026 16:15:00
  Periodo: 01/04/2026 -> 05/04/2026
  Meta individual: $100,000
-------------------------------------------------------

  ✓  Fabian           $    34.210   34.2%  ->  vendedor1@gmail.com
  ✓  Cesar            $    28.950   29.0%  ->  vendedor2@gmail.com
  ✓  Barbara          $    41.680   41.7%  ->  vendedor3@gmail.com

-------------------------------------------------------
  Fin: 26/04/2026 16:15:05
-------------------------------------------------------
```

## Automatización con el Programador de tareas (Windows)

El archivo `ejecutar_reporte.bat` ejecuta el script y guarda el log en `logs/reporte.log`.

Para programarlo mensualmente:
1. Abrir **Programador de tareas** de Windows
2. Crear una tarea nueva → acción: ejecutar `ejecutar_reporte.bat`
3. Configurar el gatillo (ej. primer lunes de cada mes a las 08:00)

## Dependencias

| Librería | Uso |
|----------|-----|
| `pandas` | Lectura del Excel y manipulación de datos |
| `openpyxl` | Generación del Excel adjunto con estilos |
| `python-dotenv` | Carga de variables de entorno desde `.env` |

## Seguridad

### Credenciales
- Las credenciales se cargan desde `.env`, que está excluido del repositorio vía `.gitignore`
- Nunca subas tu `.env` ni contraseñas al repositorio
- Usa siempre **App Passwords** de Gmail, no tu contraseña real
- El archivo `Prueba_envio_de_mails.xlsx` contiene datos ficticios únicamente

### Deliverability (anti-spam)
El email se construye con estructura MIME correcta para maximizar entrega:

```
multipart/mixed
├── multipart/alternative
│   ├── text/plain    ← fallback para clientes que no renderizan HTML
│   └── text/html     ← versión principal
└── application/xlsx  ← adjunto
```

- **Texto plano incluido:** Gmail y otros clientes penalizan emails solo-HTML enviándolos a spam
- **STARTTLS:** la conexión SMTP se encripta antes de enviar credenciales
- **Timeout 30s:** la conexión SMTP falla limpiamente si el servidor no responde en 30 segundos

### Checklist antes de usar en producción

- [ ] `.env` creado con App Password de Gmail (no la contraseña real)
- [ ] Verificar que `.env` no aparece en `git status`
- [ ] Activar verificación en dos pasos en la cuenta de Gmail
- [ ] Reemplazar `Prueba_envio_de_mails.xlsx` con datos reales (sin commitear si contiene datos personales)

## Licencia

MIT
