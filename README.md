# Automatización de Reportes de Ventas por Email

Sistema que genera reportes de ventas personalizados por vendedor y los envía automáticamente por correo electrónico con un archivo Excel adjunto.

## Características

- Lee datos de ventas desde un archivo Excel
- Genera un email HTML con métricas clave por vendedor (total vendido, meta, avance %)
- Crea un Excel con el detalle de transacciones, con estilos profesionales
- Envía los reportes vía SMTP de Gmail
- Registro de ejecución en log con estado de cada envío

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
├── reporte_ventas_vendedor.html  # Template del email
├── Prueba_envio_de_mails.xlsx    # Datos de ejemplo (ventas + emails)
├── ejecutar_reporte.bat          # Script para programar con el Programador de tareas
├── requirements.txt              # Dependencias Python
├── .env.example                  # Plantilla de variables de entorno
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
git clone https://github.com/fabian-mendez/mails-automaticos.git
cd mails-automaticos

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

## Variables de entorno (`.env`)

```env
SMTP_USER=tu_correo@gmail.com
SMTP_PASS=xxxx xxxx xxxx xxxx   # App Password de Gmail
REMITENTE=Reportes de Venta <tu_correo@gmail.com>

# Opcionales
SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
EXCEL_FILE=Prueba_envio_de_mails.xlsx
HTML_FILE=reporte_ventas_vendedor.html
META=100000
```

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
  Inicio envío: 26/04/2026 16:15:00
  Período: 01/04/2026 → 05/04/2026
  Meta individual: $100,000
-------------------------------------------------------

  ✓  Fabian           $    34.210   34.2%  →  vendedor1@gmail.com
  ✓  Cesar            $    28.950   29.0%  →  vendedor2@gmail.com
  ✓  Barbara          $    41.680   41.7%  →  vendedor3@gmail.com

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

- Las credenciales se cargan desde `.env`, que está excluido del repositorio vía `.gitignore`
- Nunca subas tu `.env` ni contraseñas al repositorio
- Usa siempre **App Passwords** de Gmail, no tu contraseña real

## Licencia

MIT
