import urllib.request
import ssl
import os
import sys
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import subprocess
from datetime import datetime

# --- CONFIGURACIÓN DE CORREO GMAIL ---
# Leídos de variables de entorno para ejecución segura en Actions, con fallback para local
GMAIL_USER = os.getenv('GMAIL_USER', 'cesartroyasherdek@gmail.com')
GMAIL_APP_PASSWORD = os.getenv('GMAIL_APP_PASSWORD', 'ekvo gsau yfsc iatp')
RECEIVER_USER = os.getenv('RECEIVER_USER', 'cesar.troya@gm.com')

# --- RUTAS DE ARCHIVOS ---
# Modificado para ser totalmente relativo al directorio donde se ejecuta el script
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_PATH = os.path.join(BASE_DIR, 'SRI_Vehiculos_Nuevos_2026.csv')
SIZE_FILE = os.path.join(BASE_DIR, 'last_size.txt')
LOG_FILE = os.path.join(BASE_DIR, 'change_history.log')
LAST_RUN_TIME_FILE = os.path.join(BASE_DIR, 'last_run_time.txt')
REPORT_HTML_PATH = os.path.join(BASE_DIR, 'index.html')

URL = 'https://descargas.sri.gob.ec/download/datosAbiertos/SRI_Vehiculos_Nuevos_2026.csv'

def generate_html_report():
    if not os.path.exists(LOG_FILE):
        return

    try:
        with open(LOG_FILE, 'r', encoding='utf-8') as f:
            lines = f.readlines()

        logs = []
        for line in lines:
            line = line.strip()
            if not line:
                continue
            if line.startswith('[') and ']' in line:
                parts = line.split(']', 1)
                timestamp = parts[0][1:]
                message = parts[1].strip()
                
                badge_class = "badge-gray"
                if "Incremento" in message or "Descargando" in message or "enviado" in message:
                    badge_class = "badge-green"
                elif "Inicializando" in message:
                    badge_class = "badge-blue"
                elif "Error" in message or "Fallo" in message or "abortado" in message:
                    badge_class = "badge-red"
                
                logs.append({
                    "timestamp": timestamp,
                    "message": message,
                    "badge_class": badge_class
                })

        logs.reverse()

        last_check = logs[0]["timestamp"] if logs else "Nunca"
        last_status = logs[0]["message"] if logs else "Sin registros"

        # Calcular próxima ejecución esperada (frecuencia de 5 min)
        from datetime import datetime, timedelta
        next_check = "Desconocido"
        next_check_subtext = "Programación activa"
        
        # Consultar estado de la tarea programada local (si es Windows)
        task_state = "Ready"
        if sys.platform == 'win32':
            try:
                res = subprocess.run(["powershell", "-Command", "(Get-ScheduledTask -TaskName 'MonitorSRIVehiculosWeb').State"], capture_output=True, text=True)
                task_state = res.stdout.strip()
            except Exception:
                pass
        else:
            task_state = "GitHub Actions"
            
        if task_state == "Disabled":
            next_check = "Pausado"
            next_check_subtext = "Monitoreo detenido por incremento"
        elif last_check != "Nunca":
            try:
                last_check_dt = datetime.strptime(last_check, '%Y-%m-%d %H:%M:%S')
                next_check_dt = last_check_dt + timedelta(minutes=5)
                next_check_subtext = "Frecuencia: cada 5 min"
                next_check = next_check_dt.strftime('%Y-%m-%d %H:%M:%S')
            except Exception as ex:
                next_check_subtext = f"Error al calcular: {ex}"
        
        current_size = "Desconocido"
        if os.path.exists(SIZE_FILE):
            with open(SIZE_FILE, 'r') as f:
                sz = f.read().strip()
                if sz.isdigit():
                    current_size = f"{int(sz)/1024/1024:.2f} MB ({int(sz):,} bytes)"

        rows_html = ""
        for log in logs[:50]:
            badge_text = "INFO"
            if log["badge_class"] == "badge-green":
                badge_text = "ALERTA/CAMBIO"
            elif log["badge_class"] == "badge-blue":
                badge_text = "INICIALIZACIÓN"
            elif log["badge_class"] == "badge-red":
                badge_text = "ERROR"
            else:
                badge_text = "CHEQUEO OK"
                
            rows_html += f"""
            <tr>
                <td>{log["timestamp"]}</td>
                <td><span class="badge {log["badge_class"]}">{badge_text}</span></td>
                <td>{log["message"]}</td>
            </tr>
            """

        html_content = f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="refresh" content="60">
    <title>Reporte de Ejecuciones SRI - Web</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap" rel="stylesheet">
    <style>
        :root {{
            --bg-color: #0b0f19;
            --card-bg: rgba(255, 255, 255, 0.03);
            --card-border: rgba(255, 255, 255, 0.08);
            --text-primary: #f8fafc;
            --text-secondary: #94a3b8;
            --green: #10b981;
            --blue: #3b82f6;
            --red: #ef4444;
            --gray: #64748b;
        }}
        
        body {{
            font-family: 'Inter', sans-serif;
            background-color: var(--bg-color);
            color: var(--text-primary);
            margin: 0;
            padding: 2rem;
            min-height: 100vh;
        }}

        .container {{
            max-width: 1000px;
            margin: 0 auto;
        }}

        header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 2rem;
            border-bottom: 1px solid var(--card-border);
            padding-bottom: 1rem;
        }}

        h1 {{
            font-size: 1.8rem;
            font-weight: 700;
            margin: 0;
            background: linear-gradient(135deg, #3b82f6, #10b981);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }}

        .last-update-header {{
            font-size: 0.85rem;
            color: var(--text-secondary);
        }}

        /* Dashboard Cards */
        .dashboard-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
            gap: 1.5rem;
            margin-bottom: 2.5rem;
        }}

        .card {{
            background: var(--card-bg);
            border: 1px solid var(--card-border);
            border-radius: 12px;
            padding: 1.5rem;
            backdrop-filter: blur(12px);
            transition: transform 0.2s ease, border-color 0.2s ease;
        }}

        .card:hover {{
            transform: translateY(-2px);
            border-color: rgba(255, 255, 255, 0.15);
        }}

        .card-title {{
            font-size: 0.85rem;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            color: var(--text-secondary);
            margin-bottom: 0.5rem;
            font-weight: 600;
        }}

        .card-value {{
            font-size: 1.25rem;
            font-weight: 700;
            color: var(--text-primary);
        }}

        .card-subtext {{
            font-size: 0.8rem;
            color: var(--text-secondary);
            margin-top: 0.25rem;
        }}

        /* Log Table */
        .table-container {{
            background: var(--card-bg);
            border: 1px solid var(--card-border);
            border-radius: 12px;
            overflow: hidden;
            box-shadow: 0 4px 30px rgba(0, 0, 0, 0.5);
        }}

        table {{
            width: 100%;
            border-collapse: collapse;
            text-align: left;
            font-size: 0.9rem;
        }}

        th, td {{
            padding: 1rem 1.5rem;
            border-bottom: 1px solid var(--card-border);
        }}

        th {{
            background: rgba(0, 0, 0, 0.2);
            color: var(--text-primary);
            font-weight: 600;
            text-transform: uppercase;
            font-size: 0.75rem;
            letter-spacing: 0.05em;
        }}

        tr:last-child td {{
            border-bottom: none;
        }}

        tr:hover td {{
            background: rgba(255, 255, 255, 0.015);
        }}

        /* Badges */
        .badge {{
            display: inline-block;
            padding: 0.25rem 0.6rem;
            border-radius: 6px;
            font-size: 0.75rem;
            font-weight: 700;
            letter-spacing: 0.02em;
        }}

        .badge-gray {{
            background: rgba(100, 116, 139, 0.15);
            color: #cbd5e1;
            border: 1px solid rgba(100, 116, 139, 0.3);
        }}

        .badge-green {{
            background: rgba(16, 185, 129, 0.15);
            color: #34d399;
            border: 1px solid rgba(16, 185, 129, 0.3);
        }}

        .badge-blue {{
            background: rgba(59, 130, 246, 0.15);
            color: #60a5fa;
            border: 1px solid rgba(59, 130, 246, 0.3);
        }}

        .badge-red {{
            background: rgba(239, 68, 68, 0.15);
            color: #f87171;
            border: 1px solid rgba(239, 68, 68, 0.3);
        }}

        /* Botón de Ejecución Forzada */
        .btn-force {{
            background: linear-gradient(135deg, #3b82f6, #1d4ed8);
            color: white;
            border: none;
            padding: 0.6rem 1.2rem;
            border-radius: 8px;
            font-weight: 600;
            cursor: pointer;
            transition: transform 0.1s ease, filter 0.2s ease;
            font-size: 0.85rem;
            box-shadow: 0 4px 12px rgba(59, 130, 246, 0.3);
        }}
        .btn-force:hover {{
            filter: brightness(1.1);
            transform: translateY(-1px);
        }}
        .btn-force:active {{
            transform: translateY(0);
        }}

        /* Modal / Overlay */
        .modal-overlay {{
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(15, 23, 42, 0.85);
            backdrop-filter: blur(8px);
            justify-content: center;
            align-items: center;
            z-index: 1000;
        }}
        .modal-content {{
            background: #1e293b;
            border: 1px solid var(--card-border);
            border-radius: 16px;
            padding: 2rem;
            max-width: 450px;
            text-align: center;
            box-shadow: 0 10px 25px rgba(0, 0, 0, 0.5);
            animation: scaleUp 0.3s ease-out;
        }}
        @keyframes scaleUp {{
            from {{ transform: scale(0.9); opacity: 0; }}
            to {{ transform: scale(1); opacity: 1; }}
        }}
        .modal-title {{
            font-size: 1.25rem;
            font-weight: 700;
            margin-top: 0;
            color: #f8fafc;
            margin-bottom: 0.75rem;
        }}
        .modal-body {{
            font-size: 0.9rem;
            color: var(--text-secondary);
            line-height: 1.5;
            margin: 1rem 0 1.5rem 0;
            text-align: left;
        }}
        .btn-close {{
            background: var(--blue);
            color: white;
            border: none;
            padding: 0.6rem 1.8rem;
            border-radius: 8px;
            font-weight: 600;
            cursor: pointer;
            font-size: 0.85rem;
            transition: background-color 0.2s;
        }}
        .btn-close:hover {{
            background-color: #1d4ed8;
        }}
    </style>
</head>
<body>
    <div class="container">
        <header>
            <div>
                <h1>Reporte de Monitoreo - SRI (Web Version)</h1>
                <div class="last-update-header">Reporte generado automáticamente desde GitHub Actions o ejecución local.</div>
            </div>
            <div>
                <button class="btn-force" onclick="showModal()">Forzar Ejecución</button>
            </div>
        </header>

        <div class="dashboard-grid">
            <div class="card">
                <div class="card-title">Última Verificación</div>
                <div class="card-value">{last_check}</div>
                <div class="card-subtext">Reloj del sistema ejecutor</div>
            </div>
            <div class="card">
                <div class="card-title">Último Estado Registrado</div>
                <div class="card-value" style="font-size: 1.05rem;">{last_status}</div>
            </div>
            <div class="card">
                <div class="card-title">Tamaño Registrado del Archivo</div>
                <div class="card-value">{current_size}</div>
                <div class="card-subtext">Solo aumenta ante nuevas inserciones</div>
            </div>
            <div class="card" style="border-color: rgba(59, 130, 246, 0.2); background: rgba(59, 130, 246, 0.01);">
                <div class="card-title" style="color: var(--blue);">Próxima Verificación Programada</div>
                <div class="card-value" style="color: #60a5fa;" id="next-check-time">{next_check}</div>
                <div class="card-subtext" style="color: var(--text-secondary); display: flex; justify-content: space-between; align-items: center; margin-top: 0.5rem;">
                    <span id="next-check-sub">{next_check_subtext}</span>
                    <span id="countdown" style="font-weight: 700; color: #60a5fa; background: rgba(59, 130, 246, 0.15); padding: 0.15rem 0.5rem; border-radius: 4px; font-size: 0.75rem; min-width: 70px; text-align: center;">--:--</span>
                </div>
            </div>
        </div>

        <h2 style="font-size: 1.2rem; margin-bottom: 1rem; font-weight: 600;">Historial de Ejecuciones (Últimas 50)</h2>
        <div class="table-container">
            <table>
                <thead>
                    <tr>
                        <th style="width: 200px;">Fecha y Hora</th>
                        <th style="width: 150px;">Tipo de Evento</th>
                        <th>Mensaje de Log</th>
                    </tr>
                </thead>
                <tbody>
                    {rows_html}
                </tbody>
            </table>
        </div>
    </div>

    <!-- Modal para ejecución manual -->
    <div id="forceModal" class="modal-overlay" onclick="closeModal(event)">
        <div class="modal-content" onclick="event.stopPropagation()">
            <div class="modal-title">¿Cómo forzar la ejecución?</div>
            <div class="modal-body">
                Si estás en local:<br>
                1. Abre la carpeta del proyecto local.<br>
                2. Haz doble clic sobre <strong>Reiniciar_Monitor.bat</strong>.<br><br>
                Si deseas forzarlo en la nube:<br>
                1. Ve a tu repositorio en GitHub.<br>
                2. Entra a la pestaña <strong>Actions</strong>.<br>
                3. Selecciona el workflow <strong>SRI Monitor Web</strong> y presiona <strong>Run workflow</strong>.
            </div>
            <button class="btn-close" onclick="closeModal(event)">Entendido</button>
        </div>
    </div>

    <script>
        function showModal() {{
            document.getElementById('forceModal').style.display = 'flex';
        }}
        function closeModal(event) {{
            document.getElementById('forceModal').style.display = 'none';
        }}

        // Countdown Timer Logic
        function startCountdown() {{
            const nextCheckStr = document.getElementById('next-check-time').innerText;
            if (nextCheckStr === "Desconocido" || nextCheckStr.includes("Nunca") || nextCheckStr === "Pausado" || nextCheckStr.includes("GitHub")) return;
            
            const parts = nextCheckStr.split(' ');
            const dateParts = parts[0].split('-');
            const timeParts = parts[1].split(':');
            
            const targetDate = new Date(
                parseInt(dateParts[0]),
                parseInt(dateParts[1]) - 1,
                parseInt(dateParts[2]),
                parseInt(timeParts[0]),
                parseInt(timeParts[1]),
                parseInt(timeParts[2])
            );
            
            const countdownEl = document.getElementById('countdown');
            
            function update() {{
                const now = new Date();
                const diffMs = targetDate - now;
                
                if (diffMs <= 0) {{
                    countdownEl.innerText = "Ejecutando...";
                    countdownEl.style.background = "rgba(16, 185, 129, 0.15)";
                    countdownEl.style.color = "#34d399";
                    
                    if (!window.reloadScheduled) {{
                        window.reloadScheduled = true;
                        setTimeout(function() {{
                            location.reload();
                        }}, 15000);
                    }}
                    return;
                }}
                
                const totalSeconds = Math.floor(diffMs / 1000);
                const hours = Math.floor(totalSeconds / 3600);
                const minutes = Math.floor((totalSeconds % 3600) / 60);
                const seconds = totalSeconds % 60;
                
                let timeStr = "";
                if (hours > 0) {{
                    timeStr += hours + "h ";
                }}
                if (minutes > 0 || hours > 0) {{
                    timeStr += minutes + "m ";
                }}
                timeStr += seconds + "s";
                
                countdownEl.innerText = "En " + timeStr;
            }}
            
            update();
            setInterval(update, 1000);
        }}
        
        window.addEventListener('DOMContentLoaded', startCountdown);
    </script>
</body>
</html>
"""
        with open(REPORT_HTML_PATH, 'w', encoding='utf-8') as f:
            f.write(html_content)
    except Exception as e:
        print(f"Error al generar reporte HTML: {e}")

def write_log(message):
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    log_message = f"[{timestamp}] {message}\n"
    print(log_message.strip())
    with open(LOG_FILE, 'a', encoding='utf-8') as f:
        f.write(log_message)
    generate_html_report()

def send_windows_notification(title, message):
    if sys.platform != 'win32':
        return
    # Genera una notificación nativa en Windows usando PowerShell
    ps_script = f'''
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    $objNotifyIcon = New-Object System.Windows.Forms.NotifyIcon
    $objNotifyIcon.Icon = [System.Drawing.SystemIcons]::Information
    $objNotifyIcon.BalloonTipIcon = "Info"
    $objNotifyIcon.BalloonTipTitle = "{title}"
    $objNotifyIcon.BalloonTipText = "{message}"
    $objNotifyIcon.Visible = $True
    $objNotifyIcon.ShowBalloonTip(10000)
    '''
    try:
        subprocess.run(["powershell", "-Command", ps_script], capture_output=True)
    except Exception as e:
        write_log(f"No se pudo mostrar la notificación de Windows: {e}")

def send_email_notification(old_size, new_size, month, year, total_sales, table):
    if not GMAIL_USER or not GMAIL_APP_PASSWORD:
        write_log("Configuración de correo no establecida o incompleta. Saltando envío de email.")
        return

    subject = f"Resultados SRI {month} {year}".strip()
    if not year:
        subject = f"Resultados SRI {month}".strip()

    body = (
        f"Se ha detectado un incremento en el archivo de vehículos nuevos del SRI (Versión Web/Actions).\n\n"
        f"Detalles de la descarga:\n"
        f"- Tamaño anterior: {old_size / 1024 / 1024:.2f} MB ({old_size} bytes)\n"
        f"- Tamaño nuevo: {new_size / 1024 / 1024:.2f} MB ({new_size} bytes)\n"
        f"- Ruta: {CSV_PATH}\n\n"
        f"Se ha ejecutado el script de cálculo y consolidación. El archivo Excel resultante se guardó en:\n"
        f"{os.path.join(BASE_DIR, 'Procesador_SRI_Result.xlsx')}\n\n"
        f"==================================================\n"
        f"RESUMEN DEL ÚLTIMO MES DISPONIBLE: {month.upper()} {year}\n"
        f"Total de ventas del mes: {total_sales}\n"
        f"==================================================\n"
        f"{table}\n"
        f"==================================================\n"
    )

    msg = MIMEMultipart()
    msg['From'] = GMAIL_USER
    msg['To'] = RECEIVER_USER
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    try:
        write_log("Enviando correo de notificación...")
        with smtplib.SMTP_SSL('smtp.gmail.com', 465, timeout=15) as server:
            server.login(GMAIL_USER, GMAIL_APP_PASSWORD)
            server.sendmail(GMAIL_USER, RECEIVER_USER, msg.as_string())
        write_log("Correo enviado con éxito.")
    except Exception as e:
        write_log(f"Error al enviar el correo: {e}")


def download_file(ctx):
    req = urllib.request.Request(URL, headers={'User-Agent': 'Mozilla/5.0'})
    os.makedirs(os.path.dirname(CSV_PATH), exist_ok=True)
    with urllib.request.urlopen(req, context=ctx, timeout=30) as response:
        with open(CSV_PATH, 'wb') as out_file:
            out_file.write(response.read())

LOCK_FILE = os.path.join(BASE_DIR, 'sri_monitor.lock')

def is_process_alive(pid):
    if sys.platform == 'win32':
        try:
            res = subprocess.run(["tasklist", "/FI", f"PID eq {pid}"], capture_output=True, text=True, errors='ignore')
            return str(pid) in res.stdout
        except:
            return False
    else:
        try:
            os.kill(pid, 0)
            return True
        except OSError:
            return False

def kill_process(pid):
    if sys.platform == 'win32':
        try:
            subprocess.run(["taskkill", "/F", "/PID", str(pid)], capture_output=True)
        except:
            pass
    else:
        try:
            import signal
            os.kill(pid, signal.SIGKILL)
        except:
            pass

def acquire_lock():
    import time
    try:
        if os.path.exists(LOCK_FILE):
            try:
                with open(LOCK_FILE, 'r') as f:
                    content = f.read().strip()
                if content:
                    parts = content.split(',')
                    old_pid = int(parts[0])
                    start_time_epoch = float(parts[1])
                    
                    if is_process_alive(old_pid):
                        if time.time() - start_time_epoch > 600:
                            write_log(f"[WATCHDOG] Detectada instancia previa colgada (PID: {old_pid}). Forzando terminación para reiniciar...")
                            kill_process(old_pid)
                            time.sleep(2)
                        else:
                            print("Otra verificación de SRI ya se encuentra en ejecución activa. Saliendo.")
                            sys.exit(0)
            except Exception as ex:
                write_log(f"[WATCHDOG] Error al validar archivo de bloqueo existente: {ex}")
                try:
                    os.remove(LOCK_FILE)
                except:
                    pass
                    
        with open(LOCK_FILE, 'w') as f:
            f.write(f"{os.getpid()},{time.time()}")
        return True
    except Exception as e:
        write_log(f"[WATCHDOG] Error grave al adquirir bloqueo de ejecución: {e}")
        sys.exit(1)

def release_lock():
    try:
        if os.path.exists(LOCK_FILE):
            os.remove(LOCK_FILE)
    except:
        pass

def ensure_task_enabled():
    if sys.platform != 'win32':
        return
    try:
        res = subprocess.run(["powershell", "-Command", "(Get-ScheduledTask -TaskName 'MonitorSRIVehiculosWeb').State"], capture_output=True, text=True)
        state = res.stdout.strip()
        if state == "Disabled":
            write_log("[WATCHDOG] La tarea programada 'MonitorSRIVehiculosWeb' estaba deshabilitada. Reactivándola...")
            subprocess.run(["powershell", "-Command", "Enable-ScheduledTask -TaskName 'MonitorSRIVehiculosWeb'"], capture_output=True)
    except Exception as e:
        write_log(f"[WATCHDOG] Error al verificar/habilitar la tarea programada: {e}")

def main():
    ensure_task_enabled()
    acquire_lock()
    
    now = datetime.now()
    
    try:
        with open(LAST_RUN_TIME_FILE, 'w') as f:
            f.write(now.strftime('%Y-%m-%d %H:%M:%S'))
    except Exception as e:
        write_log(f"Error al registrar última ejecución: {e}")

    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE

    req = urllib.request.Request(URL, headers={'User-Agent': 'Mozilla/5.0'})
    
    try:
        remote_size = 0
        attempts = 3
        for i in range(attempts):
            try:
                with urllib.request.urlopen(req, context=ctx, timeout=15) as response:
                    headers = response.info()
                    size_header = headers.get('Content-Length')
                    if size_header:
                        remote_size = int(size_header)
                        if remote_size > 0:
                            break
                        else:
                            write_log(f"Intento {i+1}: El tamaño devuelto es 0 bytes (posible error del servidor).")
                    else:
                        write_log(f"Intento {i+1}: Content-Length faltante.")
            except Exception as e:
                write_log(f"Intento {i+1}: Fallo en la conexión ({e}).")
            
            if i < attempts - 1:
                import time
                time.sleep(5)

        if remote_size <= 0:
            write_log("Monitoreo abortado: No se pudo obtener un tamaño de archivo válido (> 0 bytes) tras los reintentos.")
            release_lock()
            return

        last_size = None
        if os.path.exists(SIZE_FILE):
            with open(SIZE_FILE, 'r') as f:
                content = f.read().strip()
                if content.isdigit():
                    last_size = int(content)

        if last_size is None:
            write_log(f"Inicializando. Guardando tamaño actual: {remote_size} bytes.")
            with open(SIZE_FILE, 'w') as f:
                f.write(str(remote_size))
            if not os.path.exists(CSV_PATH):
                write_log("Descargando archivo por primera vez...")
                download_file(ctx)
        elif remote_size > last_size:
            write_log(f"¡Incremento detectado! Tamaño anterior: {last_size} bytes. Nuevo tamaño: {remote_size} bytes.")
            
            write_log("Descargando el archivo actualizado...")
            download_file(ctx)
            
            import re
            month = "Desconocido"
            year = ""
            total_sales = "0"
            table = ""
            stdout_content = ""
            
            calc_command = [
                "python",
                os.path.join(BASE_DIR, "procesador_sri.py"),
                "--input-dir",
                BASE_DIR,
                "--output-file",
                os.path.join(BASE_DIR, "Procesador_SRI_Result")
            ]
            
            try:
                write_log("Ejecutando script de cálculo y consolidación...")
                result = subprocess.run(
                    calc_command,
                    capture_output=True,
                    text=True,
                    encoding='utf-8',
                    errors='ignore',
                    cwd=BASE_DIR
                )
                if result.returncode == 0:
                    stdout_content = result.stdout
                    write_log("Script de cálculo ejecutado con éxito.")
                    
                    m_month = re.search(r"MES DISPONIBLE:\s*([A-Za-z]+)\s+(\d{4})", stdout_content, re.IGNORECASE)
                    if m_month:
                        month = m_month.group(1).capitalize()
                        year = m_month.group(2)
                    
                    m_sales = re.search(r"Total de ventas del mes:\s*([\d,]+)", stdout_content)
                    if m_sales:
                        total_sales = m_sales.group(1)

                    m_table = re.search(r"(\| Pos \| Marca \|[\s\S]+?)(?===)", stdout_content)
                    if m_table:
                        table = m_table.group(1).strip()
                else:
                    write_log(f"Error al ejecutar el script de cálculo: {result.stderr}")
            except Exception as e:
                write_log(f"Excepción al ejecutar el script de cálculo: {e}")
            
            with open(SIZE_FILE, 'w') as f:
                f.write(str(remote_size))
            
            msg_text = f"El archivo SRI aumentó de {last_size/1024/1024:.2f}MB a {remote_size/1024/1024:.2f}MB. Resumen de {month} {year} enviado por correo."
            send_windows_notification("Incremento en archivo SRI", msg_text)
            send_email_notification(last_size, remote_size, month, year, total_sales, table)
            
        else:
            if remote_size < last_size:
                write_log(f"El tamaño remoto disminuyó ({remote_size} bytes) en comparación con el registrado ({last_size} bytes). No se descarga.")
                with open(SIZE_FILE, 'w') as f:
                    f.write(str(remote_size))
            else:
                write_log("El archivo no ha cambiado de tamaño desde la última verificación.")

    except Exception as e:
        write_log(f"Error durante la verificación: {e}")
    
    release_lock()

if __name__ == '__main__':
    main()
