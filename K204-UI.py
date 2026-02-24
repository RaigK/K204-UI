import serial
import time
import struct
import json
import os
import threading
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog

# Externe Bibliotheken
from serial.tools import list_ports
from openpyxl import Workbook

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

CONFIG_FILE = 'k204_config.json'

# --- Lokalisierung ---
TEXTS = {
    "de": {
        "title": "VOLTCRAFT K204 Professional Logger",
        "conn_frame": " Verbindung ",
        "set_frame": " Einstellungen ",
        "chan_frame": " Kanäle ",
        "port": "COM Port:",
        "prefix": "Datei-Prefix:",
        "suffix": "Dateiendung:",
        "path": "Speicherpfad:",
        "cycles": "Zyklen (0=∞):",
        "interval": "Intervall (s):",
        "start": "START",
        "stop": "STOP",
        "refresh": "↻",
        "browse": "Durchsuchen",
        "suffix_time": "Zeitstempel",
        "suffix_num": "Fortl. Nummer",
        "err_port": "Bitte COM-Port wählen!",
        "err_path": "Bitte gültigen Speicherpfad wählen!",
        "log_start": "Logger gestartet an {port}",
        "log_stop": "Messung gestoppt.",
        "log_err": "Fehler: {msg}",
        "log_data_err": "Fehler: Ungültige Daten.",
        "plot_title": "Echtzeit Temperaturverlauf",
        "temp": "Temperatur"
    },
    "en": {
        "title": "VOLTCRAFT K204 Professional Logger",
        "conn_frame": " Connection ",
        "set_frame": " Settings ",
        "chan_frame": " Channels ",
        "port": "COM Port:",
        "prefix": "File Prefix:",
        "suffix": "File Suffix:",
        "path": "Save Path:",
        "cycles": "Cycles (0=∞):",
        "interval": "Interval (s):",
        "start": "START",
        "stop": "STOP",
        "refresh": "↻",
        "browse": "Browse",
        "suffix_time": "Timestamp",
        "suffix_num": "Sequential No.",
        "err_port": "Please select COM port!",
        "err_path": "Please select valid save path!",
        "log_start": "Logger started on {port}",
        "log_stop": "Measurement stopped.",
        "log_err": "Error: {msg}",
        "log_data_err": "Error: Invalid data.",
        "plot_title": "Real-time Temperature Trace",
        "temp": "Temperature"
    }
}

class K204App:
    def __init__(self, root):
        self.root = root
        self.running = False
        self.config = self.load_config()
        self.lang = self.config["settings"].get("language", "de")
        
        self.root.title(TEXTS[self.lang]["title"])
        self.root.geometry("1150x800")
        
        # Daten für Plot
        self.x_data = []
        self.y_data = {f"T{i}": [] for i in range(1, 5)}
        
        self.setup_ui()
        self.refresh_ports()
        
        # Zuletzt genutzten Port auswählen
        last_port = self.config["settings"].get("last_port")
        if last_port in self.combo_port['values']:
            self.combo_port.set(last_port)

    def load_config(self):
        defaults = {
            "channels": {"T1": "Kanal 1", "T2": "Kanal 2", "T3": "Kanal 3", "T4": "Kanal 4"},
            "settings": {
                "cycles": 0, 
                "prefix": "messung", 
                "interval": 1.0, 
                "language": "de",
                "save_path": os.getcwd(),
                "suffix_type": "Zeitstempel",
                "last_port": ""
            }
        }
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    if "channels" in data: defaults["channels"].update(data["channels"])
                    if "settings" in data: defaults["settings"].update(data["settings"])
            except: pass
        return defaults

    def save_config(self):
        self.config["settings"]["prefix"] = self.ent_prefix.get()
        self.config["settings"]["language"] = self.lang
        self.config["settings"]["save_path"] = self.ent_path.get()
        self.config["settings"]["suffix_type"] = self.combo_suffix.get()
        self.config["settings"]["last_port"] = self.combo_port.get()
        try:
            self.config["settings"]["cycles"] = int(self.ent_cycles.get())
            self.config["settings"]["interval"] = float(self.ent_interval.get())
        except ValueError: pass
        
        for i in range(1, 5):
            self.config["channels"][f"T{i}"] = self.chan_entries[i-1].get()
        
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, indent=4)

    def change_language(self, event=None):
        self.lang = "de" if self.combo_lang.get() == "Deutsch" else "en"
        self.save_config()
        self.root.title(TEXTS[self.lang]["title"])
        self.update_ui_texts()

    def update_ui_texts(self):
        t = TEXTS[self.lang]
        self.lbl_conn.config(text=t["conn_frame"])
        self.lbl_set.config(text=t["set_frame"])
        self.lbl_chan.config(text=t["chan_frame"])
        self.lbl_port_text.config(text=t["port"])
        self.lbl_prefix.config(text=t["prefix"])
        self.lbl_suffix.config(text=t["suffix"])
        self.lbl_path.config(text=t["path"])
        self.lbl_cycles.config(text=t["cycles"])
        self.lbl_interval.config(text=t["interval"])
        self.btn_start.config(text=t["start"])
        self.btn_stop.config(text=t["stop"])
        self.btn_browse.config(text=t["browse"])
        
        curr_s = self.combo_suffix.get()
        s_opts = [t["suffix_time"], t["suffix_num"]]
        self.combo_suffix['values'] = s_opts
        if curr_s not in s_opts: self.combo_suffix.current(0)
        
        self.ax.set_title(t["plot_title"])
        self.ax.set_ylabel(t["temp"])
        self.canvas.draw_idle()

    def setup_ui(self):
        header = ttk.Frame(self.root, padding=5)
        header.pack(fill="x")
        self.combo_lang = ttk.Combobox(header, values=["Deutsch", "English"], state="readonly", width=10)
        self.combo_lang.set("Deutsch" if self.lang == "de" else "English")
        self.combo_lang.pack(side="right")
        self.combo_lang.bind("<<ComboboxSelected>>", self.change_language)

        main_paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_paned.pack(fill="both", expand=True)

        left_frame = ttk.Frame(main_paned, padding=10)
        main_paned.add(left_frame, weight=1)

        # Verbindung
        self.lbl_conn = ttk.LabelFrame(left_frame, text=" Verbindung ", padding=5)
        self.lbl_conn.pack(fill="x", pady=5)
        self.lbl_port_text = ttk.Label(self.lbl_conn, text="Port:")
        self.lbl_port_text.pack(side="left")
        self.combo_port = ttk.Combobox(self.lbl_conn)
        self.combo_port.pack(side="left", fill="x", expand=True, padx=2)
        ttk.Button(self.lbl_conn, text="↻", width=3, command=self.refresh_ports).pack(side="left")

        # Einstellungen
        self.lbl_set = ttk.LabelFrame(left_frame, text=" Einstellungen ", padding=5)
        self.lbl_set.pack(fill="x", pady=5)
        
        self.lbl_prefix = ttk.Label(self.lbl_set, text="Prefix:")
        self.lbl_prefix.grid(row=0, column=0, sticky="w")
        self.ent_prefix = ttk.Entry(self.lbl_set)
        self.ent_prefix.insert(0, self.config["settings"]["prefix"])
        self.ent_prefix.grid(row=0, column=1, sticky="ew", padx=5)

        self.lbl_suffix = ttk.Label(self.lbl_set, text="Suffix:")
        self.lbl_suffix.grid(row=1, column=0, sticky="w")
        self.combo_suffix = ttk.Combobox(self.lbl_set, state="readonly")
        self.combo_suffix.grid(row=1, column=1, sticky="ew", padx=5, pady=2)

        self.lbl_path = ttk.Label(self.lbl_set, text="Pfad:")
        self.lbl_path.grid(row=2, column=0, sticky="w")
        path_sub = ttk.Frame(self.lbl_set)
        path_sub.grid(row=2, column=1, sticky="ew", padx=5)
        self.ent_path = ttk.Entry(path_sub)
        self.ent_path.insert(0, self.config["settings"]["save_path"])
        self.ent_path.pack(side="left", fill="x", expand=True)
        self.btn_browse = ttk.Button(path_sub, text="...", width=3, command=self.browse_path)
        self.btn_browse.pack(side="right")

        self.lbl_cycles = ttk.Label(self.lbl_set, text="Zyklen:")
        self.lbl_cycles.grid(row=3, column=0, sticky="w")
        self.ent_cycles = ttk.Entry(self.lbl_set)
        self.ent_cycles.insert(0, str(self.config["settings"]["cycles"]))
        self.ent_cycles.grid(row=3, column=1, sticky="ew", padx=5)

        self.lbl_interval = ttk.Label(self.lbl_set, text="Intervall:")
        self.lbl_interval.grid(row=4, column=0, sticky="w")
        self.ent_interval = ttk.Entry(self.lbl_set)
        self.ent_interval.insert(0, str(self.config["settings"]["interval"]))
        self.ent_interval.grid(row=4, column=1, sticky="ew", padx=5)
        self.lbl_set.columnconfigure(1, weight=1)

        # Kanäle
        self.lbl_chan = ttk.LabelFrame(left_frame, text=" Kanäle ", padding=5)
        self.lbl_chan.pack(fill="x", pady=5)
        self.chan_entries = []
        for i in range(1, 5):
            ttk.Label(self.lbl_chan, text=f"T{i}:").grid(row=i-1, column=0)
            en = ttk.Entry(self.lbl_chan)
            en.insert(0, self.config["channels"][f"T{i}"])
            en.grid(row=i-1, column=1, sticky="ew", padx=5, pady=2)
            self.chan_entries.append(en)
        self.lbl_chan.columnconfigure(1, weight=1)

        # Buttons
        btn_frame = ttk.Frame(left_frame)
        btn_frame.pack(fill="x", pady=10)
        self.btn_start = ttk.Button(btn_frame, text="START", command=self.start_measurement)
        self.btn_start.pack(side="left", expand=True, fill="x", padx=2)
        self.btn_stop = ttk.Button(btn_frame, text="STOP", state="disabled", command=self.stop_measurement)
        self.btn_stop.pack(side="left", expand=True, fill="x", padx=2)

        self.log_area = scrolledtext.ScrolledText(left_frame, height=12, font=("Consolas", 8))
        self.log_area.pack(fill="both", expand=True)

        # Plot
        right_frame = ttk.Frame(main_paned, padding=10)
        main_paned.add(right_frame, weight=3)
        self.fig = Figure(figsize=(6, 4), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.ax.grid(True, linestyle='--', alpha=0.7)
        self.plot_colors = ['#e74c3c', '#3498db', '#2ecc71', '#f39c12']
        self.lines = []
        for i in range(4):
            line, = self.ax.plot([], [], color=self.plot_colors[i], label=f"T{i+1}")
            self.lines.append(line)
        self.canvas = FigureCanvasTkAgg(self.fig, master=right_frame)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

        self.update_ui_texts()
        s_val = self.config["settings"].get("suffix_type", "Zeitstempel")
        if s_val in self.combo_suffix['values']: self.combo_suffix.set(s_val)
        else: self.combo_suffix.current(0)

    def browse_path(self):
        p = filedialog.askdirectory(initialdir=self.ent_path.get())
        if p:
            self.ent_path.delete(0, tk.END)
            self.ent_path.insert(0, p)

    def refresh_ports(self):
        current_val = self.combo_port.get()
        ports = [p.device for p in list_ports.comports()]
        self.combo_port['values'] = ports
        if current_val in ports:
            self.combo_port.set(current_val)
        elif ports:
            self.combo_port.current(0)

    def log(self, message):
        self.log_area.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} {message}\n")
        self.log_area.see(tk.END)

    def get_next_filename(self, path, prefix, s_type):
        if s_type in [TEXTS["de"]["suffix_time"], TEXTS["en"]["suffix_time"]]:
            return f"{prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        else:
            idx = 1
            while True:
                fname = f"{prefix}_{idx:03d}.xlsx"
                if not os.path.exists(os.path.join(path, fname)):
                    return fname
                idx += 1

    def start_measurement(self):
        t = TEXTS[self.lang]
        if not self.combo_port.get():
            messagebox.showerror("Error", t["err_port"])
            return
        if not os.path.isdir(self.ent_path.get()):
             messagebox.showerror("Error", t["err_path"])
             return
        
        self.save_config()
        self.running = True
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        
        self.x_data = []
        for k in self.y_data: self.y_data[k] = []
        for i, line in enumerate(self.lines):
            line.set_label(f"T{i+1}: {self.chan_entries[i].get()}")
        self.ax.legend(loc="upper left", fontsize='small')

        self.thread = threading.Thread(target=self.measurement_worker, daemon=True)
        self.thread.start()

    def stop_measurement(self):
        self.running = False
        self.btn_start.config(state="normal")
        self.btn_stop.config(state="disabled")

    def measurement_worker(self):
        t = TEXTS[self.lang]
        port = self.combo_port.get()
        prefix = self.ent_prefix.get()
        path = self.ent_path.get()
        s_type = self.combo_suffix.get()
        cycles = int(self.ent_cycles.get() or 0)
        interval = float(self.ent_interval.get() or 1.0)
        
        fname = self.get_next_filename(path, prefix, s_type)
        full_path = os.path.join(path, fname)
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        ws.append(["Time", "Runtime", "Seconds", "T1", "T2", "T3", "T4", "Unit"])
        
        start_time = datetime.now()
        count = 0
        
        try:
            with serial.Serial(port, 9600, timeout=2) as ser:
                self.log(t["log_start"].format(port=port))
                
                while self.running:
                    if cycles > 0 and count >= cycles: break
                    count += 1
                    ser.reset_input_buffer()
                    ser.write(b'\x41')
                    time.sleep(0.4)
                    raw = ser.read(45)
                    
                    if len(raw) == 45 and raw[0] == 0x02:
                        temps_raw = struct.unpack('>hhhh', raw[7:15])
                        ol_byte, res_byte = raw[39], raw[43]
                        unit = '°C' if (raw[1] & 0x80) else '°F'
                        
                        current_vals = []
                        now = datetime.now()
                        elapsed_total = (now - start_time).total_seconds()
                        self.x_data.append(elapsed_total)

                        for i in range(4):
                            is_ol = bool(ol_byte & (1 << i))
                            # Korrektur Dezimalstelle: Falls Bit 0 -> x10 (Divisor 10)
                            # Meist ist Bit auf 0 bei 0.1 Auflösung im K204 Protokoll
                            divisor = 1.0 if bool(res_byte & (1 << i)) else 10.0
                            val = None if is_ol else temps_raw[i]/divisor
                            current_vals.append(val)
                            self.y_data[f"T{i+1}"].append(val)
                            
                        ws.append([now.strftime("%H:%M:%S"), str(now-start_time).split('.')[0], round(elapsed_total,1)] + current_vals + [unit])
                        wb.save(full_path)
                        self.root.after(0, self.update_ui_elements, count, current_vals, elapsed_total)
                    else:
                        self.log(t["log_data_err"])
                    
                    time.sleep(max(0.1, interval - 0.5))
                    
        except Exception as e:
            self.log(t["log_err"].format(msg=str(e)))
        finally:
            self.root.after(0, self.stop_measurement)

    def update_ui_elements(self, count, vals, elapsed):
        val_str = " | ".join([f"{v:.1f}" if isinstance(v, float) else "OL" for v in vals])
        self.log(f"#{count} | {val_str}")
        for i, key in enumerate(self.y_data):
            self.lines[i].set_data(self.x_data, self.y_data[key])
        self.ax.relim()
        self.ax.autoscale_view()
        self.canvas.draw_idle()

if __name__ == "__main__":
    root = tk.Tk()
    app = K204App(root)
    root.mainloop()