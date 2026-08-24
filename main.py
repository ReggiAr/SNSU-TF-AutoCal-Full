import sys
import ctypes
from datetime import datetime
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QTextEdit,
                               QHBoxLayout, QGridLayout, QGroupBox, QFormLayout, QPlainTextEdit, QTableWidgetItem,
                               QLineEdit, QComboBox, QSpinBox, QPushButton, QProgressBar, 
                               QLabel, QTableWidget, QHeaderView, QTabWidget, QMessageBox, QCheckBox)
from PySide6.QtCore import Qt, QTimer, QThread, Signal
import openpyxl
import os
from pathlib import Path
from PySide6.QtGui import QPixmap, QIcon
import json
import pyvisa
import time
import winsound

def resource_path(relative_path):
    """Mendapatkan path absolut ke resource, berfungsi saat dijalankan sebagai .py maupun .exe"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# ==========================================
# WORKER QTHREAD UNTUK PENGUKURAN
# ==========================================
class WorkerPengukuranSemi(QThread):
    status_update = Signal(str)
    data_ready = Signal(dict)
    progress_update = Signal(int) 
    error_occurred = Signal(str)
    finished_task = Signal()

    def __init__(self, inst_counter, cmd_dict, fungsi, ch, osc, gate_mode, gate_val,
                 imp, coup, filter_, auto_trig, trig_lvl, slope, sens, pengulangan, titik_ukur, mode_manual=False):
        super().__init__()
        self.inst_counter = inst_counter
        self.cmd_dict = cmd_dict
        self.fungsi = fungsi
        self.ch = ch
        self.osc = osc
        self.gate_mode = gate_mode
        self.gate_val = gate_val
        
        self.imp = imp
        self.coup = coup
        self.filter_ = filter_
        self.auto_trig = auto_trig
        self.trig_lvl = trig_lvl
        self.slope = slope
        self.sens = sens
        
        self.pengulangan = pengulangan 
        self.titik_ukur = titik_ukur  
        self.mode_manual = mode_manual
        self._is_running = True

    def run(self):
        try:
            if not self.mode_manual:
                self.status_update.emit("Status: konfigurasi Instrument...")
            
            # --- PENGIRIMAN KONFIGURASI KE INSTRUMEN ---
                if "reset" in self.cmd_dict:
                    for r_cmd in self.cmd_dict["reset"].split(","):
                        if r_cmd.strip():
                            self.inst_counter.write(r_cmd.strip())
                time.sleep(0.5)
                
                #ini di modif biar bisa buat 53150a   
                func_data = self.cmd_dict.get("functions", {}).get(self.fungsi)
                if isinstance(func_data, dict):
                    cmd_fungsi = func_data.get(self.ch, "")
                    if cmd_fungsi:
                        self.inst_counter.write(cmd_fungsi.replace("{channel}", self.ch))
                elif isinstance(func_data, str) and func_data:
                    self.inst_counter.write(func_data.replace("{channel}", self.ch))
                
                cmd_osc = self.cmd_dict.get("set_oscillator", {}).get(self.osc, "")
                if cmd_osc:
                    self.inst_counter.write(cmd_osc)
                
                cmd_imp = self.cmd_dict.get("set_impedance", {}).get(self.imp, "")
                if cmd_imp:
                    self.inst_counter.write(cmd_imp.replace("{channel}", self.ch))
                
                cmd_coup = self.cmd_dict.get("set_coupling", {}).get(self.coup, "")
                if cmd_coup:
                    self.inst_counter.write(cmd_coup.replace("{channel}", self.ch))
                
                cmd_filter = self.cmd_dict.get("set_filter", {}).get(self.filter_, "")
                if cmd_filter:
                    self.inst_counter.write(cmd_filter.replace("{channel}", self.ch))
                
                cmd_auto = self.cmd_dict.get("set_auto_trigger", {}).get(self.auto_trig, "")
                if cmd_auto:
                    self.inst_counter.write(cmd_auto.replace("{channel}", self.ch))
                
                cmd_lvl = self.cmd_dict.get("set_trigger_level", "")
                if cmd_lvl and self.trig_lvl:
                    self.inst_counter.write(cmd_lvl.replace("{channel}", self.ch).replace("{value}", self.trig_lvl))
                
                cmd_slope = self.cmd_dict.get("set_slope", {}).get(self.slope, "")
                if cmd_slope:
                    self.inst_counter.write(cmd_slope.replace("{channel}", self.ch))
                
                sens_data = self.cmd_dict.get("set_sensitivity")
                if isinstance(sens_data, dict):
                    cmd_sens = sens_data.get(self.sens, "")
                    if cmd_sens:
                        self.inst_counter.write(cmd_sens.replace("{channel}", self.ch))
                elif isinstance(sens_data, str) and sens_data:
                    cmd_sens = sens_data.replace("{channel}", self.ch).replace("{value}", self.sens)
                    self.inst_counter.write(cmd_sens)

                cmd_gate = self.cmd_dict.get("set_gate_time", {}).get(self.gate_mode, "")
                if cmd_gate:
                    if self.gate_mode == "Time" and "{value}" in cmd_gate:
                        cmd_gate = cmd_gate.replace("{value}", self.gate_val)
                
                    perintah_gate = cmd_gate.split(",")
                    for cmd in perintah_gate:
                        if cmd.strip():
                            self.inst_counter.write(cmd.strip())
                            time.sleep(0.1) 
            else:
                self.status_update.emit("Status: Mode Manual. Melewati Konfigurasi")
                time.sleep(1)    
            # ---------------------------------------------
            self.status_update.emit("Status: Membaca data...")
            
            try:
                if self.mode_manual:
                    # Jika mode manual, beri batas waktu sangat lama (2000 detik/ 33 menit)
                    # Karena GUI tidak tahu berapa gate time fisik di Instrumen
                    self.inst_counter.timeout = None
                elif self.gate_mode == "Time" and self.gate_val:
                    gate_time_float = float(self.gate_val)
                    # Waktu time out = (Gate Time x 1000 ms) + Margin aman 300.000 ms (300 detik)
                    self.inst_counter.timeout = int((gate_time_float*1000) + 500000) 
                else:
                    self.inst_counter.timeout = 1000000
            except:
                self.inst_counter.timeout = 1000000
                
            # --- LOOPING PENGUKURAN ---
            for i in range(self.pengulangan):
                if not self._is_running:
                    break

                cmd_read = self.cmd_dict.get("read", "READ?")
                hasil_str = self.inst_counter.query(cmd_read).strip()
                    
                try:
                    hasil_float = float(hasil_str)
                except ValueError:
                    hasil_float = hasil_str 
                    
                waktu_ukur = datetime.now().strftime("%H:%M:%S")
                unit = "Hz" if "Freq" in self.fungsi else ("s" if "Period" in self.fungsi or "Time" in self.fungsi else "")
                
                if self._is_running:
                    # Menambahkan variabel index (i) untuk acuan kolom Excel
                    self.data_ready.emit({
                        "index": i,
                        "waktu": waktu_ukur, 
                        "titik_ukur": self.titik_ukur,
                        "hasil": hasil_float, 
                        "unit": unit
                    })
                    
                    progress = int(((i + 1) / self.pengulangan) * 100)
                    self.progress_update.emit(progress)
                
                time.sleep(0.5) 
            
            if self._is_running:
                self.status_update.emit("Status: Selesai")
            
        except Exception as e:
            self.error_occurred.emit(str(e))
            self.status_update.emit("Status: Error")
        finally:
            self.finished_task.emit()

    def stop(self):
        self._is_running = False

# ==========================================
# KELAS UTAMA APLIKASI
# ==========================================
class CalibrationApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SNSU TF-AutoCal | SNSU Time and Frequency Automated Calibration Software")
        self.resize(1100, 700) 

        self.setWindowIcon(QIcon(resource_path("logo.ico"))) 
        
        self.setStyleSheet("""
            QWidget { background-color: #f8f9fa; color: #2c3e50; font-family: 'Segoe UI', Arial; }
            QGroupBox {
                font-weight: bold; font-size: 13px;
                border: 1px solid #ced4da; border-radius: 6px;
                margin-top: 15px; background-color: #ffffff;
            }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 5px; color: #0f3460; }
            QComboBox, QLineEdit, QSpinBox {
                background-color: #ffffff; border: 1px solid #adb5bd;
                border-radius: 4px; padding: 4px; color: #000000; min-height: 20px;
            }
            QComboBox:focus, QLineEdit:focus, QSpinBox:focus { border: 2px solid #3498db; }
            QPushButton {
                background-color: #e9ecef; border: 1px solid #adb5bd;
                border-radius: 4px; padding: 8px; color: #000000; font-weight: bold;
            }
            QPushButton:hover { background-color: #dee2e6; }
            QPushButton#btnCreateFile { background-color: #f39c12; color: white; }
            QPushButton#btnCreateFile:hover { background-color: #e67e22; }
            QPushButton#btnStart { background-color: #27ae60; color: white; font-size: 14px; }
            QPushButton#btnStart:hover { background-color: #2ecc71; }
            QPushButton#btnStop { background-color: #c0392b; color: white; }
            QTableWidget { background-color: #ffffff; color: #000000; gridline-color: #dee2e6; }
            QHeaderView::section { background-color: #e9ecef; color: #000000; font-weight: bold; border: 1px solid #dee2e6; }
            QProgressBar { border: 1px solid #bbb; border-radius: 5px; text-align: center; font-weight: bold; background-color: #e0e0e0; height: 15px; }
            QProgressBar::chunk { background-color: #3498db; border-radius: 4px; }
        """)

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)

        # --- HEADER GLOBAL ---
        header_layout = QHBoxLayout()
        
        lbl_title = QLabel("SNSU Time and Frequency Automated Calibration Software \nNational Measurement Standards of Indonesia - Laboratory of Time and Frequency")
        lbl_title.setStyleSheet("font-size: 18px; font-weight: bold; color: #0f3460; background: transparent; padding-left: 10px;")
        
        self.lbl_clock = QLabel("Tanggal & Waktu")
        self.lbl_clock.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.lbl_clock.setStyleSheet("font-size: 16px; font-weight: bold; color: #2c3e50; background: transparent;")
        
        header_layout.addWidget(lbl_title)
        header_layout.addStretch()
        header_layout.addWidget(self.lbl_clock)
        main_layout.addLayout(header_layout)

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_waktu)
        self.timer.start(1000)
        self.update_waktu() 

        # --- TAB SYSTEM ---
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)

        self.data_instrumen = {} 
        self.inst_counter = None 
        self.file_excel_aktif = None
        self.worker_semi = None 
        self.current_excel_row = 2 
        
        try:
            self.rm = pyvisa.ResourceManager()
        except ValueError:
            try:
                self.rm = pyvisa.ResourceManager('@py')
            except Exception:
                self.rm = None
        except Exception as e:
            self.rm = None

        self.setup_tab_konfigurasi()
        self.setup_tab_semi_auto()
        self.setup_tab_full_auto()
        self.setup_tab_informasi()

        self.muat_instrumen_json()
        self.update_pilihan_instrumen() 
        # --- FOOTER ---
        lbl_footer = QLabel("SNSU Time and Frequency Automated Calibration Software | National Measurement Standards of Indonesia | © 2026")
        lbl_footer.setAlignment(Qt.AlignCenter)
        lbl_footer.setStyleSheet("font-size: 10px; font-weight: 500; color: #7f8c8d; margin-top: 8px; margin-bottom: 5px;")
        main_layout.addWidget(lbl_footer)

    # ==========================================
    # 1. SETUP TAB KONFIGURASI ALAT
    # ==========================================
    def setup_tab_konfigurasi(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        lbl_info = QLabel("Catatan: Pilihan di bawah ini ditarik otomatis dari file instrumen.json.")
        lbl_info.setStyleSheet("color: #27ae60; font-style: italic; font-weight: bold;")
        layout.addWidget(lbl_info)

        group_counter = QGroupBox("1. Pengaturan Frequency Counter")
        form_counter = QFormLayout(group_counter)
        
        self.cb_counter_name = QComboBox() 
        self.in_counter_address = QLineEdit()
        self.in_counter_address.setPlaceholderText("Misal: GPIB1::3::INSTR")
        
        lay_rs232_c = QHBoxLayout()
        self.cb_baud_c = QComboBox(); self.cb_baud_c.addItems(["9600", "19200", "38400", "115200"])
        self.cb_parity_c = QComboBox(); self.cb_parity_c.addItems(["None", "Even", "Odd"])
        self.cb_data_c = QComboBox(); self.cb_data_c.addItems(["8", "7"])
        lay_rs232_c.addWidget(QLabel("Baud Rate:")); lay_rs232_c.addWidget(self.cb_baud_c)
        lay_rs232_c.addWidget(QLabel("Parity:")); lay_rs232_c.addWidget(self.cb_parity_c)
        lay_rs232_c.addWidget(QLabel("Data Bits:")); lay_rs232_c.addWidget(self.cb_data_c)
        lay_rs232_c.addStretch()

        lay_btn_c = QHBoxLayout()
        self.btn_connect_c = QPushButton("🔌 Hubungkan & Uji Alat")
        self.btn_connect_c.setStyleSheet("background-color: #3498db; color: white;")
        self.btn_disconnect_c = QPushButton("❌ Putuskan")
        lay_btn_c.addWidget(self.btn_connect_c)
        lay_btn_c.addWidget(self.btn_disconnect_c)

        form_counter.addRow("Pilih Instrumen:", self.cb_counter_name)
        form_counter.addRow("Alamat (Address):", self.in_counter_address)
        form_counter.addRow("Setelan RS232:", lay_rs232_c)
        form_counter.addRow("", lay_btn_c)
        layout.addWidget(group_counter)

        group_source = QGroupBox("2. Pengaturan Frequency Source")
        vbox_source = QVBoxLayout(group_source)
        lbl_under_dev_source = QLabel("🚧\nUNDER DEVELOPMENT\nKoneksi Frequency Source sedang dalam pengembangan")
        lbl_under_dev_source.setAlignment(Qt.AlignCenter)
        lbl_under_dev_source.setStyleSheet("font-size: 16px; font-weight: bold; color: #e67e22; padding: 20px;")
        vbox_source.addStretch()
        vbox_source.addWidget(lbl_under_dev_source)
        vbox_source.addStretch()
        layout.addWidget(group_source)

        layout.addStretch()
        self.tabs.addTab(tab, "1. Konfigurasi Alat")

        self.btn_connect_c.clicked.connect(self.cek_koneksi_nyata)
        self.btn_disconnect_c.clicked.connect(self.putuskan_koneksi)

    # ==========================================
    # 2. SETUP TAB MODE SEMI OTOMATIS
    # ==========================================
    def setup_tab_semi_auto(self):
        tab = QWidget()
        main_layout = QVBoxLayout(tab)
        
        top_container = QWidget()
        top_vlayout = QVBoxLayout(top_container)
        top_vlayout.setContentsMargins(0, 0, 0, 0)
        top_vlayout.setSpacing(10)

        settings_container = QWidget()
        settings_grid = QGridLayout(settings_container)
        settings_grid.setContentsMargins(0, 0, 0, 0)

        # --- KOLOM 1: Measurement & Input ---
        group_meas = QGroupBox("1. Measurement & Input")
        form_meas = QFormLayout(group_meas)
        form_meas.setSpacing(2)
        
        self.cb_func = QComboBox(); self.cb_func.addItems(["Frequency","Ratio 1-2","Ratio 1-3","Ratio 2-1","Ratio 3-1","Totalize", "Phase 1-2", "Duty Cycle", "Volt Peaks","TI 1-2","Period","Rise Time","Fall Time","Pos Width", "Neg Width"])
        self.cb_channel = QComboBox(); self.cb_channel.addItems(["1", "2", "3"])
        self.cb_osc = QComboBox(); self.cb_osc.addItems(["EXT", "INT"])
        lay_ch_osc = QHBoxLayout()
        lay_ch_osc.addWidget(self.cb_channel); lay_ch_osc.addWidget(QLabel("Osc:")); lay_ch_osc.addWidget(self.cb_osc)
        
        self.cb_gate_mode = QComboBox()
        self.cb_gate_mode.addItems(["Time","Auto"])
        self.in_gate = QLineEdit("10"); self.in_gate.setPlaceholderText("s")
        lay_gate = QHBoxLayout()
        lay_gate.addWidget(self.cb_gate_mode)
        lay_gate.addWidget(self.in_gate)
        self.cb_gate_mode.currentIndexChanged.connect(self.update_gate_ui)
        
        self.cb_imp = QComboBox(); self.cb_imp.addItems(["50 OHM", "1 M OHM"])
        self.cb_coup = QComboBox(); self.cb_coup.addItems(["DC", "AC"])
        self.cb_filter = QComboBox(); self.cb_filter.addItems(["OFF", "ON"])
        lay_imp_coup = QHBoxLayout()
        lay_imp_coup.addWidget(self.cb_imp); lay_imp_coup.addWidget(self.cb_coup)

        form_meas.addRow("Func:", self.cb_func)
        form_meas.addRow("Ch/Osc:", lay_ch_osc)
        form_meas.addRow("Gate Time:", lay_gate)
        form_meas.addRow("Imp/Coup:", lay_imp_coup)
        form_meas.addRow("Filter:", self.cb_filter)
        settings_grid.addWidget(group_meas, 0, 0)

        # --- KOLOM 2: Trigger ---
        group_trig = QGroupBox("2. Trigger")
        form_trig = QFormLayout(group_trig)
        form_trig.setSpacing(2)
        self.cb_auto_trig = QComboBox(); self.cb_auto_trig.addItems(["OFF", "ON"])
        self.in_trig_lvl = QLineEdit("0")
        self.cb_slope = QComboBox(); self.cb_slope.addItems(["POS", "NEG"])
        self.cb_sens = QComboBox(); self.cb_sens.addItems(["MED","HI","LOW"])
        form_trig.addRow("Auto Trig:", self.cb_auto_trig)
        form_trig.addRow("Level (V):", self.in_trig_lvl)
        form_trig.addRow("Slope:", self.cb_slope)
        form_trig.addRow("Sensitivity:", self.cb_sens)
        settings_grid.addWidget(group_trig, 0, 1)

        # --- KOLOM 3: Excel & Environment ---
        group_excel = QGroupBox("3. File Excel & Lingkungan")
        form_excel = QFormLayout(group_excel)
        form_excel.setSpacing(2)
        self.in_order_no = QLineEdit(); self.in_order_no.setPlaceholderText("E-26-06-006")
        self.btn_create_excel = QPushButton("Buka/Buat Excel"); self.btn_create_excel.setObjectName("btnCreateFile")
        self.btn_create_excel.clicked.connect(self.buat_atau_buka_excel)
        self.in_temp_awal = QLineEdit("20.0"); self.in_temp_awal.setPlaceholderText("°C")
        self.in_rh_awal = QLineEdit("50.0"); self.in_rh_awal.setPlaceholderText("%")
        self.btn_env_awal = QPushButton("Catat Lingkungan Awal")
        self.btn_env_awal.clicked.connect(self.catat_lingkungan_ke_excel)
        
        form_excel.addRow("No Order:", self.in_order_no)
        form_excel.addRow("", self.btn_create_excel)
        form_excel.addRow("Suhu (°C):", self.in_temp_awal)
        form_excel.addRow("RH (%):", self.in_rh_awal)
        form_excel.addRow("", self.btn_env_awal)
        settings_grid.addWidget(group_excel, 0, 2)

        # --- KOLOM 4: Eksekusi ---
        group_exec = QGroupBox("4. Eksekusi Pengukuran")
        v_exec = QVBoxLayout(group_exec)

        # mode manual
        self.chk_manual = QCheckBox("Mode Manual Instrument")
        self.chk_manual.setStyleSheet("color: #d35400; font-weight: bold; margin-bottom: 10px;")
        self.chk_manual.toggled.connect(self.toggle_manual_mode)
        v_exec.addWidget(self.chk_manual)
        
        form_exec_inputs = QFormLayout()
        
        self.in_titik_ukur_val = QLineEdit("10")
        self.cb_titik_ukur_unit = QComboBox()
        self.cb_titik_ukur_unit.addItems(["Hz", "kHz", "MHz", "GHz", "s", "rpm"])
        
        lay_tu = QHBoxLayout()
        lay_tu.addWidget(self.in_titik_ukur_val)
        lay_tu.addWidget(self.cb_titik_ukur_unit)
        
        self.spin_pengulangan = QSpinBox()
        self.spin_pengulangan.setRange(1, 10000)
        self.spin_pengulangan.setValue(10)
        
        form_exec_inputs.addRow("Titik Ukur:", lay_tu)
        form_exec_inputs.addRow("Pengulangan:", self.spin_pengulangan)
        v_exec.addLayout(form_exec_inputs)

        self.btn_start_semi = QPushButton("▶ Mulai Pengukuran"); self.btn_start_semi.setObjectName("btnStart")
        self.btn_start_semi.clicked.connect(self.jalankan_pengukuran_semi_otomatis)
        self.btn_stop_semi = QPushButton("⏹ Berhenti"); self.btn_stop_semi.setObjectName("btnStop")
        self.btn_stop_semi.clicked.connect(self.hentikan_pengukuran)
        self.lbl_status = QLabel("Status: Siap")
        
        v_exec.addWidget(self.btn_start_semi)
        v_exec.addWidget(self.btn_stop_semi)
        v_exec.addStretch()
        self.chk_suara = QCheckBox("🔊 Notification")
        self.chk_suara.setChecked(True) # Dicentang secara default
        v_exec.addWidget(self.chk_suara)
        v_exec.addWidget(self.lbl_status)
        settings_grid.addWidget(group_exec, 0, 3)

        top_vlayout.addWidget(settings_container)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setAlignment(Qt.AlignCenter)
        top_vlayout.addWidget(self.progress_bar)
        
        main_layout.addWidget(top_container, stretch=1)

        # --- PANEL TABEL DATA ---
        self.table_data = QTableWidget(0, 5) 
        self.table_data.setHorizontalHeaderLabels(["No", "Waktu", "Titik Ukur", "Hasil Ukur", "Unit"])
        self.table_data.setColumnWidth(0, 40)
        self.table_data.setColumnWidth(1, 100)
        self.table_data.setColumnWidth(2, 100)
        self.table_data.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.table_data.setColumnWidth(4, 70)
        self.table_data.setAlternatingRowColors(True)
        main_layout.addWidget(self.table_data, stretch=4)
        
        self.tabs.addTab(tab, "2. Mode Semi-Otomatis")
        self.update_gate_ui()

    def update_gate_ui(self):
        if self.cb_gate_mode.currentText() == "Auto":
            self.in_gate.setEnabled(False)
        else:
            self.in_gate.setEnabled(True)

    def toggle_manual_mode(self, checked):
        # Membekukan/membuka inputan konfigurasi sesuai status checkbox
        self.cb_func.setEnabled(not checked)
        self.cb_channel.setEnabled(not checked)
        self.cb_osc.setEnabled(not checked)
        self.cb_gate_mode.setEnabled(not checked)
        self.in_gate.setEnabled(not checked)
        self.cb_imp.setEnabled(not checked)
        self.cb_coup.setEnabled(not checked)
        self.cb_filter.setEnabled(not checked)
        self.cb_auto_trig.setEnabled(not checked)
        self.in_trig_lvl.setEnabled(not checked)
        self.cb_slope.setEnabled(not checked)
        self.cb_sens.setEnabled(not checked)

    # ==========================================
    # 3. SETUP TAB MODE FULL OTOMATIS
    # ==========================================
    def setup_tab_full_auto(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        lbl_under_dev = QLabel("🛠️\nMODE FULL OTOMATIS\nSEDANG DALAM TAHAP PENGEMBANGAN\n🚧")
        lbl_under_dev.setAlignment(Qt.AlignCenter)
        lbl_under_dev.setStyleSheet("font-size: 24px; font-weight: bold; color: #c0392b;")
        layout.addStretch()
        layout.addWidget(lbl_under_dev)
        layout.addStretch()
        self.tabs.addTab(tab, "3. Mode Full Otomatis")

    # ==========================================
    # 4. SETUP TAB INFORMASI 
    # ==========================================
    def setup_tab_informasi(self):
        tab = QWidget()
        layout = QHBoxLayout(tab)
        
        # --- Kolom Kiri: Informasi & JSON Monitoring ---
        gb_info = QGroupBox("Informasi Aplikasi & Status Database")
        v_info = QVBoxLayout(gb_info)
        
        lbl_info = QLabel(
            "<b>SNSU TF-AutoCal</b><br>"
            "Time and Frequency Automated Calibration Software<br><br>"
            "Versi          : 3.0.0 (Build 2026)<br>"
            "Developer      : Reggi Aryunadi<br>"
            "Institusi      : Laboratorium SNSU Waktu dan Frekuensi <br><br>"
            "Software ini digunakan bersama Keysight Connection Expert.<br>"
        )
        lbl_info.setWordWrap(True)
        v_info.addWidget(lbl_info)

        # ==========================================
        # PENAMBAHAN TOMBOL USER GUIDE
        # ==========================================
        self.btn_user_guide = QPushButton("📖 Buka User Guide (Buku Panduan)")
        self.btn_user_guide.setStyleSheet("background-color: #3498db; color: white; padding: 6px; font-weight: bold; margin-top: 5px;")
        self.btn_user_guide.clicked.connect(self.buka_user_guide)
        v_info.addWidget(self.btn_user_guide)
        
        # UI Monitoring JSON Meter
        self.lbl_update_meter = QLabel("<b>Update JSON Meter:</b> -")
        self.lbl_update_meter.setStyleSheet("margin-top: 10px; color: #2980b9;")
        self.txt_meter_list = QTextEdit()
        self.txt_meter_list.setReadOnly(True)
        self.txt_meter_list.setMaximumHeight(80) # Dibatasi agar rapi
        
        # UI Monitoring JSON Sumber
        self.lbl_update_sumber = QLabel("<b>Update JSON Sumber:</b> -")
        self.lbl_update_sumber.setStyleSheet("margin-top: 10px; color: #d35400;")
        self.txt_sumber_list = QTextEdit()
        self.txt_sumber_list.setReadOnly(True)
        self.txt_sumber_list.setMaximumHeight(80)
        
        v_info.addWidget(self.lbl_update_meter)
        v_info.addWidget(self.txt_meter_list)
        v_info.addWidget(self.lbl_update_sumber)
        v_info.addWidget(self.txt_sumber_list)
        v_info.addStretch()

        # --- Kolom Kanan: Bantuan ---
        gb_help = QGroupBox("Laporan Kendala (Log Error)")
        form_help = QFormLayout(gb_help)
        self.in_help_nama = QLineEdit()
        self.in_help_desc = QTextEdit()
        self.btn_save_help = QPushButton("Simpan Laporan Kendala")
        self.btn_save_help.clicked.connect(self.simpan_laporan_kendala)

        form_help.addRow("Nama Operator:", self.in_help_nama)
        form_help.addRow("Deskripsi Kendala:", self.in_help_desc)
        form_help.addRow("", self.btn_save_help)

        layout.addWidget(gb_info)
        layout.addWidget(gb_help)
        self.tabs.addTab(tab, "4. Informasi & Bantuan")

    # ==========================================
    # FUNGSI BACKEND & LOGIKA JSON
    # ==========================================
    def muat_instrumen_json(self):
        folder_dokumen = Path.home() / "Documents" / "Data_Kalibrasi_BSN"
        if not folder_dokumen.exists():
            folder_dokumen.mkdir(parents=True, exist_ok=True)
            
        path_meter = folder_dokumen / "instrumen_meter.json"
        path_sumber = folder_dokumen / "instrumen_sumber.json"
        
        # 1. Muat JSON Meter
        try:
            if path_meter.exists():
                with open(path_meter, 'r') as f:
                    self.data_instrumen = json.load(f)
                
                # Update Info GUI
                tgl_update = self.data_instrumen.get("last_updated", "Tidak diketahui")
                self.lbl_update_meter.setText(f"<b>Update JSON Meter:</b> {tgl_update}")
                
                if "frequency_counters" in self.data_instrumen:
                    daftar = list(self.data_instrumen["frequency_counters"].keys())
                    self.txt_meter_list.setPlainText("\n".join(daftar))
            else:
                self.txt_meter_list.setPlainText("File instrumen_meter.json tidak ditemukan.")
                QMessageBox.warning(self, "Peringatan", f"File instrumen_meter.json hilang di:\n{path_meter}")
        except Exception as e:
            QMessageBox.critical(self, "Error JSON Meter", str(e))

        # 2. Muat JSON Sumber
        try:
            if path_sumber.exists():
                with open(path_sumber, 'r') as f:
                    self.data_sumber = json.load(f) # Simpan di variabel baru
                
                # Update Info GUI
                tgl_update = self.data_sumber.get("last_updated", "Tidak diketahui")
                self.lbl_update_sumber.setText(f"<b>Update JSON Sumber:</b> {tgl_update}")
                
                if "frequency_sources" in self.data_sumber:
                    daftar = list(self.data_sumber["frequency_sources"].keys())
                    self.txt_sumber_list.setPlainText("\n".join(daftar))
            else:
                self.txt_sumber_list.setPlainText("File instrumen_sumber.json tidak ditemukan.")
        except Exception as e:
            # Tidak perlu memblokir aplikasi jika sumber belum ada (sedang tahap pengembangan)
            pass

    def update_pilihan_instrumen(self):
        self.cb_counter_name.clear()
        if "frequency_counters" in self.data_instrumen:
            daftar_counter = list(self.data_instrumen["frequency_counters"].keys())
            self.cb_counter_name.addItems(daftar_counter)
            
        self.cb_counter_name.currentIndexChanged.connect(self.update_alamat_default)
        self.update_alamat_default()

    def update_alamat_default(self):
        nama_alat = self.cb_counter_name.currentText()
        if nama_alat in self.data_instrumen.get("frequency_counters", {}):
            alamat_default = self.data_instrumen["frequency_counters"][nama_alat].get("default_address", "")
            self.in_counter_address.setText(alamat_default)

    def update_waktu(self):
        sekarang = datetime.now()
        self.lbl_clock.setText(sekarang.strftime("%A, %d %B %Y | %H:%M:%S"))

    def cek_koneksi_nyata(self):
        if not self.rm:
            QMessageBox.warning(self, "Error", "Backend VISA tidak terdeteksi.")
            return
            
        alamat = self.in_counter_address.text().strip()
        if not alamat:
            QMessageBox.warning(self, "Peringatan", "Alamat instrumen tidak boleh kosong.")
            return
            
        try:
            self.inst_counter = self.rm.open_resource(alamat)
            if alamat.upper().startswith("ASRL") or alamat.upper().startswith("COM"):
                self.inst_counter.baud_rate = int(self.cb_baud_c.currentText())
                
            self.inst_counter.timeout = 2000
            respon_idn = self.inst_counter.query('*IDN?')
            QMessageBox.information(self, "Koneksi Berhasil!", f"Terhubung ke: {alamat}\nRespon: {respon_idn}")
            self.btn_connect_c.setText("✓ Terhubung")
            self.btn_connect_c.setStyleSheet("background-color: #27ae60; color: white;")
        except pyvisa.errors.VisaIOError as e:
            QMessageBox.critical(self, "Gagal Terhubung", f"Detail Error:\n{str(e)}")
            self.inst_counter = None
        except Exception as e:
            QMessageBox.critical(self, "Error Sistem", str(e))
            self.inst_counter = None

    def putuskan_koneksi(self):
        if self.inst_counter:
            try:
                self.inst_counter.close()
                self.inst_counter = None
                self.btn_connect_c.setText("🔌 Hubungkan & Uji Alat")
                self.btn_connect_c.setStyleSheet("background-color: #3498db; color: white;")
                QMessageBox.information(self, "Info", "Koneksi berhasil diputuskan.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Gagal memutuskan: {str(e)}")
        else:
            QMessageBox.information(self, "Info", "Memang tidak ada koneksi aktif.")
            
    def buat_atau_buka_excel(self):
        no_order = self.in_order_no.text().strip()
        if not no_order:
            QMessageBox.warning(self, "Peringatan", "Nomor Order tidak boleh kosong!\nSilakan isi dengan format E-TAHUN-BULAN-NO (Contoh: E-26-04-100).")
            return
            
        folder_dokumen = Path.home() / "Documents" / "Data_Kalibrasi_BSN"
        if not folder_dokumen.exists():
            folder_dokumen.mkdir(parents=True, exist_ok=True)
            
        nama_file = f"{no_order}_raw data.xlsx"
        path_file = folder_dokumen / nama_file
        self.file_excel_aktif = path_file
        
        try:
            if not path_file.exists():
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Data Pengukuran"
                
                # Setup format Header Horizontal
                ws.cell(row=1, column=2, value="Titik Ukur")
                ws.cell(row=1, column=3, value="Satuan")
                ws.cell(row=1, column=4, value="Hasil Pengukuran ->")
                
                ws_env = wb.create_sheet("Kondisi Lingkungan")
                ws_env.append(["Fase", "Waktu", "Suhu (°C)", "Kelembapan (%)"])
                wb.save(path_file)
                QMessageBox.information(self, "Berhasil", f"File Excel baru dibuat:\n{path_file}")
            else:
                QMessageBox.information(self, "Berhasil", f"File Excel ditemukan dan siap digunakan:\n{path_file}")
                
            self.btn_create_excel.setStyleSheet("background-color: #27ae60; color: white;")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Gagal membuat atau membuka file Excel:\n{e}")

    def catat_lingkungan_ke_excel(self):
        if not self.file_excel_aktif:
            QMessageBox.warning(self, "Peringatan", "File Excel belum dibuat/dibuka!")
            return
            
        # Tentukan fase berdasarkan teks yang sedang tampil di tombol
        fase = "AWAL" if "Awal" in self.btn_env_awal.text() else "AKHIR"
        
        try:
            wb = openpyxl.load_workbook(self.file_excel_aktif)
            if "Kondisi Lingkungan" not in wb.sheetnames:
                ws_env = wb.create_sheet("Kondisi Lingkungan")
                ws_env.append(["Fase", "Waktu", "Suhu (°C)", "Kelembapan (%)"])
            else:
                ws_env = wb["Kondisi Lingkungan"]
                
            waktu_sekarang = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            suhu = self.in_temp_awal.text()
            rh = self.in_rh_awal.text()
            
            ws_env.append([fase, waktu_sekarang, suhu, rh])
            wb.save(self.file_excel_aktif)
            QMessageBox.information(self, "Berhasil", f"Kondisi lingkungan ({fase}) berhasil dicatat!")
            
            # Ubah teks tombol menjadi sebaliknya untuk pencatatan berikutnya
            if fase == "AWAL":
                self.btn_env_awal.setText("Catat Lingkungan Akhir")
            else:
                self.btn_env_awal.setText("Catat Lingkungan Awal")
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Gagal mencatat kondisi lingkungan: {str(e)}")
    # ---------------------------------------------------------
    # EKSEKUSI PENGUKURAN DENGAN THREAD
    # ---------------------------------------------------------
    def jalankan_pengukuran_semi_otomatis(self):
        if not self.file_excel_aktif:
            QMessageBox.warning(self, "Peringatan", "Buka file Excel terlebih dahulu!")
            return
            
        if not self.inst_counter:
            QMessageBox.warning(self, "Peringatan", "Alat belum terhubung di tab Konfigurasi!")
            return
            
        nama_alat = self.cb_counter_name.currentText()
        if nama_alat not in self.data_instrumen.get("frequency_counters", {}):
            QMessageBox.warning(self, "Peringatan", "Perintah untuk alat ini tidak ditemukan di konfigurasi JSON.")
            return
            
        # --- PERBAIKAN: RESET TABEL & HITUNG BARIS EXCEL ---
        self.table_data.setRowCount(0)
        
        try:
            wb = openpyxl.load_workbook(self.file_excel_aktif)
            ws = wb["Data Pengukuran"]
            # Memastikan selalu mulai dari baris kosong (dimulai dari baris ke-2 jika data masih kosong)
            if ws.max_row < 2:
                self.current_excel_row = 2
            else:
                self.current_excel_row = ws.max_row + 1
        except Exception as e:
            self.current_excel_row = 2
        # ---------------------------------------------------

        cmd_dict = self.data_instrumen["frequency_counters"][nama_alat].get("commands", {})
        
        fungsi = self.cb_func.currentText()
        ch = self.cb_channel.currentText()
        osc = self.cb_osc.currentText()
        gate_mode = self.cb_gate_mode.currentText()
        gate_val = self.in_gate.text()
        
        imp = self.cb_imp.currentText()
        coup = self.cb_coup.currentText()
        filter_ = self.cb_filter.currentText()
        auto_trig = self.cb_auto_trig.currentText()
        trig_lvl = self.in_trig_lvl.text()
        slope = self.cb_slope.currentText()
        sens = self.cb_sens.currentText()
        
        pengulangan = self.spin_pengulangan.value()
        titik_ukur = f"{self.in_titik_ukur_val.text().strip()} {self.cb_titik_ukur_unit.currentText()}"
        
        self.btn_start_semi.setEnabled(False)
        self.progress_bar.setValue(0) 

        self.worker_semi = WorkerPengukuranSemi(
            inst_counter=self.inst_counter,
            cmd_dict=cmd_dict,
            fungsi=fungsi,
            ch=ch,
            osc=osc,
            gate_mode=gate_mode,
            gate_val=gate_val,
            imp=imp,
            coup=coup,
            filter_=filter_,
            auto_trig=auto_trig,
            trig_lvl=trig_lvl,
            slope=slope,
            sens=sens,
            pengulangan=pengulangan,
            titik_ukur=titik_ukur,
            mode_manual=self.chk_manual.isChecked()
        )
        
        self.worker_semi.status_update.connect(self.update_status_ui)
        self.worker_semi.progress_update.connect(self.progress_bar.setValue) 
        self.worker_semi.data_ready.connect(self.proses_data_masuk)
        self.worker_semi.error_occurred.connect(self.tampilkan_error_worker)
        self.worker_semi.finished_task.connect(self.worker_selesai)
        
        self.worker_semi.start()

    def update_status_ui(self, pesan):
        self.lbl_status.setText(pesan)

    def proses_data_masuk(self, data):
        waktu_ukur = data["waktu"]
        titik_ukur = data["titik_ukur"]
        hasil = data["hasil"]
        unit = data["unit"]
        idx = data["index"] # Indeks data ke-n

        # --- 1. CETAK KE GUI ---
        row_idx = self.table_data.rowCount()
        self.table_data.insertRow(row_idx)
        self.table_data.setItem(row_idx, 0, QTableWidgetItem(str(row_idx + 1)))
        self.table_data.setItem(row_idx, 1, QTableWidgetItem(waktu_ukur))
        self.table_data.setItem(row_idx, 2, QTableWidgetItem(titik_ukur))
        self.table_data.setItem(row_idx, 3, QTableWidgetItem(repr(hasil) if isinstance(hasil, float) else str(hasil)))
        self.table_data.setItem(row_idx, 4, QTableWidgetItem(unit))
        self.table_data.scrollToBottom()
        
        # --- 2. CETAK KE EXCEL (HORIZONTAL) ---
        try:
            wb = openpyxl.load_workbook(self.file_excel_aktif)
            ws = wb["Data Pengukuran"]
            
            excel_row = self.current_excel_row
            
            # Jika ini iterasi pertama (index = 0) dalam set pengukuran, masukkan angka dan satuan titik ukurnya di B dan C
            if idx == 0:
                val_str = self.in_titik_ukur_val.text().strip()
                try:
                    val_num = float(val_str)
                except ValueError:
                    val_num = val_str
                
                ws.cell(row=excel_row, column=2, value=val_num) 
                ws.cell(row=excel_row, column=3, value=self.cb_titik_ukur_unit.currentText()) 
            
            # Tulis hasil pengukuran secara menyamping mulai dari kolom ke-4 (D, E, F, ...) tanpa timestamp
            ws.cell(row=excel_row, column=4 + idx, value=hasil)
            
            wb.save(self.file_excel_aktif)
        except Exception as e:
            QMessageBox.warning(self, "Error Excel", f"Gagal menyimpan ke Excel: {str(e)}")

    def tampilkan_error_worker(self, error_msg):
        QMessageBox.critical(self, "Error Pengukuran", f"Gagal menjalankan pengukuran:\n{error_msg}")

    def worker_selesai(self):
        self.btn_start_semi.setEnabled(True)
        try:
            if self.chk_suara.isChecked():
                winsound.PlaySound(resource_path("notif.wav"), winsound.SND_FILENAME | winsound.SND_ASYNC)
        except Exception:
            pass
        # ------------------------------------

    def hentikan_pengukuran(self):
        if self.worker_semi and self.worker_semi.isRunning():
            self.worker_semi.stop()
            self.lbl_status.setText("Status: Dihentikan oleh pengguna.")
        else:
            self.lbl_status.setText("Status: Tidak ada pengukuran aktif.")

    # ---------------------------------------------------------
    def simpan_laporan_kendala(self):
        nama = self.in_help_nama.text().strip()
        kendala = self.in_help_desc.toPlainText().strip()
        
        if not nama or not kendala:
            QMessageBox.warning(self, "Peringatan", "Nama dan Deskripsi tidak boleh kosong.")
            return
            
        try:
            folder_dokumen = Path.home() / "Documents" / "Data_Kalibrasi_BSN" / "Log_Kendala"
            folder_dokumen.mkdir(parents=True, exist_ok=True)
            
            nama_file = folder_dokumen / f"ErrorLog_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            
            with open(nama_file, 'w') as f:
                f.write(f"LAPORAN KENDALA APLIKASI KALIBRASI\n")
                f.write(f"================================\n")
                f.write(f"Waktu Laporan : {datetime.now().strftime('%d %B %Y, %H:%M:%S')}\n")
                f.write(f"Nama Petugas  : {nama}\n")
                f.write(f"--------------------------------\n")
                f.write(f"Deskripsi Kendala:\n{kendala}\n")
                f.write(f"--------------------------------\n")
                f.write(f"System Info: Windows OS / PySide6 Framework\n")
                
            QMessageBox.information(self, "Berhasil", f"Laporan kendala telah disimpan di:\n{nama_file}\n\nTerima kasih, developer akan segera meninjaunya.")
            self.in_help_nama.clear()
            self.in_help_desc.clear()
        except Exception as e:
            QMessageBox.critical(self, "Gagal Simpan", f"Gagal menyimpan file log:\n{str(e)}")


    def buka_user_guide(self):
        try:
            # Menggunakan resource_path agar bisa melacak PDF di dalam perut .exe
            pdf_path = resource_path("I.MF.0.70 User Guide for SNSU TF-AutoCal.pdf")
            
            if os.path.exists(pdf_path):
                # Membuka file menggunakan aplikasi PDF bawaan Windows (Edge/Acrobat)
                os.startfile(pdf_path)
            else:
                QMessageBox.warning(self, "Tidak Ditemukan", "File User Guide tidak ditemukan di dalam sistem.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Gagal membuka User Guide:\n{str(e)}")
        # ---------------------------------------------------------

if __name__ == "__main__":
    try:
        myappid = 'laboratorium.waktudanfrekuensi.kalibrasi.2'
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except Exception:
        pass
        
    app = QApplication(sys.argv)

    # Ikon taskbar (dibungkus resource_path)
    app.setWindowIcon(QIcon(resource_path("logo.ico")))

    window = CalibrationApp()
    window.show()
    sys.exit(app.exec())