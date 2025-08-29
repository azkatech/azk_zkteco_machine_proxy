import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText
import sqlite3
import os
from datetime import datetime, timedelta
import xmlrpc.client
import json
import threading
import time
import webbrowser

# --- IMPORTANT ---
# This script requires pyzk and pytz libraries.
# Please install them using pip:
# pip install pyzk pytz
from zk import ZK
import pytz


# --- Configuration ---
DB_FILE = 'local_zkteco_proxy.db'
VERSION_NUM = '1.0.2'

# --- Database Functions ---
def init_db():
    """Initializes the database and creates/alters tables if they don't exist."""
    with sqlite3.connect(DB_FILE, check_same_thread=False) as conn:
        cursor = conn.cursor()
        # Enable foreign key support
        cursor.execute("PRAGMA foreign_keys = ON")

        # Create zkteco_machines table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS zkteco_machines (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT,
                ip TEXT NOT NULL,
                port TEXT NOT NULL,
                password TEXT,
                serial_number TEXT,
                last_connected TEXT,
                odoo_machine_name TEXT,
                odoo_machine_id INTEGER,
                machine_timezone TEXT
            )
        ''')
        
        # Add machine_timezone column for backward compatibility
        try:
            cursor.execute("SELECT machine_timezone FROM zkteco_machines LIMIT 1")
        except sqlite3.OperationalError:
            cursor.execute("ALTER TABLE zkteco_machines ADD COLUMN machine_timezone TEXT")
        
        # Create odoo config table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS odoo_config (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        ''')
        
        # Create settings table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        ''')

        # Create users table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                connection_id INTEGER NOT NULL,
                uid INTEGER NOT NULL,
                user_id TEXT NOT NULL,
                name TEXT,
                synched_time DATETIME,
                FOREIGN KEY (connection_id) REFERENCES zkteco_machines (id) ON DELETE CASCADE,
                UNIQUE(connection_id, user_id) 
            )
        ''')

        # Create attendance table
        cursor.execute('''
             CREATE TABLE IF NOT EXISTS attendance (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                connection_id INTEGER NOT NULL,
                user_id TEXT NOT NULL,
                att_id TEXT NOT NULL,
                timestamp DATETIME NOT NULL,
                synched_time DATETIME,
                FOREIGN KEY (connection_id) REFERENCES zkteco_machines (id) ON DELETE CASCADE,
                UNIQUE(connection_id, user_id, timestamp)
            )
        ''')

        # Add synched_time columns for backward compatibility
        try:
            cursor.execute("SELECT synched_time FROM users LIMIT 1")
        except sqlite3.OperationalError:
            cursor.execute("ALTER TABLE users ADD COLUMN synched_time DATETIME")
        try:
            cursor.execute("SELECT synched_time FROM attendance LIMIT 1")
        except sqlite3.OperationalError:
            cursor.execute("ALTER TABLE attendance ADD COLUMN synched_time DATETIME")
        try:
            cursor.execute("SELECT att_id FROM attendance LIMIT 1")
        except sqlite3.OperationalError:
            cursor.execute("ALTER TABLE attendance ADD COLUMN att_id TEXT")

        # Create logs table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                connection_id INTEGER,
                timestamp DATETIME NOT NULL,
                operation TEXT NOT NULL,
                message TEXT,
                FOREIGN KEY (connection_id) REFERENCES zkteco_machines (id) ON DELETE CASCADE
            )
        ''')
        conn.commit()

def db_execute(query, params=(), fetch=None):
    """A helper function to execute database queries."""
    with sqlite3.connect(DB_FILE, check_same_thread=False) as conn:
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute("PRAGMA foreign_keys = ON")
        cursor.execute(query, params)
        if fetch == 'one':
            result = cursor.fetchone()
        elif fetch == 'all':
            result = cursor.fetchall()
        else:
            conn.commit()
            result = None
        return result

# --- Main Application Class ---
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Azkatech's ZKTeco Machine proxy")
        self.geometry("950x750")

        # Initialize DB
        init_db()
        
        self.editing_connection_id = None

        # --- Scheduler Setup ---
        self.scheduler_thread = None
        self.scheduler_running = threading.Event()

        # --- UI Variables ---
        self.users_to_sync_var = tk.BooleanVar()
        self.attendance_to_sync_var = tk.BooleanVar()

        # --- UI Setup ---
        self.main_notebook = ttk.Notebook(self)
        self.main_notebook.pack(pady=10, padx=10, expand=True, fill="both")
        
        # --- Create Main Frames/Tabs ---
        self.machines_manager_frame = ttk.Frame(self.main_notebook, padding="10")
        self.data_frame = ttk.Frame(self.main_notebook)
        self.config_frame = ttk.Frame(self.main_notebook)
        self.scheduler_frame = ttk.Frame(self.main_notebook, padding="10")
        
        self.main_notebook.add(self.machines_manager_frame, text='Machines Manager')
        self.main_notebook.add(self.data_frame, text='Data')
        self.main_notebook.add(self.config_frame, text='Configuration')
        self.main_notebook.add(self.scheduler_frame, text='Scheduler')
        
        # --- Create Nested Notebook for Data ---
        self.data_notebook = ttk.Notebook(self.data_frame)
        self.data_notebook.pack(expand=True, fill="both", padx=5, pady=5)
        self.users_frame = ttk.Frame(self.data_notebook, padding="10")
        self.attendance_frame = ttk.Frame(self.data_notebook, padding="10")
        self.logs_frame = ttk.Frame(self.data_notebook, padding="10")
        self.data_notebook.add(self.users_frame, text='Device Users')
        self.data_notebook.add(self.attendance_frame, text='Attendance Logs')
        self.data_notebook.add(self.logs_frame, text='Logs')
        
        # --- Create Nested Notebook for Configuration ---
        self.config_notebook = ttk.Notebook(self.config_frame)
        self.config_notebook.pack(expand=True, fill="both", padx=5, pady=5)
        self.odoo_frame = ttk.Frame(self.config_notebook, padding="10")
        self.settings_frame = ttk.Frame(self.config_notebook, padding="10")
        self.config_notebook.add(self.odoo_frame, text='Odoo Connection')
        self.config_notebook.add(self.settings_frame, text='Settings')

        # Create all UI elements FIRST
        self.create_connections_tab()
        self.create_users_tab()
        self.create_attendance_tab()
        self.create_scheduler_tab()
        self.create_logs_tab()
        self.create_settings_tab()
        self.create_odoo_tab()
        
        # NOW, load data and populate the UI
        self.refresh_all_data()

        # Handle window closing
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Start the scheduler automatically on launch
        self.start_scheduler()

    def on_closing(self):
        """Handle cleanup when the application window is closed."""
        if self.scheduler_running.is_set():
            self.stop_scheduler()
        self.destroy()

    def refresh_all_data(self):
        """Loads all data from DB and updates all tables."""
        self.load_settings_from_db()
        self.load_connections_from_db()
        self.load_odoo_details_from_db()
        
        self.update_connections_table()
        self.update_users_table()
        self.update_attendance_table()
        self.update_logs_table()

    def create_connections_tab(self):
        self.add_edit_frame = ttk.LabelFrame(self.machines_manager_frame, text="Add new Machine", padding="10")
        self.add_edit_frame.pack(fill="x", expand="yes", padx=5, pady=5)

        ttk.Label(self.add_edit_frame, text="Name:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.name_entry = ttk.Entry(self.add_edit_frame)
        self.name_entry.grid(row=0, column=1, columnspan=3, padx=5, pady=5, sticky="ew")

        ttk.Label(self.add_edit_frame, text="IP Address:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.ip_entry = ttk.Entry(self.add_edit_frame)
        self.ip_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(self.add_edit_frame, text="Port:").grid(row=1, column=2, padx=5, pady=5, sticky="w")
        self.port_entry = ttk.Entry(self.add_edit_frame, width=10)
        self.port_entry.grid(row=1, column=3, padx=5, pady=5, sticky="ew")
        
        ttk.Label(self.add_edit_frame, text="Password:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.pass_entry = ttk.Entry(self.add_edit_frame, show="*")
        self.pass_entry.grid(row=2, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(self.add_edit_frame, text="Odoo Machine Name:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.odoo_name_entry = ttk.Entry(self.add_edit_frame)
        self.odoo_name_entry.grid(row=3, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(self.add_edit_frame, text="Odoo Machine ID:").grid(row=3, column=2, padx=5, pady=5, sticky="w")
        self.odoo_id_entry = ttk.Entry(self.add_edit_frame)
        self.odoo_id_entry.grid(row=3, column=3, padx=5, pady=5, sticky="ew")

        ttk.Label(self.add_edit_frame, text="Machine Timezone:").grid(row=4, column=0, padx=5, pady=5, sticky="w")
        self.timezone_entry = ttk.Entry(self.add_edit_frame, state="readonly")
        self.timezone_entry.grid(row=4, column=1, padx=5, pady=5, sticky="ew")

        self.add_edit_frame.columnconfigure(1, weight=1)
        self.clear_connection_entries() # Set defaults
        
        self.action_button_frame = ttk.Frame(self.add_edit_frame)
        self.action_button_frame.grid(row=5, column=0, columnspan=4, pady=10)
        
        self.add_button = ttk.Button(self.action_button_frame, text="Add Machine", command=self.add_connection)
        self.save_button = ttk.Button(self.action_button_frame, text="Save Changes", command=self.save_connection_changes)
        self.cancel_button = ttk.Button(self.action_button_frame, text="Cancel", command=self.cancel_edit)
        self.add_button.pack()

        table_frame = ttk.Frame(self.machines_manager_frame)
        table_frame.pack(fill="both", expand=True, padx=5, pady=5)

        columns = ('name', 'ip', 'port', 'odoo_name', 'odoo_id', 'timezone', 'serial', 'last_conn')
        self.connections_table = ttk.Treeview(table_frame, columns=columns, show='headings')
        self.connections_table.heading('name', text='Name')
        self.connections_table.heading('ip', text='IP Address')
        self.connections_table.heading('port', text='Port')
        self.connections_table.heading('odoo_name', text='Odoo Name')
        self.connections_table.heading('odoo_id', text='Odoo ID')
        self.connections_table.heading('timezone', text='Timezone')
        self.connections_table.heading('serial', text='Serial Number')
        self.connections_table.heading('last_conn', text='Last Connected')
        self.connections_table.column('name', width=120)
        self.connections_table.column('ip', width=100)
        self.connections_table.column('port', width=50)
        self.connections_table.column('odoo_name', width=120)
        self.connections_table.column('odoo_id', width=60)
        self.connections_table.column('timezone', width=120)
        self.connections_table.column('serial', width=120)
        self.connections_table.column('last_conn', width=130)


        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.connections_table.yview)
        scrollbar.pack(side="right", fill="y")
        self.connections_table.configure(yscrollcommand=scrollbar.set)
        self.connections_table.pack(side="left", fill="both", expand=True)

        # --- Button Groups and Status Bar ---
        bottom_frame = ttk.Frame(self.machines_manager_frame)
        bottom_frame.pack(fill="x", padx=5, pady=(10,0))

        group1_frame = ttk.LabelFrame(bottom_frame, text="Machine Actions", padding=5)
        group1_frame.pack(side="left", padx=5)
        ttk.Button(group1_frame, text="Test Selected", command=self.test_selected_connection).pack(side="left", padx=5)
        ttk.Button(group1_frame, text="Edit Selected", command=self.edit_connection).pack(side="left", padx=5)
        ttk.Button(group1_frame, text="Delete Selected", command=self.delete_connection).pack(side="left", padx=5)

        group2_frame = ttk.LabelFrame(bottom_frame, text="Data Actions", padding=5)
        group2_frame.pack(side="left", padx=5)
        ttk.Button(group2_frame, text="Link Machine to Odoo", command=self.link_machine_to_odoo).pack(side="left", padx=5)
        ttk.Button(group2_frame, text="Fetch New Data", command=self.fetch_data_from_device_manual).pack(side="left", padx=5)
        ttk.Button(group2_frame, text="Sync to Odoo", command=self.sync_to_odoo).pack(side="left", padx=5)

        status_frame = ttk.LabelFrame(bottom_frame, text="Status", padding=5)
        status_frame.pack(side="left", padx=5, fill="x", expand=True)
        self.status_text = tk.StringVar()
        self.status_text.set("Ready")
        status_label = ttk.Label(status_frame, textvariable=self.status_text, anchor="w")
        status_label.pack(fill="x", expand=True, padx=5, pady=2)

    def create_users_tab(self):
        filter_frame = ttk.Frame(self.users_frame, padding="5")
        filter_frame.pack(fill="x")
        ttk.Label(filter_frame, text="Filter by Machine:").pack(side="left")
        self.user_machine_filter = ttk.Combobox(filter_frame, state="readonly")
        self.user_machine_filter.pack(side="left", padx=5)
        self.user_machine_filter.bind("<<ComboboxSelected>>", self.update_users_table)
        
        ttk.Checkbutton(filter_frame, text="Show only records to sync", variable=self.users_to_sync_var, command=self.update_users_table).pack(side="left", padx=10)
        
        table_frame = ttk.Frame(self.users_frame)
        table_frame.pack(fill="both", expand=True, padx=5, pady=5)
        columns = ('machine_name', 'user_id', 'name', 'uid', 'synched_time')
        self.users_table = ttk.Treeview(table_frame, columns=columns, show='headings')
        self.users_table.heading('machine_name', text='Machine')
        self.users_table.heading('user_id', text='User ID')
        self.users_table.heading('name', text='Name')
        self.users_table.heading('uid', text='Internal UID')
        self.users_table.heading('synched_time', text='Synched Time')
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.users_table.yview)
        scrollbar.pack(side="right", fill="y")
        self.users_table.configure(yscrollcommand=scrollbar.set)
        self.users_table.pack(side="left", fill="both", expand=True)

    def create_attendance_tab(self):
        filter_frame = ttk.Frame(self.attendance_frame, padding="5")
        filter_frame.pack(fill="x")
        ttk.Label(filter_frame, text="Filter by Machine:").pack(side="left")
        self.attendance_machine_filter = ttk.Combobox(filter_frame, state="readonly")
        self.attendance_machine_filter.pack(side="left", padx=5)
        self.attendance_machine_filter.bind("<<ComboboxSelected>>", self.update_attendance_table)

        ttk.Checkbutton(filter_frame, text="Show only records to sync", variable=self.attendance_to_sync_var, command=self.update_attendance_table).pack(side="left", padx=10)

        # Deletion controls for attendance
        delete_frame = ttk.Frame(filter_frame)
        delete_frame.pack(side="right")

        ttk.Button(delete_frame, text="Delete All", command=self.delete_all_attendance).pack(side="right", padx=5)
        ttk.Button(delete_frame, text="Delete", command=self.delete_old_attendance).pack(side="right", padx=2)
        self.att_delete_period = ttk.Combobox(delete_frame, state="readonly", width=10, values=["30 Days", "60 Days", "90 Days", "180 Days", "365 Days"])
        self.att_delete_period.pack(side="right")
        self.att_delete_period.set("90 Days") 
        ttk.Label(delete_frame, text="Delete Older Than:").pack(side="right", padx=2)

        table_frame = ttk.Frame(self.attendance_frame)
        table_frame.pack(fill="both", expand=True, padx=5, pady=5)
        columns = ('machine_name', 'user_id', 'timestamp', 'synched_time')
        self.attendance_table = ttk.Treeview(table_frame, columns=columns, show='headings')
        self.attendance_table.heading('machine_name', text='Machine')
        self.attendance_table.heading('user_id', text='User ID')
        self.attendance_table.heading('timestamp', text='Timestamp (Machine Time)')
        self.attendance_table.heading('synched_time', text='Synched Time')
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.attendance_table.yview)
        scrollbar.pack(side="right", fill="y")
        self.attendance_table.configure(yscrollcommand=scrollbar.set)
        self.attendance_table.pack(side="left", fill="both", expand=True)

    def create_scheduler_tab(self):
        status_frame = ttk.LabelFrame(self.scheduler_frame, text="Scheduler Control", padding=10)
        status_frame.pack(fill="x", pady=10)
        
        # --- UI Variables ---
        self.scheduler_status_var = tk.StringVar(value="Stopped")
        self.last_run_var = tk.StringVar(value="N/A")
        self.next_run_var = tk.StringVar(value="N/A")
        
        # --- Control Buttons ---
        self.start_scheduler_btn = ttk.Button(status_frame, text="Start Scheduler", command=self.start_scheduler)
        self.start_scheduler_btn.grid(row=0, column=0, padx=5, pady=5)
        self.stop_scheduler_btn = ttk.Button(status_frame, text="Stop Scheduler", command=self.stop_scheduler, state="disabled")
        self.stop_scheduler_btn.grid(row=0, column=1, padx=5, pady=5)

        # --- Status Display ---
        ttk.Label(status_frame, text="Status:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        ttk.Label(status_frame, textvariable=self.scheduler_status_var, font=("TkDefaultFont", 10, "bold")).grid(row=1, column=1, sticky="w", padx=5, pady=5)
        
        ttk.Label(status_frame, text="Last Run:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        ttk.Label(status_frame, textvariable=self.last_run_var).grid(row=2, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(status_frame, text="Next Run:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        ttk.Label(status_frame, textvariable=self.next_run_var).grid(row=3, column=1, sticky="w", padx=5, pady=5)

        # --- Last Run Logs Display ---
        last_run_logs_frame = ttk.LabelFrame(self.scheduler_frame, text="Last Run Logs", padding=10)
        last_run_logs_frame.pack(fill="both", expand=True, pady=10)
        
        log_columns = ('machine_name', 'timestamp', 'operation', 'message')
        self.last_run_logs_table = ttk.Treeview(last_run_logs_frame, columns=log_columns, show='headings')
        self.last_run_logs_table.heading('machine_name', text='Machine')
        self.last_run_logs_table.heading('timestamp', text='Timestamp')
        self.last_run_logs_table.heading('operation', text='Operation')
        self.last_run_logs_table.heading('message', text='Message')
        self.last_run_logs_table.column('message', width=400)
        self.last_run_logs_table.column('operation', width=120)
        self.last_run_logs_table.column('machine_name', width=120)
        log_scrollbar = ttk.Scrollbar(last_run_logs_frame, orient=tk.VERTICAL, command=self.last_run_logs_table.yview)
        log_scrollbar.pack(side="right", fill="y")
        self.last_run_logs_table.configure(yscrollcommand=log_scrollbar.set)
        self.last_run_logs_table.pack(fill="both", expand=True)

    def create_logs_tab(self):
        controls_frame = ttk.Frame(self.logs_frame, padding="5")
        controls_frame.pack(fill="x")

        # Filter
        ttk.Label(controls_frame, text="Filter by Machine:").pack(side="left")
        self.log_machine_filter = ttk.Combobox(controls_frame, state="readonly")
        self.log_machine_filter.pack(side="left", padx=5)
        self.log_machine_filter.bind("<<ComboboxSelected>>", self.update_logs_table)

        # Deletion Controls
        delete_frame = ttk.Frame(controls_frame)
        delete_frame.pack(side="right")

        ttk.Button(delete_frame, text="Delete All Logs", command=self.delete_all_logs).pack(side="right", padx=5)
        
        ttk.Button(delete_frame, text="Delete", command=self.delete_old_logs).pack(side="right", padx=2)
        self.log_delete_period = ttk.Combobox(delete_frame, state="readonly", width=10, values=["1 Day", "7 Days", "30 Days"])
        self.log_delete_period.pack(side="right")
        self.log_delete_period.set("7 Days")
        ttk.Label(delete_frame, text="Delete Older Than:").pack(side="right", padx=2)

        # Log Table
        table_frame = ttk.Frame(self.logs_frame)
        table_frame.pack(fill="both", expand=True, padx=5, pady=5)
        columns = ('machine_name', 'timestamp', 'operation', 'message')
        self.logs_table = ttk.Treeview(table_frame, columns=columns, show='headings')
        self.logs_table.heading('machine_name', text='Machine')
        self.logs_table.heading('timestamp', text='Timestamp')
        self.logs_table.heading('operation', text='Operation')
        self.logs_table.heading('message', text='Message')
        self.logs_table.column('message', width=400)
        self.logs_table.column('operation', width=120)
        self.logs_table.column('machine_name', width=120)

        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.logs_table.yview)
        scrollbar.pack(side="right", fill="y")
        self.logs_table.configure(yscrollcommand=scrollbar.set)
        self.logs_table.pack(side="left", fill="both", expand=True)
    
    def create_settings_tab(self):
        settings_frame = ttk.LabelFrame(self.settings_frame, text="Application Settings", padding="10")
        settings_frame.pack(fill="x", padx=10, pady=10)

        # Days to go back
        ttk.Label(settings_frame, text="Days to Go Back on First Pull:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.setting_days_back = ttk.Combobox(settings_frame, state="readonly", values=[1, 5, 10, 31, 60, 90])
        self.setting_days_back.grid(row=0, column=1, sticky="w", padx=5, pady=5)

        # Batch size
        ttk.Label(settings_frame, text="Batch Size to Upload to Odoo:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.setting_batch_size = ttk.Combobox(settings_frame, state="readonly", values=[1, 50, 100, 500, 1000])
        self.setting_batch_size.grid(row=1, column=1, sticky="w", padx=5, pady=5)

        # Scheduler delay
        ttk.Label(settings_frame, text="Scheduler Delay (minutes):").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.setting_scheduler_delay = ttk.Combobox(settings_frame, state="readonly", values=[5, 15, 30, 60, 120, 240, 720, 1440])
        self.setting_scheduler_delay.grid(row=2, column=1, sticky="w", padx=5, pady=5)
        
        # Clean logs older than
        ttk.Label(settings_frame, text="Clean Logs Older Than (days):").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.setting_clean_logs_days = ttk.Combobox(settings_frame, state="readonly", values=[5, 30, 60, 90])
        self.setting_clean_logs_days.grid(row=3, column=1, sticky="w", padx=5, pady=5)
        
        # Delete attendance older than
        ttk.Label(settings_frame, text="Delete Attendance Older Than (days):").grid(row=4, column=0, sticky="w", padx=5, pady=5)
        self.setting_delete_attendance_days = ttk.Combobox(settings_frame, state="readonly", values=[30, 60, 90, 180, 360])
        self.setting_delete_attendance_days.grid(row=4, column=1, sticky="w", padx=5, pady=5)

        ttk.Button(settings_frame, text="Save Settings", command=self.save_settings).grid(row=5, column=0, columnspan=2, pady=10)

    def create_odoo_tab(self):
        # Main frame to hold the two columns
        main_odoo_frame = ttk.Frame(self.odoo_frame)
        main_odoo_frame.pack(fill="both", expand=True)
        main_odoo_frame.grid_columnconfigure(0, weight=1, minsize=400) # Left column for details
        main_odoo_frame.grid_columnconfigure(1, weight=1) # Right column for promo

        # --- Left Column: Details and Results ---
        left_column_frame = ttk.Frame(main_odoo_frame)
        left_column_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        odoo_details_frame = ttk.LabelFrame(left_column_frame, text="Odoo Server Details", padding="10")
        odoo_details_frame.pack(fill="x", pady=5)
        
        ttk.Label(odoo_details_frame, text="Odoo Server URL:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.odoo_url_entry = ttk.Entry(odoo_details_frame, width=40)
        self.odoo_url_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        
        ttk.Label(odoo_details_frame, text="DB Name:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.odoo_db_entry = ttk.Entry(odoo_details_frame)
        self.odoo_db_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=5)

        ttk.Label(odoo_details_frame, text="Username:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.odoo_user_entry = ttk.Entry(odoo_details_frame)
        self.odoo_user_entry.grid(row=2, column=1, sticky="ew", padx=5, pady=5)

        ttk.Label(odoo_details_frame, text="Password:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.odoo_pass_entry = ttk.Entry(odoo_details_frame, show="*")
        self.odoo_pass_entry.grid(row=3, column=1, sticky="ew", padx=5, pady=5)
        
        odoo_details_frame.columnconfigure(1, weight=1)

        button_frame = ttk.Frame(left_column_frame)
        button_frame.pack(pady=10, fill="x")
        ttk.Button(button_frame, text="Save Odoo Details", command=self.save_odoo_details).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Test Odoo Connection", command=self.test_odoo_connection).pack(side="left", padx=5)

        # Odoo Test Result Area
        result_frame = ttk.LabelFrame(left_column_frame, text="Test Result", padding="10")
        result_frame.pack(fill="both", expand=True, pady=10)
        self.odoo_result_text = ScrolledText(result_frame, height=8, state='disabled', wrap=tk.WORD)
        self.odoo_result_text.pack(fill="both", expand=True)
        
        # --- Right Column: Promotional Message ---
        right_column_frame = ttk.Frame(main_odoo_frame)
        right_column_frame.grid(row=0, column=1, sticky="nsew", padx=(10, 0))

        promo_frame = ttk.LabelFrame(right_column_frame, text="Azkatech Integration", padding="15")
        promo_frame.pack(fill="both", expand=True)

        promo_text = "This Software syncs with Azkatech's all-in-one Attendance integration for Odoo, which can be found on this link:"
        promo_label = ttk.Label(promo_frame, text=promo_text, wraplength=300, justify="left")
        promo_label.pack(pady=(5,10), anchor="w")

        link_url = "https://apps.odoo.com/apps/modules/18.0/azk_zkteco_attendance"
        link_label = ttk.Label(promo_frame, text=link_url, foreground="blue", cursor="hand2", wraplength=300, justify="left")
        link_label.pack(pady=5, anchor="w")
        link_label.bind("<Button-1>", lambda e: self.open_link(link_url))

        # This is a placeholder for a logo. You could load an image here if you have one.
        logo_label = ttk.Label(promo_frame, text="AZKATECH", font=("TkDefaultFont", 16, "bold"))
        logo_label.pack(pady=10)
        
        # Container for copyright and version
        footer_frame = ttk.Frame(promo_frame)
        footer_frame.pack(side="bottom", pady=(10,5))
        
        copyright_label = ttk.Label(footer_frame, text="© 2025 Azkatech")
        copyright_label.pack()

        version_label = ttk.Label(footer_frame, text=f"Version: {VERSION_NUM}")
        version_label.pack()


    def open_link(self, url):
        webbrowser.open_new(url)

    # --- Data Loading and Table Updates ---
    def load_connections_from_db(self):
        rows = db_execute("SELECT * FROM zkteco_machines ORDER BY name", fetch='all')
        self.connections_list = [dict(row) for row in rows]
        # Update filter dropdowns
        machine_names = ["All"] + [conn['name'] for conn in self.connections_list]
        self.user_machine_filter['values'] = machine_names
        self.user_machine_filter.set("All")
        self.attendance_machine_filter['values'] = machine_names
        self.attendance_machine_filter.set("All")
        self.log_machine_filter['values'] = machine_names
        self.log_machine_filter.set("All")

    def load_odoo_details_from_db(self):
        rows = db_execute("SELECT key, value FROM odoo_config", fetch='all')
        self.odoo_details = {row['key']: row['value'] for row in rows}
        self.odoo_url_entry.delete(0, tk.END)
        self.odoo_db_entry.delete(0, tk.END)
        self.odoo_user_entry.delete(0, tk.END)
        self.odoo_pass_entry.delete(0, tk.END)
        self.odoo_url_entry.insert(0, self.odoo_details.get("url", ""))
        self.odoo_db_entry.insert(0, self.odoo_details.get("db", ""))
        self.odoo_user_entry.insert(0, self.odoo_details.get("username", ""))
        self.odoo_pass_entry.insert(0, self.odoo_details.get("password", ""))

    def load_settings_from_db(self):
        rows = db_execute("SELECT key, value FROM settings", fetch='all')
        self.settings = {row['key']: row['value'] for row in rows}
        
        # Set UI elements with defaults if not found in DB
        self.setting_days_back.set(self.settings.get("days_back", 31))
        self.setting_batch_size.set(self.settings.get("batch_size", 1000))
        self.setting_scheduler_delay.set(self.settings.get("scheduler_delay", 10))
        self.setting_clean_logs_days.set(self.settings.get("clean_logs_days", 30))
        self.setting_delete_attendance_days.set(self.settings.get("delete_attendance_days", 180))
    
    def update_connections_table(self):
        self.connections_table.delete(*self.connections_table.get_children())
        self.load_connections_from_db() # Reload from DB to get fresh data
        for conn_dict in self.connections_list:
            values = (conn_dict.get('name', ''), conn_dict.get('ip', ''), conn_dict.get('port', ''),
                      conn_dict.get('odoo_machine_name', ''), conn_dict.get('odoo_machine_id', ''),
                      conn_dict.get('machine_timezone', 'N/A'),
                      conn_dict.get('serial_number', 'N/A'), conn_dict.get('last_connected', 'N/A'))
            self.connections_table.insert('', tk.END, iid=conn_dict['id'], values=values)

    def update_users_table(self, event=None):
        self.users_table.delete(*self.users_table.get_children())
        selected_machine = self.user_machine_filter.get()
        query = "SELECT u.user_id, u.name, u.uid, u.synched_time, c.name as machine_name FROM users u JOIN zkteco_machines c ON u.connection_id = c.id"
        
        where_clauses = []
        params = []

        if selected_machine != "All":
            where_clauses.append("c.name = ?")
            params.append(selected_machine)
        
        if self.users_to_sync_var.get():
            where_clauses.append("u.synched_time IS NULL")

        if where_clauses:
            query += " WHERE " + " AND ".join(where_clauses)
        
        query += " ORDER BY c.name, u.name"

        for row in db_execute(query, tuple(params), fetch='all'):
            self.users_table.insert('', tk.END, values=(row['machine_name'], row['user_id'], row['name'], row['uid'], row['synched_time'] or ''))

    def update_attendance_table(self, event=None):
        self.attendance_table.delete(*self.attendance_table.get_children())
        selected_machine = self.attendance_machine_filter.get()
        query = """SELECT a.user_id, a.timestamp, a.synched_time, c.name as machine_name FROM attendance a 
        JOIN zkteco_machines c ON a.connection_id = c.id"""
        
        where_clauses = []
        params = []
        
        if selected_machine != "All":
            where_clauses.append("c.name = ?")
            params.append(selected_machine)

        if self.attendance_to_sync_var.get():
            where_clauses.append("a.synched_time IS NULL")

        if where_clauses:
            query += " WHERE " + " AND ".join(where_clauses)

        query += " ORDER BY a.timestamp desc"

        for row in db_execute(query, tuple(params), fetch='all'):
            # Insert at the top (index 0) to show newest first
            self.attendance_table.insert('', 0, values=(row['machine_name'], row['user_id'], row['timestamp'], row['synched_time'] or ''))

    def update_logs_table(self, event=None):
        self.logs_table.delete(*self.logs_table.get_children())
        selected_machine = self.log_machine_filter.get()
        query = "SELECT l.timestamp, l.operation, l.message, c.name as machine_name FROM logs l LEFT JOIN zkteco_machines c ON l.connection_id = c.id"
        params = ()
        if selected_machine != "All":
            machine_id = self.get_connection_id_from_name(selected_machine)
            if machine_id:
                query += " WHERE c.id = ?"
                params = (machine_id,)
        query += " ORDER BY l.timestamp DESC"

        for row in db_execute(query, params, fetch='all'):
            machine_name = row['machine_name'] if row['machine_name'] else "System"
            # Insert at the top (index 0) to show newest first
            self.logs_table.insert('', 0, values=(machine_name, row['timestamp'], row['operation'], row['message']))

    def _update_last_run_logs_table(self, run_start_time):
        self.last_run_logs_table.delete(*self.last_run_logs_table.get_children())
        start_time_str = run_start_time.strftime("%Y-%m-%d %H:%M:%S")
        query = """SELECT l.timestamp, l.operation, l.message, c.name as machine_name 
        FROM logs l LEFT JOIN zkteco_machines c ON l.connection_id = c.id WHERE l.timestamp >= ? 
        ORDER BY l.timestamp desc"""
        
        for row in db_execute(query, (start_time_str,), fetch='all'):
            machine_name = row['machine_name'] if row['machine_name'] else "System"
            self.last_run_logs_table.insert('', tk.END, values=(machine_name, row['timestamp'], row['operation'], row['message']))

    # --- Connection Tab Methods ---
    def clear_connection_entries(self):
        self.name_entry.delete(0, tk.END)
        self.ip_entry.delete(0, tk.END)
        self.port_entry.delete(0, tk.END)
        self.pass_entry.delete(0, tk.END)
        self.odoo_name_entry.delete(0, tk.END)
        self.odoo_id_entry.delete(0, tk.END)
        self.timezone_entry.config(state="normal")
        self.timezone_entry.delete(0, tk.END)
        self.timezone_entry.config(state="readonly")
        self.port_entry.insert(0, "4370")
        self.pass_entry.insert(0, "0")

    def add_connection(self):
        name = self.name_entry.get().strip()
        ip = self.ip_entry.get().strip()
        port = self.port_entry.get().strip() or "4370"
        password = self.pass_entry.get()
        odoo_name = self.odoo_name_entry.get().strip()
        odoo_id_str = self.odoo_id_entry.get().strip()
        odoo_id = int(odoo_id_str) if odoo_id_str.isdigit() else None

        if not name or not ip:
            messagebox.showwarning("Warning", "Name and IP Address are required.")
            return
        query = "INSERT INTO zkteco_machines (name, ip, port, password, odoo_machine_name, odoo_machine_id) VALUES (?, ?, ?, ?, ?, ?)"
        db_execute(query, (name, ip, port, password, odoo_name, odoo_id))
        self.refresh_all_data()
        self.clear_connection_entries()

    def edit_connection(self):
        selected_iid = self.connections_table.focus()
        if not selected_iid:
            messagebox.showinfo("Info", "Please select a machine to edit.")
            return
        self.editing_connection_id = int(selected_iid)
        conn_details = db_execute("SELECT * FROM zkteco_machines WHERE id = ?", (self.editing_connection_id,), fetch='one')
        
        self.clear_connection_entries()
        self.name_entry.insert(0, conn_details['name'])
        self.ip_entry.insert(0, conn_details['ip'])
        self.port_entry.insert(0, conn_details['port'])
        self.pass_entry.insert(0, conn_details['password'])
        self.odoo_name_entry.insert(0, conn_details['odoo_machine_name'] or '')
        self.odoo_id_entry.insert(0, str(conn_details['odoo_machine_id'] or ''))
        self.timezone_entry.config(state="normal")
        self.timezone_entry.insert(0, conn_details['machine_timezone'] or '')
        self.timezone_entry.config(state="readonly")

        self.add_edit_frame.config(text="Edit Machine")
        self.add_button.pack_forget()
        self.save_button.pack(side="left", padx=5)
        self.cancel_button.pack(side="left", padx=5)

    def save_connection_changes(self):
        if self.editing_connection_id is None: return
        name = self.name_entry.get().strip()
        ip = self.ip_entry.get().strip()
        port = self.port_entry.get().strip() or "4370"
        password = self.pass_entry.get()
        odoo_name = self.odoo_name_entry.get().strip()
        odoo_id_str = self.odoo_id_entry.get().strip()
        odoo_id = int(odoo_id_str) if odoo_id_str.isdigit() else None

        if not name or not ip:
            messagebox.showwarning("Warning", "Name and IP Address are required.")
            return
        query = "UPDATE zkteco_machines SET name=?, ip=?, port=?, password=?, odoo_machine_name=?, odoo_machine_id=? WHERE id=?"
        db_execute(query, (name, ip, port, password, odoo_name, odoo_id, self.editing_connection_id))
        self.refresh_all_data()
        self.cancel_edit()

    def cancel_edit(self):
        self.editing_connection_id = None
        self.clear_connection_entries()
        self.add_edit_frame.config(text="Add new Machine")
        self.save_button.pack_forget()
        self.cancel_button.pack_forget()
        self.add_button.pack()

    def delete_connection(self):
        selected_iid = self.connections_table.focus()
        if not selected_iid:
            messagebox.showinfo("Info", "Please select a machine to delete.")
            return
        if messagebox.askyesno("Confirm", "Are you sure you want to delete this machine and all its related user/attendance/log data?"):
            db_execute("DELETE FROM zkteco_machines WHERE id = ?", (int(selected_iid),))
            self.refresh_all_data()

    # --- Log and Data Cleanup Methods ---
    def log_operation(self, operation, message, conn_id=None):
        """Adds a record to the logs table and refreshes the UI."""
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        db_execute("INSERT INTO logs (connection_id, timestamp, operation, message) VALUES (?, ?, ?, ?)",
                   (conn_id, now, operation, message))
        self.after(0, self.update_logs_table)
        self.after(0, self.update_status, message)

    def update_status(self, message):
        self.status_text.set(message)

    def get_connection_id_from_name(self, name):
        """Helper to get a connection ID from its name."""
        if not name: return None
        for conn in self.connections_list:
            if conn['name'] == name:
                return conn['id']
        return None

    def delete_all_logs(self):
        if messagebox.askyesno("Confirm", "Are you sure you want to delete ALL log entries? This cannot be undone."):
            db_execute("DELETE FROM logs")
            self.update_logs_table()

    def delete_old_logs(self):
        period = self.log_delete_period.get()
        days_map = {"1 Day": 1, "7 Days": 7, "30 Days": 30}
        days = days_map.get(period)
        if not days:
            messagebox.showwarning("Warning", "Please select a valid period.")
            return
        
        filter_machine_name = self.log_machine_filter.get()
        confirm_message = f"Are you sure you want to delete logs older than {days} day(s)"
        query = "DELETE FROM logs WHERE timestamp < ?"
        params = [datetime.now() - timedelta(days=days)]

        if filter_machine_name != "All":
            conn_id = self.get_connection_id_from_name(filter_machine_name)
            if conn_id:
                query += " AND connection_id = ?"
                params.append(conn_id)
                confirm_message += f" for machine '{filter_machine_name}'?"
            else:
                confirm_message += "?"
        else:
            confirm_message += " for ALL machines?"
        
        if messagebox.askyesno("Confirm", confirm_message):
            db_execute(query, tuple(params))
            self.update_logs_table()

    def delete_all_attendance(self):
        """Deletes all attendance records after confirmation."""
        if messagebox.askyesno("Confirm Deletion", "Are you sure you want to delete ALL attendance records from the local database? This cannot be undone."):
            db_execute("DELETE FROM attendance")
            self.log_operation("Data Cleanup", "Deleted all attendance records.")
            self.update_attendance_table()

    def delete_old_attendance(self):
        """Deletes attendance records older than the selected period."""
        period = self.att_delete_period.get()
        days_map = {"30 Days": 30, "60 Days": 60, "90 Days": 90, "180 Days": 180, "365 Days": 365}
        days = days_map.get(period)
        if not days:
            messagebox.showwarning("Warning", "Please select a valid period to delete.")
            return
        
        filter_machine_name = self.attendance_machine_filter.get()
        confirm_message = f"Are you sure you want to delete attendance records older than {days} day(s)"
        query = "DELETE FROM attendance WHERE timestamp < ?"
        params = [datetime.now() - timedelta(days=days)]

        if filter_machine_name != "All":
            conn_id = self.get_connection_id_from_name(filter_machine_name)
            if conn_id:
                query += " AND connection_id = ?"
                params.append(conn_id)
                confirm_message += f" for machine '{filter_machine_name}'?"
            else: # Should not happen if UI is consistent
                confirm_message += "?"
        else:
            confirm_message += " for ALL machines?"
        
        if messagebox.askyesno("Confirm Deletion", confirm_message):
            db_execute(query, tuple(params))
            self.log_operation("Data Cleanup", f"Deleted attendance older than {days} days for '{filter_machine_name}'.")
            self.update_attendance_table()


    # --- Device Interaction ---
    def _get_selected_connection(self):
        selected_iid = self.connections_table.focus()
        if not selected_iid:
            messagebox.showinfo("Info", "Please select a machine from the table first.")
            return None, None
        conn_id = int(selected_iid)
        conn_dict = db_execute("SELECT * FROM zkteco_machines WHERE id = ?", (conn_id,), fetch='one')
        return conn_id, conn_dict
    
    def test_selected_connection(self):
        conn_id, conn_dict = self._get_selected_connection()
        if not conn_dict: return
        threading.Thread(target=self._test_connection_thread, args=(conn_id, conn_dict), daemon=True).start()
        
    def _test_connection_thread(self, conn_id, conn_dict):
        self.log_operation("Test Connection", f"Attempting to connect to {conn_dict['name']}...", conn_id)
        zk = None
        try:
            zk = ZK(conn_dict['ip'], port=int(conn_dict['port']), password=conn_dict['password'], timeout=5)
            conn = zk.connect()
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            serial_num = conn.get_serialnumber()
            db_execute("UPDATE zkteco_machines SET last_connected = ?, serial_number = ? WHERE id = ?", (now, serial_num, conn_id))
            self.log_operation("Test Connection", f"Success! Serial: {serial_num}", conn_id)
            self.after(0, self.update_connections_table)
            self.after(0, lambda: messagebox.showinfo("Success", f"Successfully connected to {conn_dict['name']}!"))
        except Exception as e:
            self.log_operation("Test Connection", f"Failed: {e}", conn_id)
            self.after(0, lambda: messagebox.showerror("Connection Failed", f"Could not connect to the device.\nError: {e}"))
        finally:
            if zk and zk.is_connect:
                zk.disconnect()
            self.after(0, self.update_status, "Ready")
    
    def fetch_data_from_device_manual(self):
        conn_id, conn_dict = self._get_selected_connection()
        if not conn_dict: return
        threading.Thread(target=self._fetch_data_for_machine, args=(conn_id, conn_dict, True), daemon=True).start()

    def _fetch_data_for_machine(self, conn_id, conn_dict, is_manual):
        self.log_operation("Fetch Data", f"Starting data fetch for {conn_dict['name']}", conn_id)
        last_timestamp_row = db_execute("SELECT MAX(timestamp) as last_ts FROM attendance WHERE connection_id = ?", (conn_id,), fetch='one')
        date_from = None
        if last_timestamp_row and last_timestamp_row['last_ts']:
            last_timestamp = datetime.strptime(last_timestamp_row['last_ts'], "%Y-%m-%d %H:%M:%S")
            date_from = last_timestamp
            self.log_operation("Fetch Data", f"Last record at {last_timestamp}. Fetching newer.", conn_id)
        else:
            days_back = int(self.settings.get("days_back", 31))
            date_from = datetime.now() - timedelta(days=days_back)
            self.log_operation("Fetch Data", f"No previous records. Fetching last {days_back} days.", conn_id)

        #always pull one full day
        date_from = date_from.replace(hour=0, minute=0, second=0, microsecond=0)

        zk = None
        try:
            zk = ZK(conn_dict['ip'], port=int(conn_dict['port']), password=conn_dict['password'], timeout=15)
            self.log_operation("Fetch Data", "Connecting to device...", conn_id)
            conn = zk.connect()
            self.log_operation("Fetch Data", "Synchronizing users...", conn_id)
            users = conn.get_users()
            user_count = 0
            user_query = "INSERT OR REPLACE INTO users (connection_id, uid, user_id, name, synched_time) VALUES (?, ?, ?, ?, (SELECT synched_time FROM users WHERE connection_id=? AND user_id=?))"
            for user in users:
                db_execute(user_query, (conn_id, user.uid, user.user_id, user.name, conn_id, user.user_id))
                user_count += 1
            self.log_operation("Fetch Data", f"Users synchronized: {user_count}", conn_id)

            self.log_operation("Fetch Data", "Downloading attendance...", conn_id)
            all_device_attendance = conn.get_attendance()
            now = datetime.now()
            
            new_attendance_records = [att for att in all_device_attendance if att.timestamp > date_from and att.timestamp <= now]
            
            att_count = 0
            if new_attendance_records:
                att_query = "INSERT OR IGNORE INTO attendance (connection_id, user_id, att_id, timestamp) VALUES (?, ?, ?, ?)"
                for att in new_attendance_records:
                    # The timestamp from the device is naive, representing local time on the device
                    # Create a unique ID using connection_id, user_id, and timestamp
                    att_id = f"{conn_id}-{att.user_id}-{att.timestamp.strftime('%Y%m%d%H%M%S')}"
                    timestamp_str = att.timestamp.strftime("%Y-%m-%d %H:%M:%S")
                    db_execute(att_query, (conn_id, att.user_id, att_id, timestamp_str))
                    att_count += 1
            
            self.log_operation("Success", f"Fetch complete. New records: {att_count}", conn_id)
            self.after(0, self.refresh_all_data)
            if is_manual:
                self.after(0, lambda: messagebox.showinfo("Success", f"Data fetch complete for {conn_dict['name']}.\n- Users: {user_count}\n- New attendance: {att_count}"))

        except Exception as e:
            self.log_operation("Error", f"Fetch Failed: {e}", conn_id)
            if is_manual:
                err_msg = f"An error occurred. Check logs for details.\nError: {e}"
                self.after(0, lambda: messagebox.showerror("Operation Failed", err_msg))
        finally:
            if zk and zk.is_connect:
                zk.disconnect()
            self.after(0, self.update_status, "Ready")

    # --- Odoo Methods ---
    def save_settings(self):
        settings_to_save = {
            "days_back": self.setting_days_back.get(),
            "batch_size": self.setting_batch_size.get(),
            "scheduler_delay": self.setting_scheduler_delay.get(),
            "clean_logs_days": self.setting_clean_logs_days.get(),
            "delete_attendance_days": self.setting_delete_attendance_days.get()
        }
        query = "INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)"
        for key, value in settings_to_save.items():
            db_execute(query, (key, value))
        
        self.load_settings_from_db() # Reload settings into memory
        messagebox.showinfo("Success", "Settings saved successfully.")

    def save_odoo_details(self):
        details = {"url": self.odoo_url_entry.get(), "db": self.odoo_db_entry.get(), "username": self.odoo_user_entry.get(), "password": self.odoo_pass_entry.get()}
        query = "INSERT OR REPLACE INTO odoo_config (key, value) VALUES (?, ?)"
        for key, value in details.items():
            db_execute(query, (key, value))
        self.load_odoo_details_from_db() # Reload details after saving
        messagebox.showinfo("Success", "Odoo details saved successfully.")

    def test_odoo_connection(self):
        url = self.odoo_url_entry.get()
        db = self.odoo_db_entry.get()
        username = self.odoo_user_entry.get()
        password = self.odoo_pass_entry.get()

        self.odoo_result_text.config(state='normal', bg='SystemButtonFace')
        self.odoo_result_text.delete('1.0', tk.END)

        if not all([url, db, username, password]):
            self.odoo_result_text.insert(tk.END, "Error: Please fill in all Odoo connection details.")
            self.odoo_result_text.config(state='disabled', bg='#FFCCCC') # Light red
            return
        
        try:
            self.odoo_result_text.insert(tk.END, f"Attempting to connect to {url}...\n")
            self.update()

            common = xmlrpc.client.ServerProxy(f'{url}/xmlrpc/2/common')
            version_info = common.version()
            
            self.odoo_result_text.insert(tk.END, "Server reached. Checking version...\n")
            self.odoo_result_text.insert(tk.END, f"  - Odoo Version: {version_info.get('server_version')}\n")
            self.update()

            self.odoo_result_text.insert(tk.END, "Authenticating...\n")
            self.update()
            uid = common.authenticate(db, username, password, {})
            
            if not uid:
                raise Exception("Authentication failed. Please check DB, username, and password.")

            self.odoo_result_text.insert(tk.END, f"Authentication successful! User ID: {uid}\n\n")
            self.odoo_result_text.insert(tk.END, "Connection test successful!")
            self.odoo_result_text.config(bg='#CCFFCC') # Light green

        except Exception as e:
            self.odoo_result_text.config(bg='#FFCCCC') # Light red
            self.odoo_result_text.insert(tk.END, f"\nConnection Failed:\n\n{e}")
        finally:
            self.odoo_result_text.config(state='disabled')

    def link_machine_to_odoo(self):
        conn_id, conn_dict = self._get_selected_connection()
        if not conn_dict: return
        threading.Thread(target=self._link_machine_thread, args=(conn_id, conn_dict), daemon=True).start()

    def _link_machine_thread(self, conn_id, conn_dict):
        odoo_machine_name = conn_dict['odoo_machine_name']
        
        if not odoo_machine_name:
            self.after(0, lambda: messagebox.showwarning("Warning", "The selected machine does not have an 'Odoo Machine Name' set."))
            return

        url = self.odoo_details.get("url")
        db = self.odoo_details.get("db")
        username = self.odoo_details.get("username")
        password = self.odoo_details.get("password")

        if not all([url, db, username, password]):
            self.after(0, lambda: messagebox.showerror("Error", "Odoo connection details are not fully configured in the Odoo Connection tab."))
            return
        
        try:
            common = xmlrpc.client.ServerProxy(f'{url}/xmlrpc/2/common')
            uid = common.authenticate(db, username, password, {})
            if not uid:
                self.after(0, lambda: messagebox.showerror("Odoo Auth Failed", "Could not authenticate with Odoo. Check credentials."))
                return

            models = xmlrpc.client.ServerProxy(f'{url}/xmlrpc/2/object')
            
            self.log_operation("Link to Odoo", f"Searching for '{odoo_machine_name}' in Odoo.", conn_id)
            search_read_result = models.execute_kw(db, uid, password, 'azk.machine', 'search_read', 
                                                   [[['name', '=', odoo_machine_name]]], 
                                                   {'fields': ['id', 'timezone'], 'limit': 1})

            if not search_read_result:
                self.after(0, lambda: messagebox.showerror("Not Found", f"No machine named '{odoo_machine_name}' found in Odoo."))
                self.log_operation("Link to Odoo", f"Machine '{odoo_machine_name}' not found.", conn_id)
                return
            if len(search_read_result) > 1:
                self.after(0, lambda: messagebox.showwarning("Multiple Found", f"Multiple machines found with name '{odoo_machine_name}'."))
                self.log_operation("Link to Odoo", "Multiple machines found. Aborting.", conn_id)
                return

            odoo_data = search_read_result[0]
            machine_id_in_odoo = odoo_data['id']
            machine_timezone = odoo_data.get('timezone', False) or None # Odoo returns False for empty selection fields
            
            db_execute("UPDATE zkteco_machines SET odoo_machine_id = ?, machine_timezone = ? WHERE id = ?", 
                       (machine_id_in_odoo, machine_timezone, conn_id))
            
            self.log_operation("Success", f"Link successful. Odoo ID {machine_id_in_odoo} and TZ '{machine_timezone}' saved.", conn_id)
            self.after(0, self.update_connections_table)
            self.after(0, lambda: messagebox.showinfo("Success", f"Successfully linked to Odoo machine '{odoo_machine_name}' (ID: {machine_id_in_odoo})."))

        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Odoo Link Failed", f"An error occurred:\n{e}"))
            self.log_operation("Error", f"Odoo Link Failed: {e}", conn_id)
        finally:
            self.after(0, self.update_status, "Ready")

    def sync_to_odoo(self):
        """Starts the Odoo synchronization process in a background thread."""
        threading.Thread(target=self._sync_to_odoo_thread, daemon=True).start()

    def _sync_to_odoo_thread(self):
        """The actual logic for syncing data to Odoo, runs in a background thread."""
        self.log_operation("Odoo Sync", "Starting synchronization process...")
        
        url = self.odoo_details.get("url")
        db = self.odoo_details.get("db")
        username = self.odoo_details.get("username")
        password = self.odoo_details.get("password")

        if not all([url, db, username, password]):
            self.log_operation("Error", "Odoo connection details not configured.")
            self.after(0, lambda: messagebox.showerror("Error", "Odoo connection details not configured."))
            return

        try:
            # 1. Authenticate with Odoo
            common = xmlrpc.client.ServerProxy(f'{url}/xmlrpc/2/common')
            uid = common.authenticate(db, username, password, {})
            if not uid:
                self.log_operation("Error", "Odoo authentication failed.")
                self.after(0, lambda: messagebox.showerror("Odoo Auth Failed", "Could not authenticate with Odoo."))
                return
            models = xmlrpc.client.ServerProxy(f'{url}/xmlrpc/2/object')
            self.log_operation("Odoo Sync", "Authentication successful.")

            batch_size = int(self.settings.get("batch_size", 1000))
            
            # 2. Sync Users
            unsynced_users = db_execute("SELECT u.*, m.odoo_machine_id FROM users u JOIN zkteco_machines m ON u.connection_id = m.id WHERE u.synched_time IS NULL", fetch='all')
            if not unsynced_users:
                self.log_operation("Odoo Sync", "No new users to synchronize.")
            else:
                total_users = len(unsynced_users)
                self.log_operation("Odoo Sync", f"Found {total_users} new user(s) to synchronize.")
                synced_count = 0
                for user in unsynced_users:
                    if not user['odoo_machine_id']:
                        self.log_operation("Odoo Sync", f"Skipping user {user['user_id']} because their machine is not linked to Odoo.", user['connection_id'])
                        continue
                    
                    user_payload = {
                        'uid': user['uid'],
                        'user_id': user['user_id'],
                        'name': user['name'],
                        'machine_id': user['odoo_machine_id'],
                        'raw_data': ''
                    }
                    try:
                        models.execute_kw(db, uid, password, 'azk.machine.proxy.users', 'create', [user_payload])
                        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        db_execute("UPDATE users SET synched_time = ? WHERE id = ?", (now_str, user['id']))
                        synced_count += 1
                        if synced_count % 50 == 0:
                            self.log_operation("Odoo Sync", f"Uploaded {synced_count} of {total_users} users so far.")
                    except Exception as e:
                        self.log_operation("Error", f"Failed to create user {user['user_id']} in Odoo: {e}", user['connection_id'])
                
                self.log_operation("Odoo Sync", f"Successfully synchronized {synced_count} user(s).")

            # 3. Sync Attendance
            unsynced_attendance = db_execute("SELECT a.*, m.odoo_machine_id, m.machine_timezone FROM attendance a JOIN zkteco_machines m ON a.connection_id = m.id WHERE a.synched_time IS NULL", fetch='all')
            if not unsynced_attendance:
                self.log_operation("Odoo Sync", "No new attendance records to synchronize.")
            else:
                total_att = len(unsynced_attendance)
                self.log_operation("Odoo Sync", f"Found {total_att} new attendance record(s) to sync.")
                
                payload_batch = []
                record_id_batch = []
                synced_count = 0

                last_user_attendance = {} #get the last attendance for each user in the batch to avoid duplicates

                for att in unsynced_attendance:
                    if not att['odoo_machine_id']:
                        self.log_operation("Odoo Sync", f"Skipping attendance for user {att['user_id']} because their machine is not linked.", att['connection_id'])
                        continue
                    
                    # The timestamp in the database is the naive time from the device.
                    # Odoo expects UTC. We must convert it.
                    timestamp_to_send = att['timestamp']
                    if att['machine_timezone']:
                        try:
                            # Assume the stored naive time is in the machine's local timezone
                            local_tz = pytz.timezone(att['machine_timezone'])
                            naive_timestamp = datetime.strptime(att['timestamp'], "%Y-%m-%d %H:%M:%S")
                            local_dt = local_tz.localize(naive_timestamp, is_dst=None) # is_dst=None handles ambiguous times
                            utc_dt = local_dt.astimezone(pytz.utc)
                            timestamp_to_send = utc_dt.strftime("%Y-%m-%d %H:%M:%S")
                        except Exception as e:
                            self.log_operation("Error", f"Could not process timezone '{att['machine_timezone']}' for record. Sending naive time. Error: {e}", att['connection_id'])
                    
                    if att['user_id'] not in last_user_attendance:
                        records = models.execute_kw(db, uid, password, 'azk.machine.proxy.attendance', 'search_read',
                                                        [[('user_id', '=', att['user_id']), ('machine_id', '=', att['odoo_machine_id'])]],
                                                        {
                                                            'fields': ['timestamp'],
                                                            'order': 'timestamp desc',
                                                            'limit': 1
                                                        }
                                    )
                        
                        if records:
                            last_user_attendance[att['user_id']] = records[0]['timestamp']
                        else:
                            last_user_attendance[att['user_id']] = None

                    if not last_user_attendance[att['user_id']] or last_user_attendance[att['user_id']] < timestamp_to_send:
                        payload_batch.append({
                            'user_id': att['user_id'],
                            'timestamp': timestamp_to_send,
                            'machine_id': att['odoo_machine_id'],
                            'att_id': att['att_id'],
                        })
                    record_id_batch.append(att['id'])

                    if len(record_id_batch) >= batch_size:
                        try:
                            if payload_batch:
                                models.execute_kw(db, uid, password, 'azk.machine.proxy.attendance', 'create', [payload_batch])
                            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            placeholders = ','.join('?' for _ in record_id_batch)
                            db_execute(f"UPDATE attendance SET synched_time = ? WHERE id IN ({placeholders})", [now_str] + record_id_batch)
                            synced_count += len(payload_batch)
                            self.log_operation("Odoo Sync", f"Uploaded {synced_count} of {total_att} attendance records so far.")
                            payload_batch = []
                            record_id_batch = []
                        except Exception as e:
                             self.log_operation("Error", f"Failed to create attendance batch in Odoo: {e}")

                # Process any remaining records in the last batch
                if record_id_batch:
                    try:
                        if payload_batch:
                            models.execute_kw(db, uid, password, 'azk.machine.proxy.attendance', 'create', [payload_batch])
                        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        placeholders = ','.join('?' for _ in record_id_batch)
                        db_execute(f"UPDATE attendance SET synched_time = ? WHERE id IN ({placeholders})", [now_str] + record_id_batch)
                        synced_count += len(payload_batch)
                    except Exception as e:
                        self.log_operation("Error", f"Failed to create final attendance batch in Odoo: {e}")

                self.log_operation("Odoo Sync", f"Successfully synchronized {synced_count} attendance record(s).")


            self.log_operation("Success", "Synchronization process finished.")
            self.after(0, self.refresh_all_data)
            self.after(0, lambda: messagebox.showinfo("Success", "Odoo synchronization complete. Check logs for details."))

        except Exception as e:
            self.log_operation("Error", f"Odoo Sync Failed: {e}")
            error_msg = f"An error occurred:\n{e}"
            self.after(0, lambda: messagebox.showerror("Odoo Sync Failed", error_msg))
        finally:
            self.after(0, self.update_status, "Ready")


    # --- Scheduler Methods ---
    def start_scheduler(self):
        if not self.scheduler_running.is_set():
            self.scheduler_running.set()
            self.scheduler_thread = threading.Thread(target=self.scheduler_loop, daemon=True)
            self.scheduler_thread.start()
            self.start_scheduler_btn.config(state="disabled")
            self.stop_scheduler_btn.config(state="normal")
            self.scheduler_status_var.set("Running")
            self.log_operation("Scheduler", "Scheduler started.")
            self.update_scheduler_next_run()

    def stop_scheduler(self):
        if self.scheduler_running.is_set():
            self.scheduler_running.clear()
            self.start_scheduler_btn.config(state="normal")
            self.stop_scheduler_btn.config(state="disabled")
            self.scheduler_status_var.set("Stopped")
            self.next_run_var.set("N/A")
            self.log_operation("Scheduler", "Scheduler stopped.")

    def update_scheduler_next_run(self):
        if self.scheduler_running.is_set():
            delay_minutes = int(self.settings.get("scheduler_delay", 10))
            next_run = datetime.now() + timedelta(minutes=delay_minutes)
            self.next_run_var.set(next_run.strftime("%Y-%m-%d %H:%M:%S"))
        else:
             self.next_run_var.set("N/A")

    def scheduler_loop(self):
        """Runs in a background thread, triggering data fetches."""
        # Initial wait before the first run
        time.sleep(5) 
        
        while self.scheduler_running.is_set():
            self.after(0, self._execute_scheduled_run)
            
            delay_seconds = int(self.settings.get("scheduler_delay", 10)) * 60
            # Wait for the interval, but check for stop event every second
            # This makes the shutdown more responsive
            for _ in range(delay_seconds):
                if not self.scheduler_running.is_set():
                    return
                time.sleep(1)
            
    def _execute_scheduled_run(self):
        """Scheduled by the loop to run in the main UI thread."""
        if not self.scheduler_running.is_set():
            return

        run_start_time = datetime.now()
        self.last_run_var.set(run_start_time.strftime("%Y-%m-%d %H:%M:%S"))
        self.log_operation("Scheduler", "Starting scheduled run...")
        
        all_machines = db_execute("SELECT * FROM zkteco_machines", fetch='all')

        for machine in all_machines:
            # Run each machine fetch in its own thread to avoid one failed machine blocking others
            threading.Thread(target=self._fetch_data_for_machine, args=(machine['id'], dict(machine), False), daemon=True).start()

        # After fetching, run the sync to Odoo
        threading.Thread(target=self._sync_to_odoo_thread, daemon=True).start()

        self._clean_old_logs()
        self._clean_old_attendance_from_settings()

        self.log_operation("Scheduler", "Scheduled run finished.")
        self.after(0, self._update_last_run_logs_table, run_start_time)
        self.after(1000, self.update_scheduler_next_run) # Update next run time after a second
        
    def _clean_old_logs(self):
        days = int(self.settings.get("clean_logs_days", 30))
        cutoff_date = datetime.now() - timedelta(days=days)
        cutoff_date_str = cutoff_date.strftime("%Y-%m-%d %H:%M:%S")
        db_execute("DELETE FROM logs WHERE timestamp < ?", (cutoff_date_str,))
        self.log_operation("Scheduler", f"Cleaned logs older than {days} days.")
    
    def _clean_old_attendance_from_settings(self):
        days = int(self.settings.get("delete_attendance_days", 180))
        cutoff_date = datetime.now() - timedelta(days=days)
        cutoff_date_str = cutoff_date.strftime("%Y-%m-%d %H:%M:%S")
        db_execute("DELETE FROM attendance WHERE timestamp < ?", (cutoff_date_str,))
        self.log_operation("Scheduler", f"Cleaned attendance records older than {days} days.")

if __name__ == "__main__":
    app = App()
    app.mainloop()
