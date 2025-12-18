import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import csv
import os
from datetime import datetime
import hashlib
import json
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

class LoginSystem:
    def __init__(self):
        self.users_file = "users.json"
        self.current_user = None
        self.load_users()
    
    def load_users(self):
        if os.path.exists(self.users_file):
            try:
                with open(self.users_file, 'r') as f:
                    self.users = json.load(f)
            except:
                self.users = {}
    
    def save_users(self):
        with open(self.users_file, 'w') as f:
            json.dump(self.users, f)
    
    def hash_password(self, password):
        return hashlib.sha256(password.encode()).hexdigest()
    
    def register(self, username, password, nama, role="guru"):
        if username in self.users:
            return False, "Username sudah terdaftar"
        
        self.users[username] = {
            "password": self.hash_password(password),
            "role": role,
            "nama": nama
        }
        self.save_users()
        return True, "Registrasi berhasil"
    
    def login(self, username, password):
        if username not in self.users:
            return False, "Username tidak ditemukan"
        
        if self.users[username]["password"] != self.hash_password(password):
            return False, "Password salah"
        
        self.current_user = {
            "username": username,
            "role": self.users[username]["role"],
            "nama": self.users[username]["nama"]
        }
        return True, "Login berhasil"
    
    def logout(self):
        self.current_user = None
    
    def is_logged_in(self):
        return self.current_user is not None
    
    def is_admin(self):
        return self.is_logged_in() and self.current_user["role"] == "admin"
    
    def is_guru(self):
        return self.is_logged_in() and self.current_user["role"] == "guru"
class SiRapor:
    def __init__(self, Layar_Utama):
        self.Layar_Utama = Layar_Utama
        self.Layar_Utama.title("SiRapor") #JUDUL WINDOW
        self.Layar_Utama.geometry("1200x800") #UKURAN UTAMA
        self.Layar_Utama.iconbitmap("Unesa.ico") #ICON ATAS KIRI
        self.icon_app = tk.PhotoImage(file="logo.png")
        self.Layar_Utama.iconphoto(True, self.icon_app)
        self.Layar_Utama.maxsize(1200, 800) 
        self.login_system = LoginSystem()
        self.data_siswa = []
        self.csv_filename = "data_raport_siswa.csv"
    
        self.style = ttk.Style()
        self.style.configure('TFrame', background="#ffffff") #WARNA POPUP 
        self.style.configure('TLabel', background="#fffcfc", font=('Segoe UI', 10)) #WARNA LABEL
        self.style.configure('TButton', font=('Segoe UI', 12))
        self.style.configure('Header.TLabel', font=('Segoe UI', 30, 'bold'), foreground="#000000")
        
        self.show_login_screen()
        
    def show_login_screen(self):
        for widget in self.Layar_Utama.winfo_children():
            widget.destroy()
        
        bg_image = tk.PhotoImage(file="unesaBg.gif")
        bg_label = tk.Label(self.Layar_Utama, image=bg_image)
        bg_label.place(x=0, y=0, relwidth=1, relheight=1)
        bg_label.image = bg_image 

        login_frame = ttk.Frame(self.Layar_Utama, padding=(150, 40)) # frame popup login
        login_frame.place(relx=0.5, rely=0.5, anchor='center')
        
        img_bawaan = tk.PhotoImage(file="logo.png")
        self.logo_img = img_bawaan.subsample(2, 2) 
        logo_label = ttk.Label(login_frame, image=self.logo_img)
        logo_label.grid(row=0, column=0, columnspan=2, pady=(0, 15))
        
        ttk.Label(login_frame, text="SiRapor", font=('Verda', 28, 'bold')).grid(row=1, column=0, columnspan=2, pady=(0,10))
        
        ttk.Label(login_frame, text="Username:",font=('Segoe UI', 15)).grid(row=2, column=0, sticky='e', pady=5, padx=10)
        self.login_username = ttk.Entry(login_frame, width=25) 
        self.login_username.grid(row=2, column=1, sticky='w', pady=5, padx=0)
        
        ttk.Label(login_frame, text="Password:", font=('Segoe UI', 15)).grid(row=3, column=0, sticky='e', pady=5, padx=10)
        self.login_password = ttk.Entry(login_frame, width=25, show="*")
        self.login_password.grid(row=3, column=1, sticky='w', pady=5, padx=0)
        
        button_container = ttk.Frame(login_frame)
        button_container.grid(row=4, column=0, columnspan=2, pady=(25, 0))

        btn_login = ttk.Button(button_container, text="Login", command=self.do_login, width=12)
        btn_login.pack(side='left', padx=5)
        
        btn_register = ttk.Button(button_container, text="Register", command=self.Halaman_Register, width=12)
        btn_register.pack(side='left', padx=5)
        
        self.login_password.bind('<Return>', lambda e: self.do_login())
    
    def Halaman_Register(self):
        register_window = tk.Toplevel(self.Layar_Utama)
        register_window.title("Registrasi")
        register_window.geometry("500x350")
        register_window.transient(self.Layar_Utama)
        register_window.grab_set()
        
        ttk.Label(register_window, text="REGISTRASI", font=('Segoe UI', 14, 'bold')).pack(pady=10)
        
        form_frame = ttk.Frame(register_window, padding=10)
        form_frame.pack(fill='both', expand=True, anchor='center')
        
        ttk.Label(form_frame, text="Nama Lengkap:").grid(row=0, column=0, sticky='w', pady=5)
        reg_nama = ttk.Entry(form_frame, width=20)
        reg_nama.grid(row=0, column=1, pady=5, padx=5,)
        
        ttk.Label(form_frame, text="Username:").grid(row=1, column=0, sticky='w', pady=5)
        reg_username = ttk.Entry(form_frame, width=20)
        reg_username.grid(row=1, column=1, pady=5, padx=5)
        
        ttk.Label(form_frame, text="Password:").grid(row=2, column=0, sticky='w', pady=5)
        reg_password = ttk.Entry(form_frame, width=20, show="*")
        reg_password.grid(row=2, column=1, pady=5, padx=5)
        
        ttk.Label(form_frame, text="Konfirmasi Password:").grid(row=3, column=0, sticky='w', pady=5)
        reg_confirm = ttk.Entry(form_frame, width=20, show="*")
        reg_confirm.grid(row=3, column=1, pady=5, padx=5)
        
        def register():
            if reg_password.get() != reg_confirm.get():
                messagebox.showerror("Error", "Password tidak cocok!")
                return
            
            success, message = self.login_system.register(
                reg_username.get(), 
                reg_password.get(), 
                reg_nama.get()
            )
            
            if success:
                messagebox.showinfo("Sukses", message)
                register_window.destroy()
            else:
                messagebox.showerror("Error", message)
        
        ttk.Button(form_frame, text="Register", command=register).grid(row=4, column=0, columnspan=2, pady=10)
    
    def do_login(self):
        username = self.login_username.get()
        password = self.login_password.get()
        
        success, message = self.login_system.login(username, password)
        
        if success:
            messagebox.showinfo("Sukses", f"Login berhasil!\nSelamat datang {self.login_system.current_user['nama']}")
            self.Applikasi_Utama()
        else:
            messagebox.showerror("Error", message)
    
    def Applikasi_Utama(self):
        for widget in self.Layar_Utama.winfo_children():
            widget.destroy()
        
        bg_image = tk.PhotoImage(file="unesaBg.gif")
        bg_label = tk.Label(self.Layar_Utama, image=bg_image)
        bg_label.place(x=0, y=0, relwidth=1, relheight=1)
        bg_label.image = bg_image 
        
        self.notebook = ttk.Notebook(self.Layar_Utama)
        self.notebook.pack(fill='both', expand=True, padx=40, pady=40) 
        
        self.frame_input = ttk.Frame(self.notebook)
        self.frame_leaderboard = ttk.Frame(self.notebook)
        self.frame_pencariansiswa = ttk.Frame(self.notebook)
        self.frame_export = ttk.Frame(self.notebook)
        self.frame_user = ttk.Frame(self.notebook)
        
        self.notebook.add(self.frame_input, text='Input Nilai')
        self.notebook.add(self.frame_leaderboard, text='Leaderboard Kelas')
        self.notebook.add(self.frame_pencariansiswa, text='Pencarian Siswa')
        self.notebook.add(self.frame_export, text='Export/Import')
        self.notebook.add(self.frame_user, text='Pengguna')
        
     
        self.TAB_INPUT_NILAI()
        self.TAB_LEADERBOARD()
        self.TAB_PENCARIAN_SISWA()
        self.setup_export_import_tab()
        self.setup_user_tab()
        self.load_data()
    
    def setup_user_tab(self):
        header_label = ttk.Label(self.frame_user, text="INFORMASI PENGGUNA", style='Header.TLabel')
        header_label.pack(pady=10)
        
        user_info_frame = ttk.Frame(self.frame_user)
        user_info_frame.pack(pady=10, padx=20, fill='x')
        
        user_info = self.login_system.current_user
        ttk.Label(user_info_frame, text=f"Nama: {user_info['nama']}", font=('Segoe UI', 11)).pack(anchor='w', pady=5)
        ttk.Label(user_info_frame, text=f"Username: {user_info['username']}", font=('Segoe UI', 11)).pack(anchor='w', pady=5)
        ttk.Label(user_info_frame, text=f"Role: {user_info['role']}", font=('Segoe UI', 11)).pack(anchor='w', pady=5)
        
        button_frame = ttk.Frame(self.frame_user)
        button_frame.pack(pady=20)
        
        ttk.Button(button_frame, text="Logout", command=self.logout).pack(side='left', padx=5)
        
        if self.login_system.is_admin():
            ttk.Button(button_frame, text="Kelola Pengguna", command=self.show_user_management).pack(side='left', padx=5)
    
    def show_user_management(self):
        management_window = tk.Toplevel(self.Layar_Utama)
        management_window.title("Kelola Pengguna")
        management_window.geometry("500x600")
        management_window.iconbitmap("Unesa.ico") #ICON ATAS KIRI
        
        ttk.Label(management_window, text="DAFTAR PENGGUNA", font=('Segoe UI', 14, 'bold')).pack(pady=10)
        
        columns = ('Username', 'Nama', 'Role')
        tree = ttk.Treeview(management_window, columns=columns, show='headings', height=15)
        
        tree.heading('Username', text='Username')
        tree.heading('Nama', text='Nama Lengkap')
        tree.heading('Role', text='Role')
        
        tree.column('Username', width=150)
        tree.column('Nama', width=200)
        tree.column('Role', width=100)
        
        for username, user_data in self.login_system.users.items():
            tree.insert('', 'end', values=(
                username,
                user_data['nama'],
                user_data['role']
            ))
        
        tree.pack(padx=10, pady=10, fill='both', expand=True)
        
        def delete_user():
            selected = tree.selection()
            if not selected:
                messagebox.showwarning("Peringatan", "Pilih pengguna yang akan dihapus")
                return
            
            username = tree.item(selected[0])['values'][0]
            if username == self.login_system.current_user['username']:
                messagebox.showerror("Error", "Tidak dapat menghapus akun sendiri!")
                return
            
            if messagebox.askyesno("Konfirmasi", f"Hapus pengguna {username}?"):
                del self.login_system.users[username]
                self.login_system.save_users()
                tree.delete(selected[0])
                messagebox.showinfo("Sukses", "Pengguna berhasil dihapus")
        
        ttk.Button(management_window, text="Hapus Pengguna", command=delete_user).pack(pady=10)
    
    def logout(self):
        self.login_system.logout()
        self.show_login_screen()
    
    def TAB_INPUT_NILAI(self):
        header_label = ttk.Label(self.frame_input, text="MASUKAN NILAI RAPORT SISWA", style='Header.TLabel')
        header_label.pack(pady=10)
        
        input_frame = ttk.Frame(self.frame_input)
        input_frame.pack(pady=10, padx=20, fill='x')
        
        ttk.Label(input_frame, text="Nama Siswa:").grid(row=0, column=0, sticky='w', pady=5)
        self.entry_nama = ttk.Entry(input_frame, width=30, font=('Segoe UI', 10))
        self.entry_nama.grid(row=0, column=1, padx=10, pady=5, sticky='w')
        
        ttk.Label(input_frame, text="Kelas:").grid(row=0, column=2, sticky='w', pady=5, padx=(20,0))
        self.entry_kelas = ttk.Entry(input_frame, width=15, font=('Segoe UI', 10))
        self.entry_kelas.grid(row=0, column=3, padx=10, pady=5, sticky='w')
        
        self.mata_pelajaran = [
            "Matematika", "Bahasa Indonesia", "Bahasa Inggris", "Fisika", 
            "Kimia", "Biologi", "Sejarah", "Geografi", "Ekonomi", 
            "Sosiologi", "Seni Budaya", "PJOK", "Pendidikan Agama", "PKN"
        ]
        
        self.entries_nilai = {}
        
        for i, mapel in enumerate(self.mata_pelajaran):
            ttk.Label(input_frame, text=f"{mapel}:").grid(row=i+1, column=0, sticky='w', pady=2)
            entry = ttk.Entry(input_frame, width=10, font=('Segoe UI', 10))
            entry.grid(row=i+1, column=1, padx=10, pady=2, sticky='w')
            self.entries_nilai[mapel] = entry
        
        button_frame = ttk.Frame(self.frame_input)
        button_frame.pack(pady=20)
        
        ttk.Button(button_frame, text="Hitung Rata-rata", command=self.hitung_rata_rata).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Simpan Data", command=self.simpan_data).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Reset Form", command=self.reset_form).pack(side='left', padx=5)
        
        self.result_frame = ttk.Frame(self.frame_input)
        self.result_frame.pack(pady=10, padx=20, fill='x')
        
        self.result_label = ttk.Label(self.result_frame, text="", font=('Segoe UI', 11, 'bold'), foreground='#2c3e50')
        self.result_label.pack()
        
    def TAB_LEADERBOARD(self):
        header_label = ttk.Label(self.frame_leaderboard, text="LEADERBOARD NILAI SISWA PER KELAS", style='Header.TLabel')
        header_label.pack(pady=10)
        
        control_frame = ttk.Frame(self.frame_leaderboard)
        control_frame.pack(pady=10)
        
        ttk.Label(control_frame, text="Filter Kelas:").pack(side='left', padx=5)
        self.kelas_filter = ttk.Combobox(control_frame, values=["Semua"] + self.get_unique_kelas(), width=15)
        self.kelas_filter.set("Semua")
        self.kelas_filter.pack(side='left', padx=5)
        self.kelas_filter.bind('<<ComboboxSelected>>', self.refresh_leaderboard)
        
        ttk.Button(control_frame, text="Refresh Leaderboard", command=self.refresh_leaderboard).pack(side='left', padx=5)
        
        if self.login_system.is_guru():
            ttk.Button(control_frame, text="Hapus Semua Data", command=self.hapus_semua_data).pack(side='left', padx=5)
        
        stats_frame = ttk.Frame(self.frame_leaderboard)
        stats_frame.pack(pady=5, padx=20, fill='x')
        self.stats_label = ttk.Label(stats_frame, text="", font=('Segoe UI', 10, 'bold'), foreground='#2c3e50')
        self.stats_label.pack()
        
        columns = ('Rank', 'Nama', 'Kelas', 'Matematika', 'B. Indo', 'B. Inggris', 'Rata-rata', 'Status')
        
        self.tree = ttk.Treeview(self.frame_leaderboard, columns=columns, show='headings', height=20)
        
        self.tree.heading('Rank', text='Rank')
        self.tree.heading('Nama', text='Nama Siswa')
        self.tree.heading('Kelas', text='Kelas')
        self.tree.heading('Matematika', text='Matematika')
        self.tree.heading('B. Indo', text='B. Indo')
        self.tree.heading('B. Inggris', text='B. Ing')
        self.tree.heading('Rata-rata', text='Rata-rata')
        self.tree.heading('Status', text='Status')

        self.tree.column('Rank', width=50, anchor='center')
        self.tree.column('Nama', width=150)
        self.tree.column('Kelas', width=80, anchor='center')
        self.tree.column('Matematika', width=80, anchor='center')
        self.tree.column('B. Indo', width=80, anchor='center')
        self.tree.column('B. Inggris', width=80, anchor='center')
        self.tree.column('Rata-rata', width=80, anchor='center')
        self.tree.column('Status', width=120, anchor='center')

        scrollbar = ttk.Scrollbar(self.frame_leaderboard, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side='left', fill='both', expand=True, padx=10)
        scrollbar.pack(side='right', fill='y')

        self.refresh_leaderboard()
        
    def TAB_PENCARIAN_SISWA(self):
        header_label = ttk.Label(self.frame_pencariansiswa, text="PENCARIAN SISWA", style='Header.TLabel')
        header_label.pack(pady=10)
        
        search_frame = ttk.Frame(self.frame_pencariansiswa)
        search_frame.pack(pady=10, padx=20, fill='x')
        
        ttk.Label(search_frame, text="Cari Siswa:").grid(row=0, column=0, sticky='w', pady=5)
        self.search_entry = ttk.Entry(search_frame, width=30, font=('Segoe UI', 10))
        self.search_entry.grid(row=0, column=1, padx=10, pady=5, sticky='w')
        self.search_entry.bind('<KeyRelease>', self.cari_siswa)
        
        #ttk.Button(search_frame, text="Cari", command=self.cari_siswa).grid(row=0, column=2, padx=5)
        ttk.Button(search_frame, text="Reset", command=self.reset_search).grid(row=0, column=3, padx=5)
        
        columns = ('Nama', 'Kelas', 'Matematika', 'Bahasa Indonesia', 'Bahasa Inggris', 'Rata-rata', 'Status')
        
        self.search_tree = ttk.Treeview(self.frame_pencariansiswa, columns=columns, show='headings', height=20)
        
        self.search_tree.heading('Nama', text='Nama Siswa')
        self.search_tree.heading('Kelas', text='Kelas')
        self.search_tree.heading('Matematika', text='Matematika')
        self.search_tree.heading('Bahasa Indonesia', text='B. Indonesia')
        self.search_tree.heading('Bahasa Inggris', text='B. Inggris')
        self.search_tree.heading('Rata-rata', text='Rata-rata')
        self.search_tree.heading('Status', text='Status')
        
        self.search_tree.column('Nama', width=150)
        self.search_tree.column('Kelas', width=80)
        self.search_tree.column('Matematika', width=80)
        self.search_tree.column('Bahasa Indonesia', width=80)
        self.search_tree.column('Bahasa Inggris', width=80)
        self.search_tree.column('Rata-rata', width=80)
        self.search_tree.column('Status', width=120)

        scrollbar = ttk.Scrollbar(self.frame_pencariansiswa, orient='vertical', command=self.search_tree.yview)
        self.search_tree.configure(yscrollcommand=scrollbar.set)
        
        self.search_tree.pack(side='left', fill='both', expand=True, padx=10)
        scrollbar.pack(side='right', fill='y')
        
        self.search_context_menu = tk.Menu(self.frame_pencariansiswa, tearoff=0)
        self.search_context_menu.add_command(label="Edit Data", command=self.edit_selected_student)
        self.search_context_menu.add_command(label="Hapus Data", command=self.delete_selected_student)
        self.search_context_menu.add_command(label="Export PDF", command=self.export_selected_pdf)
        
        self.search_tree.bind("<Button-3>", self.show_search_context_menu)
        
        self.reset_search()
    
    def cari_siswa(self, event=None):
        query = self.search_entry.get().strip().lower()
        
        for item in self.search_tree.get_children():
            self.search_tree.delete(item)
        
        if not query:
            self.reset_search()
            return
        
        results = []
        for siswa in self.data_siswa:
            if (query in siswa['nama'].lower() or 
                query in siswa['kelas'].lower()):
                results.append(siswa)
        
        for siswa in results:
            status, tag = self.hitung_status(siswa['rata_rata'])
            
            self.search_tree.insert('', 'end', values=(
                siswa['nama'],
                siswa['kelas'],
                siswa['nilai_mapel'].get('Matematika', '-'),
                siswa['nilai_mapel'].get('Bahasa Indonesia', '-'),
                siswa['nilai_mapel'].get('Bahasa Inggris', '-'),
                f"{siswa['rata_rata']:.2f}",
                status
            ), tags=(tag,))
        
        self.search_tree.tag_configure('sangat_baik', background="#4EFF4E")  
        self.search_tree.tag_configure('baik', background="#97ED97")        
        self.search_tree.tag_configure('cukup', background="#FFF7B1")       
        self.search_tree.tag_configure('bimbingan', background="#C75768")   
    
    def reset_search(self):
        self.search_entry.delete(0, tk.END)
        for item in self.search_tree.get_children():
            self.search_tree.delete(item)
        
        for siswa in self.data_siswa:
            status, tag = self.hitung_status(siswa['rata_rata'])
            
            self.search_tree.insert('', 'end', values=(
                siswa['nama'],
                siswa['kelas'],
                siswa['nilai_mapel'].get('Matematika', '-'),
                siswa['nilai_mapel'].get('Bahasa Indonesia', '-'),
                siswa['nilai_mapel'].get('Bahasa Inggris', '-'),
                f"{siswa['rata_rata']:.2f}",
                status
            ), tags=(tag,))
           
            self.search_tree.tag_configure('sangat_baik', background="#4EFF4E")  
            self.search_tree.tag_configure('baik', background="#97ED97")        
            self.search_tree.tag_configure('cukup', background="#FFF7B1")       
            self.search_tree.tag_configure('bimbingan', background="#C75768")   
    
    def show_search_context_menu(self, event):
        item = self.search_tree.identify_row(event.y)
        if item:
            self.search_tree.selection_set(item)
            self.search_context_menu.post(event.x_root, event.y_root)
    
    def edit_selected_student(self):
        if not self.login_system.is_guru():
            messagebox.showerror("Error", "Hanya guru yang dapat mengedit data!")
            return
            
        selected = self.search_tree.selection()
        if not selected:
            messagebox.showwarning("Peringatan", "Pilih siswa yang akan diedit")
            return
        
        student_name = str(self.search_tree.item(selected[0])['values'][0])
        student_class = str(self.search_tree.item(selected[0])['values'][1])
        
        student_data = None
        for siswa in self.data_siswa:
            if str(siswa['nama']) == student_name and str(siswa['kelas']) == student_class:
                student_data = siswa
                break
        if student_data:
            self.show_edit_dialog(student_data)
        else:
            messagebox.showerror("Error", "Data tidak ditemukan untuk diedit.")
    
    def delete_selected_student(self):
        if not self.login_system.is_guru():
            messagebox.showerror("Error", "Hanya guru yang dapat menghapus data!")
            return
            
        selected = self.search_tree.selection()
        if not selected:
            messagebox.showwarning("Peringatan", "Pilih siswa yang akan dihapus")
            return
        
        student_name = str(self.search_tree.item(selected[0])['values'][0])
        student_class = str(self.search_tree.item(selected[0])['values'][1])
        
        if messagebox.askyesno("Konfirmasi", f"Hapus data {student_name} ({student_class})?"):
            data_baru = []
            deleted = False
            
            for s in self.data_siswa:
                if str(s['nama']) == student_name and str(s['kelas']) == student_class:
                    deleted = True
                    continue
                data_baru.append(s)
            
            self.data_siswa = data_baru
            
            if deleted:
                self.save_data()
                self.refresh_leaderboard()
                self.reset_search()
                self.update_kelas_filter() 
                messagebox.showinfo("Sukses", "Data siswa berhasil dihapus secara permanen.")
            else:
                messagebox.showerror("Error", f"Gagal menghapus. Data tidak ditemukan di database.")
    
    def show_edit_dialog(self, student_data):
        edit_window = tk.Toplevel(self.Layar_Utama)
        edit_window.title("Edit Data Siswa")
        edit_window.geometry("400x600")
        edit_window.transient(self.Layar_Utama)
        edit_window.grab_set()
        
        ttk.Label(edit_window, text="EDIT DATA SISWA", font=('Segoe UI', 14, 'bold')).pack(pady=10)
        
        form_frame = ttk.Frame(edit_window, padding=10)
        form_frame.pack(fill='both', expand=True)
        
        ttk.Label(form_frame, text="Nama Siswa:").grid(row=0, column=0, sticky='w', pady=5)
        edit_nama = ttk.Entry(form_frame, width=30)
        edit_nama.insert(0, student_data['nama'])
        edit_nama.grid(row=0, column=1, pady=5, padx=5)
        
        ttk.Label(form_frame, text="Kelas:").grid(row=1, column=0, sticky='w', pady=5)
        edit_kelas = ttk.Entry(form_frame, width=15)
        edit_kelas.insert(0, student_data['kelas'])
        edit_kelas.grid(row=1, column=1, pady=5, padx=5)
        
        edit_entries = {}
        for i, mapel in enumerate(self.mata_pelajaran):
            ttk.Label(form_frame, text=f"{mapel}:").grid(row=i+2, column=0, sticky='w', pady=2)
            entry = ttk.Entry(form_frame, width=10)
            nilai = student_data['nilai_mapel'].get(mapel, '')
            if nilai:
                entry.insert(0, str(nilai))
            entry.grid(row=i+2, column=1, pady=2, padx=5)
            edit_entries[mapel] = entry
        
        def save_edit():
            student_data['nama'] = edit_nama.get()
            student_data['kelas'] = edit_kelas.get()
            
            total_nilai = 0
            jumlah_mapel = 0
            for mapel, entry in edit_entries.items():
                nilai_text = entry.get().strip()
                if nilai_text:
                    try:
                        nilai = float(nilai_text)
                        if nilai < 0 or nilai > 100:
                            messagebox.showerror("Error", f"Nilai {mapel} harus antara 0-100!")
                            return
                        student_data['nilai_mapel'][mapel] = nilai
                        total_nilai += nilai
                        jumlah_mapel += 1
                    except ValueError:
                        messagebox.showerror("Error", f"Nilai {mapel} harus berupa angka!")
                        return
            
            if jumlah_mapel > 0:
                student_data['rata_rata'] = total_nilai / jumlah_mapel
            
            self.save_data()
            self.refresh_leaderboard()
            self.reset_search()

            messagebox.showinfo("Sukses", "Data berhasil diupdate!")
            edit_window.destroy()
        
        ttk.Button(edit_window, text="Simpan Perubahan", command=save_edit).pack(pady=10)
    
    def export_selected_pdf(self):
        if not self.login_system.is_guru():
            messagebox.showerror("Error", "Hanya guru yang dapat mengekpor data!")
            return
        selected = self.search_tree.selection()
        if not selected:
            messagebox.showwarning("Peringatan", "Pilih siswa yang akan diexport ke PDF")
            return
        
        student_name = str(self.search_tree.item(selected[0])['values'][0])
        student_class = str(self.search_tree.item(selected[0])['values'][1])
        
        student_data = None
        for siswa in self.data_siswa:
            if str(siswa['nama']) == student_name and str(siswa['kelas']) == student_class:
                student_data = siswa
                break
        
        if student_data:
            self.export_single_pdf(student_data)
    
    def export_single_pdf(self, student_data):
        filename = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Simpan Laporan PDF",
            initialfile=f"Raport_{student_data['nama']}_{student_data['kelas']}.pdf"
        )
        
        if filename:
            try:
                doc = SimpleDocTemplate(filename, pagesize=A4)
                elements = []
                
                styles = getSampleStyleSheet()
                title_style = ParagraphStyle(
                    'CustomTitle',
                    parent=styles['Heading1'],
                    fontSize=16,
                    spaceAfter=30,
                    alignment=1
                )
                
                title = Paragraph(f"LAPORAN NILAI SISWA", title_style)
                elements.append(title)
                
                student_info = [
                    [f"Nama Siswa: {student_data['nama']}", f"Kelas: {student_data['kelas']}"],
                    [f"Rata-rata: {student_data['rata_rata']:.2f}", f"Status: {self.hitung_status(student_data['rata_rata'])[0]}"],
                    ["", f"Tanggal: {datetime.now().strftime('%d/%m/%Y')}"]
                ]
                
                student_table = Table(student_info, colWidths=[3*inch, 3*inch])
                student_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, -1), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black)
                ]))
                elements.append(student_table)
                elements.append(Spacer(1, 20))
                
                grades_data = [['Mata Pelajaran', 'Nilai']]
                for mapel in self.mata_pelajaran:
                    nilai = student_data['nilai_mapel'].get(mapel, '-')
                    if nilai != '-':
                        grades_data.append([mapel, str(nilai)])
                
                grades_table = Table(grades_data, colWidths=[4*inch, 2*inch])
                grades_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, -1), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black)
                ]))
                elements.append(grades_table)
                
                doc.build(elements)
                messagebox.showinfo("Sukses", f"Laporan PDF berhasil disimpan di {filename}")
                
            except Exception as e:
                messagebox.shorterror("Error", f"Gagal membuat PDF: {str(e)}")
    
    def setup_export_import_tab(self):
        header_label = ttk.Label(self.frame_export, text="EXPORT DAN IMPORT DATA", style='Header.TLabel')
        header_label.pack(pady=10)
        
        pdf_frame = ttk.LabelFrame(self.frame_export, text="Export Laporan PDF", padding=10)
        pdf_frame.pack(pady=10, padx=20, fill='x')
        
        ttk.Button(pdf_frame, text="Export Semua Data ke PDF", command=self.export_all_pdf).pack(pady=5)
        ttk.Button(pdf_frame, text="Export per Kelas ke PDF", command=self.export_class_pdf).pack(pady=5)
        
        csv_frame = ttk.LabelFrame(self.frame_export, text="Export/Import CSV", padding=10)
        csv_frame.pack(pady=10, padx=20, fill='x')
        
        ttk.Button(csv_frame, text="Import Data CSV", command=self.import_from_csv).pack(pady=5)
        ttk.Button(csv_frame, text="Export dengan Nama Custom", command=self.export_custom_csv).pack(pady=5)
        ttk.Button(csv_frame, text="Buka File CSV", command=self.buka_file_csv).pack(pady=5)
        
        file_section = ttk.LabelFrame(self.frame_export, text="Informasi File", padding=10)
        file_section.pack(pady=10, padx=20, fill='x')
        
        ttk.Label(file_section, text=f"File data otomatis tersimpan di: {self.csv_filename}").pack(pady=2)
        ttk.Label(file_section, text="Data secara otomatis disimpan setiap kali menambah data baru").pack(pady=2)
        
        preview_frame = ttk.Frame(self.frame_export)
        preview_frame.pack(pady=10, padx=20, fill='both', expand=True)
    
    def export_all_pdf(self):
        if not self.login_system.is_guru():
            messagebox.showerror("Error", "Hanya guru yang dapat mengexport data!")
            return
        
        if not self.data_siswa:
            messagebox.showinfo("Info", "Tidak ada data untuk diexport!")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Simpan Laporan PDF Semua Siswa",
            initialfile="Raport_Semua_Siswa.pdf"
        )
        
        if filename:
            try:
                doc = SimpleDocTemplate(filename, pagesize=A4)
                elements = []
                
                styles = getSampleStyleSheet()
                title_style = ParagraphStyle(
                    'CustomTitle',
                    parent=styles['Heading1'],
                    fontSize=16,
                    spaceAfter=30,
                    alignment=1
                )
                
                title = Paragraph("LAPORAN NILAI SEMUA SISWA", title_style)
                elements.append(title)
                
                total_siswa = len(self.data_siswa)
                nilai_tertinggi = max(siswa['rata_rata'] for siswa in self.data_siswa)
                nilai_terendah = min(siswa['rata_rata'] for siswa in self.data_siswa)
                rata_rata_kelas = sum(siswa['rata_rata'] for siswa in self.data_siswa) / total_siswa
                
                summary_data = [
                    ["Total Siswa", "Nilai Tertinggi", "Nilai Terendah", "Rata-rata Kelas"],
                    [str(total_siswa), f"{nilai_tertinggi:.2f}", f"{nilai_terendah:.2f}", f"{rata_rata_kelas:.2f}"]
                ]
                
                summary_table = Table(summary_data, colWidths=[1.5*inch, 1.5*inch, 1.5*inch, 1.5*inch])
                summary_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, -1), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, 1), colors.beige),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black)
                ]))
                elements.append(summary_table)
                elements.append(Spacer(1, 20))
                
                student_data = [['No', 'Nama', 'Kelas', 'Rata-rata', 'Status']]
                
                sorted_data = sorted(self.data_siswa, key=lambda x: x['rata_rata'], reverse=True)
                for i, siswa in enumerate(sorted_data, 1):
                    status, _ = self.hitung_status(siswa['rata_rata'])
                    student_data.append([
                        str(i),
                        siswa['nama'],
                        siswa['kelas'],
                        f"{siswa['rata_rata']:.2f}",
                        status
                    ])
                
                student_table = Table(student_data, colWidths=[0.5*inch, 2.5*inch, 1*inch, 1*inch, 1.5*inch])
                student_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, -1), 8),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey])
                ]))
                elements.append(student_table)
                
                doc.build(elements)
                messagebox.showinfo("Sukses", f"Laporan PDF berhasil disimpan di {filename}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Gagal membuat PDF: {str(e)}")
    
    def export_class_pdf(self):
        if not self.login_system.is_guru():
            messagebox.showerror("Error", "Hanya guru yang dapat mengexport data!")
            return
        if not self.data_siswa:
            messagebox.showinfo("Info", "Tidak ada data untuk diexport!")
            return
        
        class_window = tk.Toplevel(self.Layar_Utama)
        class_window.title("Pilih Kelas")
        class_window.geometry("300x150")
        class_window.transient(self.Layar_Utama)
        class_window.grab_set()
        
        ttk.Label(class_window, text="Pilih Kelas untuk Export:").pack(pady=10)
        
        kelas_var = tk.StringVar()
        kelas_combo = ttk.Combobox(class_window, textvariable=kelas_var, values=self.get_unique_kelas())
        kelas_combo.pack(pady=5)
        
        def export_selected_class():
            selected_class = kelas_var.get()
            if not selected_class:
                messagebox.showwarning("Peringatan", "Pilih kelas terlebih dahulu!")
                return
            
            filename = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                title=f"Simpan Laporan PDF Kelas {selected_class}",
                initialfile=f"Raport_Kelas_{selected_class}.pdf"
            )
            
            if filename:
                class_window.destroy()
                self.export_class_pdf_file(selected_class, filename)
        
        ttk.Button(class_window, text="Export PDF", command=export_selected_class).pack(pady=10)
    
    def export_class_pdf_file(self, kelas, filename):
        try:
            class_students = [s for s in self.data_siswa if s['kelas'] == kelas]
            
            if not class_students:
                messagebox.showinfo("Info", f"Tidak ada data untuk kelas {kelas}")
                return
            
            doc = SimpleDocTemplate(filename, pagesize=A4)
            elements = []
            
            styles = getSampleStyleSheet()
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=16,
                spaceAfter=30,
                alignment=1
            )
            
            title = Paragraph(f"LAPORAN NILAI KELAS {kelas}", title_style)
            elements.append(title)
            
            total_siswa = len(class_students)
            nilai_tertinggi = max(siswa['rata_rata'] for siswa in class_students)
            nilai_terendah = min(siswa['rata_rata'] for siswa in class_students)
            rata_rata_kelas = sum(siswa['rata_rata'] for siswa in class_students) / total_siswa
            
            summary_data = [
                ["Total Siswa", "Nilai Tertinggi", "Nilai Terendah", "Rata-rata Kelas"],
                [str(total_siswa), f"{nilai_tertinggi:.2f}", f"{nilai_terendah:.2f}", f"{rata_rata_kelas:.2f}"]
            ]
            
            summary_table = Table(summary_data, colWidths=[1.5*inch, 1.5*inch, 1.5*inch, 1.5*inch])
            summary_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, 1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ]))
            elements.append(summary_table)
            elements.append(Spacer(1, 20))
            
            student_data = [['No', 'Nama', 'Rata-rata', 'Status'] + self.mata_pelajaran[:3]]
            
            sorted_data = sorted(class_students, key=lambda x: x['rata_rata'], reverse=True)
            for i, siswa in enumerate(sorted_data, 1):
                status, _ = self.hitung_status(siswa['rata_rata'])
                row = [
                    str(i),
                    siswa['nama'],
                    f"{siswa['rata_rata']:.2f}",
                    status
                ]
                
                for mapel in self.mata_pelajaran[:3]:
                    nilai = siswa['nilai_mapel'].get(mapel, '-')
                    row.append(str(nilai) if nilai != '-' else '-')
                
                student_data.append(row)
            
            col_widths = [0.4*inch, 2*inch, 0.8*inch, 1.2*inch] + [0.8*inch] * 3
            student_table = Table(student_data, colWidths=col_widths)
            student_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 7),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey])
            ]))
            elements.append(student_table)
            
            doc.build(elements)
            messagebox.showinfo("Sukses", f"Laporan PDF kelas {kelas} berhasil disimpan di {filename}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Gagal membuat PDF: {str(e)}")

    def hitung_status(self, rata_rata):
        """Menentukan status berdasarkan nilai rata-rata"""
        if rata_rata >= 91:
            return "SANGAT BAIK", 'sangat_baik'
        elif rata_rata >= 88:
            return "BAIK", 'baik'
        elif rata_rata >= 84:
            return "CUKUP", 'cukup'
        else:
            return "PERLU BIMBINGAN", 'bimbingan'
        
    def hitung_rata_rata(self):
        nama = self.entry_nama.get().strip()
        kelas = self.entry_kelas.get().strip()
        
        if not nama:
            messagebox.showerror("Error", "Nama siswa harus diisi!")
            return
        
        if not kelas:
            messagebox.showerror("Error", "Kelas harus diisi!")
            return
        
        total_nilai = 0
        jumlah_mapel = 0
        nilai_mapel = {}
        
        #logika meenghitung
        for mapel, entry in self.entries_nilai.items():
            nilai_text = entry.get().strip()
            if nilai_text:
                try:
                    nilai = float(nilai_text)
                    if nilai < 0 or nilai > 100:
                        messagebox.showerror("Error", f"Nilai {mapel} harus antara 0-100!")
                        return
                    total_nilai += nilai
                    jumlah_mapel += 1
                    nilai_mapel[mapel] = nilai
                except ValueError:
                    messagebox.showerror("Error", f"Nilai {mapel} harus berupa angka!")
                    return
        
        if jumlah_mapel == 0:
            messagebox.showerror("Error", "Minimal satu mata pelajaran harus diisi!")
            return
        
        rata_rata = total_nilai / jumlah_mapel
        status, _ = self.hitung_status(rata_rata)
        
        result_text = f"Rata-rata nilai {nama} ({kelas}): {rata_rata:.2f} - Status: {status}"
        self.result_label.config(text=result_text)
        
        messagebox.showinfo("Sukses", f"Rata-rata nilai {nama} adalah {rata_rata:.2f}\nsilahkan klik tombol simpan data.")
        
        self.data_sementara = {
            'nama': nama,
            'kelas': kelas,
            'rata_rata': rata_rata,
            'nilai_mapel': nilai_mapel,
            'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        
    def simpan_data(self):
        if not hasattr(self, 'data_sementara'):
            messagebox.showerror("Error", "Hitung rata-rata terlebih dahulu!")
            return
        
        self.data_siswa.append(self.data_sementara)
        
        self.save_data()
        
        messagebox.showinfo("Sukses", f"Data {self.data_sementara['nama']} berhasil disimpan!")
        self.reset_form()
        self.refresh_leaderboard()
        
        self.update_kelas_filter()
        
    def reset_form(self):
        self.entry_nama.delete(0, tk.END)
        self.entry_kelas.delete(0, tk.END)
        for entry in self.entries_nilai.values():
            entry.delete(0, tk.END)
        self.result_label.config(text="")
        if hasattr(self, 'data_sementara'):
            del self.data_sementara
    
    def get_unique_kelas(self):
        kelas_list = list(set([siswa['kelas'] for siswa in self.data_siswa]))
        return sorted(kelas_list)
    
    def update_kelas_filter(self):
        unique_kelas = self.get_unique_kelas()
        self.kelas_filter['values'] = ["Semua"] + unique_kelas
        
    def refresh_leaderboard(self, event=None):
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        selected_kelas = self.kelas_filter.get()
        if selected_kelas == "Semua":
            filtered_data = self.data_siswa
            if filtered_data:
                total_siswa = len(filtered_data)
                nilai_tertinggi = max(siswa['rata_rata'] for siswa in filtered_data)
                nilai_terendah = min(siswa['rata_rata'] for siswa in filtered_data)
                rata_rata_kelas = sum(siswa['rata_rata'] for siswa in filtered_data) / total_siswa
                
                status_count = {'SANGAT BAIK': 0, 'BAIK': 0, 'CUKUP': 0, 'PERLU BIMBINGAN': 0}
                for siswa in filtered_data:
                    status, _ = self.hitung_status(siswa['rata_rata'])
                    status_count[status] += 1
                
                stats_text = (f"Total Siswa: {total_siswa} | Nilai Tertinggi: {nilai_tertinggi:.2f} | "
                            f"Nilai Terendah: {nilai_terendah:.2f} | Rata-rata: {rata_rata_kelas:.2f}\n"
                            f"Status: {status_count['SANGAT BAIK']} Sangat Baik, {status_count['BAIK']} Baik, "
                            f"{status_count['CUKUP']} Cukup, {status_count['PERLU BIMBINGAN']} Perlu Bimbingan")
            else:
                stats_text = "Belum ada data siswa"
        else:
            filtered_data = [siswa for siswa in self.data_siswa if siswa['kelas'] == selected_kelas]
            if filtered_data:
                total_siswa = len(filtered_data)
                nilai_tertinggi = max(siswa['rata_rata'] for siswa in filtered_data)
                nilai_terendah = min(siswa['rata_rata'] for siswa in filtered_data)
                rata_rata_kelas = sum(siswa['rata_rata'] for siswa in filtered_data) / total_siswa
                
                status_count = {'SANGAT BAIK': 0, 'BAIK': 0, 'CUKUP': 0, 'PERLU BIMBINGAN': 0}
                for siswa in filtered_data:
                    status, _ = self.hitung_status(siswa['rata_rata'])
                    status_count[status] += 1
                
                stats_text = (f"Kelas {selected_kelas} - Total Siswa: {total_siswa} | "
                            f"Nilai Tertinggi: {nilai_tertinggi:.2f} | Nilai Terendah: {nilai_terendah:.2f} | "
                            f"Rata-rata Kelas: {rata_rata_kelas:.2f}\n"
                            f"Status: {status_count['SANGAT BAIK']} Sangat Baik, {status_count['BAIK']} Baik, "
                            f"{status_count['CUKUP']} Cukup, {status_count['PERLU BIMBINGAN']} Perlu Bimbingan")
            else:
                stats_text = f"Tidak ada data untuk kelas {selected_kelas}"
        
        self.stats_label.config(text=stats_text)
        
        sorted_data = sorted(filtered_data, key=lambda x: x['rata_rata'], reverse=True)
        
        for rank, siswa in enumerate(sorted_data, 1):
            status, tag = self.hitung_status(siswa['rata_rata'])
            
            self.tree.insert('', 'end', values=(
                rank,
                siswa['nama'],
                siswa['kelas'],
                siswa['nilai_mapel'].get('Matematika', '-'),
                siswa['nilai_mapel'].get('Bahasa Indonesia', '-'),
                siswa['nilai_mapel'].get('Bahasa Inggris', '-'),
                f"{siswa['rata_rata']:.2f}",
                status
            ), tags=(tag,))
        
        self.tree.tag_configure('sangat_baik', background="#4EFF4E")  
        self.tree.tag_configure('baik', background="#97ED97")        
        self.tree.tag_configure('cukup', background="#FFF7B1")       
        self.tree.tag_configure('bimbingan', background="#C75768")   
        
    def hapus_semua_data(self):
        if not self.login_system.is_guru():
            messagebox.showerror("Error", "Hanya guru yang dapat menghapus data!")
            return
            
        if not self.data_siswa:
            messagebox.showinfo("Info", "Tidak ada data untuk dihapus!")
            return
        
        if messagebox.askyesno("Konfirmasi", "Apakah Anda yakin ingin menghapus semua data?"):
            self.data_siswa = []
            self.save_data()
            self.refresh_leaderboard()
            self.update_kelas_filter()
            messagebox.showinfo("Sukses", "Semua data berhasil dihapus!")
    
    def save_data(self):
        try:
            with open(self.csv_filename, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                
                header = ['Nama', 'Kelas', 'Rata_Rata', 'Timestamp'] + self.mata_pelajaran
                writer.writerow(header)
                
                for siswa in self.data_siswa:
                    row = [
                        siswa['nama'],
                        siswa['kelas'],
                        f"{siswa['rata_rata']:.2f}",
                        siswa['timestamp']
                    ]
                    
                    for mapel in self.mata_pelajaran:
                        row.append(siswa['nilai_mapel'].get(mapel, ''))
                    
                    writer.writerow(row)
                    
        except Exception as e:
            messagebox.showerror("Error", f"Gagal menyimpan data ke CSV: {str(e)}")
    
    def load_data(self):
        try:
            if os.path.exists(self.csv_filename):
                self.data_siswa = []
                with open(self.csv_filename, 'r', newline='', encoding='utf-8') as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        nilai_mapel = {}
                        for mapel in self.mata_pelajaran:
                            if row.get(mapel):
                                try:
                                    nilai_mapel[mapel] = float(row[mapel])
                                except ValueError:
                                    nilai_mapel[mapel] = 0
                        
                        siswa_data = {
                            'nama': row['Nama'],
                            'kelas': row['Kelas'],
                            'rata_rata': float(row['Rata_Rata']),
                            'timestamp': row['Timestamp'],
                            'nilai_mapel': nilai_mapel
                        }
                        self.data_siswa.append(siswa_data)
                if hasattr(self, 'kelas_filter'):
                    self.update_kelas_filter()
                if hasattr(self, 'tree'):
                    self.refresh_leaderboard()
                if hasattr(self, 'search_tree'):
                    self.reset_search()
        except Exception as e:
            messagebox.showerror("Error", f"Gagal memuat data dari CSV: {str(e)}")
    
    def import_from_csv(self):
        if not self.login_system.is_guru():
            messagebox.showerror("Error", "Hanya guru yang dapat mengimport data!")
            return
            
        filename = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            title="Pilih file CSV untuk diimport"
        )
        
        if filename:
            try:
                imported_data = []
                with open(filename, 'r', newline='', encoding='utf-8') as f:
                    reader = csv.DictReader(f)
                    
                    expected_columns = ['Nama', 'Kelas', 'Rata_Rata', 'Timestamp'] + self.mata_pelajaran
                    if not all(col in reader.fieldnames for col in expected_columns):
                        messagebox.showerror("Error", "Format file CSV tidak sesuai! Pastikan file memiliki kolom yang benar.")
                        return
                    
                    for row in reader:
                        nilai_mapel = {}
                        for mapel in self.mata_pelajaran:
                            if row.get(mapel) and row[mapel].strip():
                                try:
                                    nilai_mapel[mapel] = float(row[mapel])
                                except ValueError:
                                    nilai_mapel[mapel] = 0
                        
                        siswa_data = {
                            'nama': row['Nama'],
                            'kelas': row['Kelas'],
                            'rata_rata': float(row['Rata_Rata']),
                            'timestamp': row.get('Timestamp', datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
                            'nilai_mapel': nilai_mapel
                        }
                        imported_data.append(siswa_data)
                
                if imported_data:
                    choice = messagebox.askyesno("Import Data", 
                        f"Ditemukan {len(imported_data)} data siswa.\n"
                        f"Ya = Ganti semua data dengan data baru\n"
                        f"Tidak = Tambahkan data baru ke data existing")
                    
                    if choice:
                        self.data_siswa = imported_data
                    else:
                        self.data_siswa.extend(imported_data)
                    
                    self.save_data()
                    self.refresh_leaderboard()
                    self.update_kelas_filter()
                    
                    messagebox.showinfo("Sukses", f"Berhasil mengimport {len(imported_data)} data siswa!")
                else:
                    messagebox.showinfo("Info", "Tidak ada data yang dapat diimport dari file tersebut.")
                    
            except Exception as e:
                messagebox.showerror("Error", f"Gagal import data: {str(e)}")
    
    def export_custom_csv(self):
        if not self.login_system.is_guru():
            messagebox.showerror("Error", "Hanya guru yang dapat mengimport data!")
            return
        
        if not self.data_siswa:
            messagebox.showinfo("Info", "Tidak ada data untuk diexport!")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            title="Simpan file CSV sebagai"
        )
        
        if filename:
            try:
                with open(filename, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f)
                    
                    header = ['Rank', 'Nama', 'Kelas', 'Rata_Rata', 'Status'] + self.mata_pelajaran + ['Timestamp']
                    writer.writerow(header)
                    
                    sorted_data = sorted(self.data_siswa, key=lambda x: x['rata_rata'], reverse=True)
                    
                    for rank, siswa in enumerate(sorted_data, 1):
                        status, _ = self.hitung_status(siswa['rata_rata'])
                        
                        row = [
                            rank,
                            siswa['nama'],
                            siswa['kelas'],
                            f"{siswa['rata_rata']:.2f}",
                            status
                        ]
                        
                        for mapel in self.mata_pelajaran:
                            row.append(siswa['nilai_mapel'].get(mapel, ''))
                        
                        row.append(siswa['timestamp'])
                        writer.writerow(row)
                
                messagebox.showinfo("Sukses", f"Data berhasil diexport ke {filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Gagal export data: {str(e)}")
    
    def buka_file_csv(self):
        if not self.login_system.is_guru():
            messagebox.showerror("Error", "Hanya guru yang bisa membuka file CSV")
            return
        try:
            os.startfile(self.csv_filename)  # windows
        except:
            try:
                # non windows
                import subprocess
                subprocess.run(['open', self.csv_filename])  # macOS
            except:
                try:
                    subprocess.run(['xdg-open', self.csv_filename])  # Linux
                except:
                    messagebox.showinfo("Info", f"File CSV tersimpan di: {os.path.abspath(self.csv_filename)}")
    
def main():
    Layar_Utama = tk.Tk()
    app = SiRapor(Layar_Utama)
    Layar_Utama.mainloop()

if __name__ == "__main__":
    main()