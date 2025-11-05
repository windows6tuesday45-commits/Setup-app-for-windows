"""
Let's Install Drivers Setup Master — Win32Shell Edition
Updated: Remove Background button, Exit Program, full save settings, restart prompt
"""

import os
import sys
import json
import time
import threading
import subprocess
import ctypes
from ctypes import wintypes
import tkinter as tk
from tkinter import ttk, filedialog, colorchooser, messagebox
import winsound
import random
import shutil

# ------------------------ Configuration ------------------------
DESKTOP = os.path.join(os.path.expanduser("~"), "Desktop")
SETTINGS_PATH = os.path.join(DESKTOP, "driver_setup_settings_win32shell.json")
RANDOM_WALLPAPER_FOLDER = os.path.join(DESKTOP, "RandomWallpapers")
FULL_TITLE = "Let's Install Drivers Setup Master — Win32Shell Edition"
SPI_SETDESKWALLPAPER = 20

# ------------------------ Win32 Helpers ------------------------
def set_wallpaper(path):
    if os.path.isfile(path):
        ctypes.windll.user32.SystemParametersInfoW(SPI_SETDESKWALLPAPER, 0, path, 3)

def remove_wallpaper():
    ctypes.windll.user32.SystemParametersInfoW(SPI_SETDESKWALLPAPER, 0, "", 3)

def run_elevated(exe_path, args=""):
    if not os.path.isabs(exe_path):
        exe_path = os.path.abspath(exe_path)
    res = ctypes.windll.shell32.ShellExecuteW(None, "runas", exe_path, args, None, 1)
    return int(res) > 32

def play_wav_async(path):
    if os.path.isfile(path):
        winsound.PlaySound(path, winsound.SND_FILENAME | winsound.SND_ASYNC)

def do_restart(delay_seconds=5):
    try:
        subprocess.Popen(["shutdown", "/r", "/t", str(max(0, int(delay_seconds)))])
        return True
    except Exception:
        return False

# ---------------------- Settings ----------------------
DEFAULT_SETTINGS = {
    "background": "",
    "text_size": 12,
    "sound": "",
    "installer_list": [],
    "window_color": "#f0f0f0"
}

def load_settings():
    s = DEFAULT_SETTINGS.copy()
    try:
        if os.path.isfile(SETTINGS_PATH):
            with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
                s.update(json.load(f))
    except Exception:
        pass
    return s

def save_settings(settings):
    try:
        with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2)
    except:
        pass

# ---------------------- GUI ----------------------
class Win32ShellSetupApp(tk.Tk):
    def __init__(self):
        super().__init__()
        if os.name != "nt":
            messagebox.showerror("Platform","Windows only")
            self.destroy()
            return
        self.settings = load_settings()
        self.title(FULL_TITLE)
        self.geometry("820x680")
        self.resizable(False, False)
        self.installer_list = list(self.settings.get("installer_list",[]))
        self.text_size = tk.IntVar(value=self.settings.get("text_size",12))
        self.window_color = self.settings.get("window_color", "#f0f0f0")
        self.configure(bg=self.window_color)
        self.create_widgets()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def create_widgets(self):
        # ---------- Top toolbar ----------
        toolbar = tk.Frame(self, bg="#d0d0d0")
        toolbar.pack(side="top", fill="x", padx=2, pady=2)
        btn_specs = [
            ("Files", self.btn_files),
            ("Change Color", self.btn_change_color),
            ("Add Desktop Background", self.btn_add_background),
            ("New Desktop Background", self.btn_new_background),
            ("Remove Background", self.btn_remove_background),
            ("Check for Updates", self.btn_check_updates),
            ("What's New?", self.btn_whats_new),
            ("About", self.btn_about),
            ("Exit Program", self.btn_exit_program)
        ]
        for txt, cmd in btn_specs:
            b = tk.Button(toolbar, text=txt, command=cmd)
            b.pack(side="left", padx=4, pady=2)

        # ---------- Header ----------
        header = tk.Label(self, text=FULL_TITLE, font=("Segoe UI", 14, "bold"), wraplength=760, justify="center", bg=self.window_color)
        header.pack(pady=(4,6))

        # ---------- Installer List ----------
        installers_frame = tk.LabelFrame(self,text="Select Drivers / Installers (from PC or USB)", bg=self.window_color)
        installers_frame.pack(fill="x", padx=12, pady=8)

        self.installer_listbox=tk.Listbox(installers_frame,height=8)
        self.installer_listbox.pack(side="left",fill="both",expand=True,padx=(6,0),pady=6)
        for p in self.installer_list:
            self.installer_listbox.insert("end",p)

        installer_buttons=tk.Frame(installers_frame, bg=self.window_color)
        installer_buttons.pack(side="right",padx=6,pady=6,fill="y")
        tk.Button(installer_buttons,text="Add Installer...",command=self.add_installer).pack(fill="x",pady=2)
        tk.Button(installer_buttons,text="Remove Selected",command=self.remove_installer).pack(fill="x",pady=2)
        tk.Button(installer_buttons,text="Run Selected (elevated)",command=self.run_selected_installer).pack(fill="x",pady=2)

        # ---------- Progress ----------
        progress_frame=tk.LabelFrame(self,text="Installation Progress", bg=self.window_color)
        progress_frame.pack(fill="x", padx=12, pady=6)
        self.status_text = tk.StringVar(value="Ready")
        tk.Label(progress_frame,textvariable=self.status_text, bg=self.window_color).pack(anchor="w")
        self.total_progress=ttk.Progressbar(progress_frame,maximum=1000,length=780)
        self.total_progress.pack(pady=6)

        # ---------- Log ----------
        log_frame=tk.LabelFrame(self,text="Installation Log", bg=self.window_color)
        log_frame.pack(fill="both", padx=12, pady=6, expand=True)
        self.log_text=tk.Text(log_frame,height=10,wrap="word",state="disabled")
        self.log_text.pack(side="left",fill="both",expand=True,padx=(6,0),pady=6)
        scrollbar=tk.Scrollbar(log_frame,command=self.log_text.yview)
        scrollbar.pack(side="right",fill="y")
        self.log_text.configure(yscrollcommand=scrollbar.set)

        # ---------- Apply & Restart Button ----------
        btn_row=tk.Frame(self, bg=self.window_color)
        btn_row.pack(fill="x", padx=12, pady=(4,12))
        tk.Button(btn_row,text="Apply Settings & Restart",command=self.apply_settings_and_restart).pack(side="right")

        self.log("Win32Shell Setup ready.")
        self.apply_text_size()

    # ---------------- Toolbar button methods ----------------
    def btn_files(self):
        filedialog.askopenfilename(title="Browse Files")

    def btn_change_color(self):
        color=colorchooser.askcolor(title="Choose window background color")
        if color[1]:
            self.window_color=color[1]
            self.configure(bg=self.window_color)
            self.log(f"Window color changed: {self.window_color}")
            self.save_settings()

    def btn_add_background(self):
        p=filedialog.askopenfilename(title="Select Desktop Background",filetypes=[("Images","*.bmp;*.jpg;*.png;*.jpeg;*.gif")])
        if p:
            try:
                set_wallpaper(p)
                self.log(f"Desktop background set to: {p}")
                self.settings["background"]=p
                self.save_settings()
            except: pass

    def btn_new_background(self):
        wallpaper=""
        if os.path.isdir(RANDOM_WALLPAPER_FOLDER):
            imgs=[os.path.join(RANDOM_WALLPAPER_FOLDER,f) for f in os.listdir(RANDOM_WALLPAPER_FOLDER)
                  if f.lower().endswith((".bmp",".jpg",".png",".jpeg",".gif"))]
            if imgs: wallpaper=random.choice(imgs)
        if wallpaper:
            try:
                set_wallpaper(wallpaper)
                self.log(f"Random desktop background set: {wallpaper}")
                self.settings["background"]=wallpaper
                self.save_settings()
            except: pass

    def btn_remove_background(self):
        try:
            remove_wallpaper()
            self.log("Desktop background removed.")
            self.settings["background"]=""
            self.save_settings()
        except: pass

    def btn_check_updates(self):
        messagebox.showinfo("Check for Updates","No updates available (placeholder).")

    def btn_whats_new(self):
        messagebox.showinfo("What's New","Version 1.1.0\n- Added Remove Background\n- Exit Program button\n- Full settings save\n- Restart prompt after installation")

    def btn_about(self):
        messagebox.showinfo("About",FULL_TITLE+"\n© 2025 Setup Master Team")

    def btn_exit_program(self):
        self.on_close()

    # ---------------- Log & Installer ----------------
    def log(self,msg):
        ts=time.strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.configure(state="normal")
        self.log_text.insert("end",f"[{ts}] {msg}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def add_installer(self):
        p=filedialog.askopenfilename(title="Select driver/installer", filetypes=[("Executables","*.exe;*.msi"),("All files","*.*")])
        if p:
            p=os.path.abspath(p)
            self.installer_list.append(p)
            self.installer_listbox.insert("end",p)
            self.save_settings()
            self.log(f"Added installer: {p}")

    def remove_installer(self):
        sel=self.installer_listbox.curselection()
        if not sel: return
        idx=sel[0]
        p=self.installer_list[idx]
        self.installer_listbox.delete(idx)
        del self.installer_list[idx]
        self.save_settings()
        self.log(f"Removed installer: {p}")

    def run_selected_installer(self):
        sel=self.installer_listbox.curselection()
        if not sel: return
        idx=sel[0]
        exe=self.installer_list[idx]
        if not os.path.isfile(exe): 
            messagebox.showerror("Error","File not found")
            return
        self.log(f"Launching installer elevated: {exe}")
        ok=run_elevated(exe)
        if ok: self.log("Installer launched successfully.")
        else: self.log("Failed to launch installer.")

    # ---------------- Settings & Restart ----------------
    def apply_settings_and_restart(self):
        self.save_settings()
        self.apply_text_size()
        messagebox.showinfo("Restart","Setup completed! Please restart your computer now.")
        if messagebox.askyesno("Restart","Restart system now?"):
            self.log("User requested restart.")
            do_restart(5)

    def apply_text_size(self):
        size=self.text_size.get()
        for w in self.winfo_children():
            if isinstance(w,tk.Label):
                try: w.configure(font=("Segoe UI",size))
                except: pass

    def save_settings(self):
        self.settings["text_size"]=self.text_size.get()
        self.settings["installer_list"]=self.installer_list
        self.settings["window_color"]=self.window_color
        save_settings(self.settings)

    def on_close(self):
        self.save_settings()
        self.destroy()

# ------------------ Main ------------------
def main():
    app=Win32ShellSetupApp()
    app.mainloop()

if __name__=="__main__":
    main()
