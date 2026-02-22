"""
╔══════════════════════════════════════════════╗
║         ASSETS COLLECTOR — ELITE EDITION     ║
║                                              ║
║  pip install psutil wmi pillow qrcode        ║
╚══════════════════════════════════════════════╝
"""

import tkinter as tk
import psutil                                           # pip install psutil
import wmi                                              # pip install wmi
from PIL import Image, ImageDraw, ImageFilter, ImageTk  # pip install pillow
import qrcode                                           # pip install qrcode
import socket, json, os, sys, math, threading, subprocess, ctypes

# ══════════════════════════════════════════════
# COLORS  — Deep Obsidian + Electric Accent
# ══════════════════════════════════════════════
C = {
    "bg":        "#04080F",
    "card":      "#090F1C",
    "card2":     "#0C1426",
    "code":      "#070D1A",
    "cyan":      "#00D4FF",
    "orange":    "#FF8C42",
    "magenta":   "#E040FB",
    "green":     "#00F5A0",
    "red":       "#FF3D5A",
    "text":      "#C8E6F5",
    "dim":       "#3A6080",
    "border":    "#0F2035",
    "header":    "#020508",
    "yellow":    "#FFD700",
    "gold":      "#C9A84C",
    "accent":    "#0066AA",
    # scrollbar dark-blue palette
    "sb_trough": "#030710",
    "sb_thumb":  "#071830",
    "sb_hover":  "#0A2848",
    "sb_line":   "#0B3060",
}

def _app_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

SAVE_FILE = os.path.join(_app_dir(), "pc_info.json")


# ══════════════════════════════════════════════
# GATHER SYSTEM INFO
# ══════════════════════════════════════════════
def gather_info():
    import pythoncom
    pythoncom.CoInitialize()
    data = {}
    W = wmi.WMI()

    # Whoami
    data["whoami"] = os.environ.get("USERNAME") or os.environ.get("USER") or "N/A"

    # Hostname
    data["hostname"] = socket.gethostname()

    # Local IP
    data["local_ip"] = "N/A"
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.settimeout(1)
        s.connect(("8.8.8.8", 80))
        data["local_ip"] = s.getsockname()[0]
        s.close()
    except Exception:
        try:
            data["local_ip"] = socket.gethostbyname(data["hostname"])
        except Exception:
            pass

    # Serial Number
    data["serial"] = "N/A"
    try:
        for b in W.Win32_BIOS():
            v = (b.SerialNumber or "").strip()
            data["serial"] = v if v else "N/A"
            break
    except Exception:
        pass

    # OS
    data["os"] = "N/A"
    try:
        for o in W.Win32_OperatingSystem():
            data["os"] = f"{o.Caption.strip()} ({o.OSArchitecture})"
            break
    except Exception:
        import platform
        data["os"] = platform.platform()

    # CPU
    data["cpu"] = "N/A"
    try:
        for p in W.Win32_Processor():
            data["cpu"] = p.Name.strip()
            break
    except Exception:
        pass

    # RAM
    data["ram"] = "N/A"
    try:
        data["ram"] = f"{round(psutil.virtual_memory().total / 1024**3)} GB"
    except Exception:
        pass

    # GPU
    data["gpu"] = "N/A"
    try:
        gpus = [(g.Name or "").strip() for g in W.Win32_VideoController()
                if (g.Name or "").strip()]
        data["gpu"] = " / ".join(gpus) if gpus else "N/A"
    except Exception:
        pass

    # Motherboard
    data["motherboard"] = "N/A"
    try:
        for b in W.Win32_BaseBoard():
            mfr = (b.Manufacturer or "").strip()
            prd = (b.Product or "").strip()
            data["motherboard"] = f"{mfr} {prd}".strip()
            break
    except Exception:
        pass

    # Disks — detection with USB Flash support
    NVME_KEYS = ["NVME", "NVM EXPRESS", "M.2 NVME", "NVME SSD"]
    USB_KEYS  = ["USB", "FLASH DRIVE", "FLASH DISK", "REMOVABLE",
                 "DATATRAVELER", "JETFLASH", "CRUZER", "ULTRA FIT",
                 "ULTRA USB", "VENGEANCE USB", "PENDRIVE", "PEN DRIVE",
                 "USB DISK", "USB DRIVE", "TRANSCEND", "TOSHIBA USB",
                 "KINGSTON USB", "SANDISK USB"]
    SSD_KEYS  = ["SSD", "SOLID STATE", "SOLID-STATE",
                 "EVO", "PRO ", "MX", "CRUCIAL", "KINGSTON",
                 "SANDISK", "WD GREEN", "WD BLUE SSD", "WD BLACK SSD",
                 "SAMSUNG 8", "SAMSUNG 9", "INTEL SSD", "MICRON",
                 "TOSHIBA SSD", "A400", "A2000", "BX500", "MX500",
                 "SA400", "ADATA SSD", "PATRIOT"]
    data["disks"] = []
    try:
        for idx, disk in enumerate(W.Win32_DiskDrive()):
            model      = (disk.Model      or "").strip()
            caption    = (disk.Caption    or "").strip()
            iface      = (disk.InterfaceType or "").strip().upper()
            media_type = (disk.MediaType  or "").lower()
            combo      = (model + " " + caption).upper()

            if any(k in combo for k in NVME_KEYS):
                dtype = "NVMe/M.2"
            elif iface == "USB":
                # Interface is USB — definitely a USB device
                # Distinguish Flash Drive vs external HDD by size & name
                try:
                    size_b = int(disk.Size or 0)
                except Exception:
                    size_b = 0
                if any(k in combo for k in USB_KEYS):
                    dtype = "USB Flash"
                elif size_b > 0 and size_b <= 256 * 1024**3:
                    # ≤256 GB on USB → likely flash drive
                    dtype = "USB Flash"
                else:
                    dtype = "USB HDD"
            elif any(k in combo for k in USB_KEYS):
                dtype = "USB Flash"
            elif any(k in combo for k in SSD_KEYS):
                dtype = "SSD"
            else:
                if "solid" in media_type or "flash" in media_type:
                    dtype = "SSD"
                elif "external" in media_type:
                    dtype = "USB HDD"
                elif "hard" in media_type or "fixed" in media_type:
                    dtype = "HDD"
                else:
                    try:
                        sp = int(disk.SpindleSpeed or -1)
                        dtype = "HDD" if sp > 0 else ("SSD" if sp == 0 else "HDD")
                    except Exception:
                        dtype = "HDD"
            try:
                size_gb = f"{round(int(disk.Size) / 1024**3):.0f} GB"
            except Exception:
                size_gb = "? GB"
            data["disks"].append({
                "idx":   idx,
                "model": model or caption,
                "size":  size_gb,
                "type":  dtype,
            })
    except Exception:
        pass

    pythoncom.CoUninitialize()
    return data


# ══════════════════════════════════════════════
# CUSTOM SCROLLBAR  (Canvas-based, dark-blue)
# ══════════════════════════════════════════════
class FancyScrollbar(tk.Canvas):
    WIDTH = 6
    PAD   = 1

    def __init__(self, parent, command, **kw):
        super().__init__(parent,
                         width=self.WIDTH,
                         bg=C["sb_trough"],
                         highlightthickness=0, bd=0, **kw)
        self._cmd      = command
        self._thumb_y0 = 0.0
        self._thumb_y1 = 1.0
        self._drag_y   = None
        self._hover    = False

        self.bind("<Configure>",        self._redraw)
        self.bind("<ButtonPress-1>",    self._on_press)
        self.bind("<B1-Motion>",        self._on_drag)
        self.bind("<ButtonRelease-1>",  self._on_release)
        self.bind("<Enter>",            self._on_enter)
        self.bind("<Leave>",            self._on_leave)
        self.bind("<MouseWheel>",       self._on_wheel)

    def set(self, lo, hi):
        self._thumb_y0 = float(lo)
        self._thumb_y1 = float(hi)
        self._redraw()

    def _redraw(self, *_):
        self.delete("all")
        H = self.winfo_height()
        if H < 2:
            return
        W = self.WIDTH
        self.create_line(W//2, 6, W//2, H-6,
                         fill=C["sb_line"], width=1)
        ty0 = int(self._thumb_y0 * H) + self.PAD
        ty1 = int(self._thumb_y1 * H) - self.PAD
        ty1 = max(ty1, ty0 + 14)
        col = C["sb_hover"] if self._hover else C["sb_thumb"]
        x0 = self.PAD
        x1 = W - self.PAD
        self.create_rectangle(x0, ty0, x1, ty1,
                              fill=col, outline=C["cyan"],
                              width=1)
        self._ty0 = ty0
        self._ty1 = ty1

    def _on_enter(self, e):  self._hover = True;  self._redraw()
    def _on_leave(self, e):  self._hover = False; self._redraw()

    def _on_press(self, e):
        H = self.winfo_height()
        if H == 0: return
        frac = e.y / H
        if hasattr(self, '_ty0') and self._ty0 <= e.y <= self._ty1:
            self._drag_y = e.y - self._ty0
        else:
            span = self._thumb_y1 - self._thumb_y0
            new  = max(0.0, min(1.0 - span, frac - span/2))
            self._cmd("moveto", new)

    def _on_drag(self, e):
        if self._drag_y is None: return
        H = self.winfo_height()
        if H == 0: return
        new  = (e.y - self._drag_y) / H
        span = self._thumb_y1 - self._thumb_y0
        new  = max(0.0, min(1.0 - span, new))
        self._cmd("moveto", new)

    def _on_release(self, e): self._drag_y = None

    def _on_wheel(self, e):
        self._cmd("scroll", int(-1*(e.delta/120)), "units")


# ══════════════════════════════════════════════
# SCROLLABLE FRAME
# ══════════════════════════════════════════════
class ScrollFrame(tk.Frame):
    def __init__(self, parent, **kw):
        super().__init__(parent, bg=C["bg"], **kw)
        self._cv = tk.Canvas(self, bg=C["bg"], highlightthickness=0, bd=0)
        self._sb = FancyScrollbar(self, command=self._cv.yview)
        self._cv.configure(yscrollcommand=self._sb.set)
        self._sb.pack(side="right", fill="y", padx=(0, 3), pady=6)
        self._cv.pack(side="left", fill="both", expand=True)
        self.inner = tk.Frame(self._cv, bg=C["bg"])
        self._id = self._cv.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>",
                        lambda e: self._cv.configure(scrollregion=self._cv.bbox("all")))
        self._cv.bind("<Configure>",
                      lambda e: self._cv.itemconfig(self._id, width=e.width))
        self.bind_all("<MouseWheel>",
                      lambda e: self._cv.yview_scroll(int(-1*(e.delta/120)), "units"))


# ══════════════════════════════════════════════
# AURORA BACKGROUND ANIMATION
# ══════════════════════════════════════════════
class AuroraBg:
    BLOBS = [
        (0.00, 0.20, 0.00028, 0.00016, 300, (0, 180, 255),   0.00),
        (1.00, 0.55, 0.00022, 0.00018, 280, (0, 140, 220),   1.57),
        (0.48, 0.00, 0.00018, 0.00014, 260, (0, 100, 180),   0.80),
        (0.02, 0.85, 0.00024, 0.00016, 310, (160, 20, 200),  2.00),
        (0.98, 0.40, 0.00020, 0.00013, 240, (100, 10, 160),  3.14),
        (0.50, 1.00, 0.00015, 0.00011, 320, (0, 40, 120),    0.50),
        (0.85, 0.10, 0.00020, 0.00015, 260, (0, 160, 240),   0.40),
        (0.15, 0.90, 0.00022, 0.00014, 240, (0, 200, 160),   1.80),
        (0.60, 0.30, 0.00018, 0.00012, 200, (180, 60, 255),  1.20),
    ]

    def __init__(self, parent, w, h):
        self.cv = tk.Canvas(parent, width=w, height=h,
                            highlightthickness=0, bd=0, bg=C["bg"])
        self.cv.place(x=0, y=0)
        self.cv.lower("all")
        self._w = w; self._h = h; self._t = 0; self._img = None
        self._tick()

    def _tick(self):
        try:
            w, h, t = self._w, self._h, self._t
            base = Image.new("RGB", (w, h), (4, 8, 15))
            lyr  = Image.new("RGB", (w, h), (0, 0, 0))
            drw  = ImageDraw.Draw(lyr)
            for xp, yp, xs, ys, r, col, ap in self.BLOBS:
                cx = int((xp + math.sin(t*xs*math.pi*2)*0.25)*w)
                cy = int((yp + math.cos(t*ys*math.pi*2)*0.20)*h)
                br = 0.45 + 0.45*math.sin(t*0.0004 + ap)
                for i in range(7):
                    s  = int(br*(i+1)/7*85)
                    sz = int(r*(1-i*0.13))
                    c2 = tuple(min(255, col[k]*s//255) for k in range(3))
                    drw.ellipse([cx-sz, cy-sz, cx+sz, cy+sz], fill=c2)
            lyr = lyr.filter(ImageFilter.GaussianBlur(radius=50))
            img = Image.blend(base, lyr, alpha=0.70)
            self._img = ImageTk.PhotoImage(img)
            self.cv.delete("all")
            self.cv.create_image(0, 0, anchor="nw", image=self._img)
            self.cv.lower("all")
            self._t += 1
        except Exception:
            pass
        self.cv.after(90, self._tick)


# ══════════════════════════════════════════════
# MAIN APPLICATION
# ══════════════════════════════════════════════
class App(tk.Tk):
    W, H = 600, 720

    def __init__(self):
        super().__init__()
        self.title("Assets Collector v3.0")
        self.resizable(False, False)
        self.configure(bg=C["bg"])
        self.overrideredirect(True)

        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass

        self._info  = None
        self._saved = {"assign_to": "", "location": "", "brand": "",
                       "note": "", "server_path": ""}
        try:
            if os.path.exists(SAVE_FILE):
                with open(SAVE_FILE, "r", encoding="utf-8") as f:
                    self._saved = json.load(f)
        except Exception:
            pass

        self._center()
        self._build_skeleton()

        # drag
        self._dx = self._dy = 0
        for w in (self._hbar, self._htitle):
            w.bind("<ButtonPress-1>", lambda e: setattr(self, '_dx', e.x_root - self.winfo_x())
                   or setattr(self, '_dy', e.y_root - self.winfo_y()))
            w.bind("<B1-Motion>",
                   lambda e: self.geometry(f"+{e.x_root-self._dx}+{e.y_root-self._dy}"))

        threading.Thread(target=self._load, daemon=True).start()

    def _center(self):
        sw = self.winfo_screenwidth(); sh = self.winfo_screenheight()
        self.geometry(f"{self.W}x{self.H}+{(sw-self.W)//2}+{(sh-self.H)//2}")

    # ══════════════════════════════════════════
    # SKELETON
    # ══════════════════════════════════════════
    def _build_skeleton(self):
        AuroraBg(self, self.W, self.H)

        # ── Outer border frame ─────────────────
        border = tk.Frame(self, bg=C["cyan"], padx=1, pady=1)
        border.place(x=0, y=0, width=self.W, height=self.H)

        inner_wrap = tk.Frame(border, bg=C["bg"])
        inner_wrap.pack(fill="both", expand=True)

        # ── Header ────────────────────────────
        self._hbar = tk.Frame(inner_wrap, bg=C["header"], height=62)
        self._hbar.pack(fill="x")
        self._hbar.pack_propagate(False)

        # top accent line
        tk.Frame(self._hbar, bg=C["cyan"], height=1).pack(side="top", fill="x")
        # bottom accent line with gold tint
        tk.Frame(self._hbar, bg=C["gold"], height=1).pack(side="bottom", fill="x")

        row = tk.Frame(self._hbar, bg=C["header"])
        row.pack(fill="both", expand=True, padx=4)

        # Icon + title group
        title_frame = tk.Frame(row, bg=C["header"])
        title_frame.pack(side="left", padx=(10, 0), pady=8)

        tk.Label(title_frame, text="◈", font=("Consolas", 18, "bold"),
                 fg=C["cyan"], bg=C["header"]).pack(side="left", padx=(0, 6))

        title_col = tk.Frame(title_frame, bg=C["header"])
        title_col.pack(side="left")

        self._htitle = tk.Label(title_col,
                                text="ASSETS  COLLECTOR",
                                font=("Consolas", 15, "bold"),
                                fg=C["cyan"], bg=C["header"])
        self._htitle.pack(anchor="w")

        tk.Label(title_col, text="v3.0  │  System Information Tool",
                 font=("Consolas", 8),
                 fg=C["dim"], bg=C["header"]).pack(anchor="w")

        # Window controls
        ctrl_frame = tk.Frame(row, bg=C["header"])
        ctrl_frame.pack(side="right", padx=10)

        for txt, fg, cmd in [("✕", C["red"],   self.destroy),
                              ("─", C["gold"],  self.iconify),
                              ("↺", C["green"], self._refresh)]:
            btn = tk.Label(ctrl_frame, text=txt,
                           font=("Consolas", 13, "bold"),
                           fg=fg, bg=C["card2"],
                           cursor="hand2", padx=12, pady=5,
                           width=2)
            btn.pack(side="right", padx=2)
            btn.bind("<Button-1>", lambda e, c=cmd: c())
            btn.bind("<Enter>",    lambda e, b=btn, f=fg: b.configure(bg=C["border"], fg=f))
            btn.bind("<Leave>",    lambda e, b=btn, f=fg: b.configure(bg=C["card2"], fg=f))

        # ── Body ─────────────────────────────
        self._sf   = ScrollFrame(inner_wrap)
        self._sf.pack(fill="both", expand=True)
        self._body = self._sf.inner

        # Loading state
        load_frame = tk.Frame(self._body, bg=C["bg"])
        load_frame.pack(pady=100)
        self._lbl_load = tk.Label(load_frame,
                                  text="⏳   Loading system information ...",
                                  font=("Consolas", 13),
                                  fg=C["dim"], bg=C["bg"])
        self._lbl_load.pack()
        tk.Label(load_frame, text="Please wait while hardware data is collected",
                 font=("Consolas", 9),
                 fg=C["border"], bg=C["bg"]).pack(pady=(4, 0))

    # ══════════════════════════════════════════
    # REFRESH
    # ══════════════════════════════════════════
    def _refresh(self):
        """Destroy all body widgets and reload system info from scratch."""
        # Clear the scroll canvas content
        for w in self._body.winfo_children():
            w.destroy()
        # Show loading indicator
        load_frame = tk.Frame(self._body, bg=C["bg"])
        load_frame.pack(pady=100)
        self._lbl_load = tk.Label(load_frame,
                                  text="↺   Refreshing system information ...",
                                  font=("Consolas", 13),
                                  fg=C["green"], bg=C["bg"])
        self._lbl_load.pack()
        tk.Label(load_frame, text="Please wait while hardware data is re-collected",
                 font=("Consolas", 9),
                 fg=C["border"], bg=C["bg"]).pack(pady=(4, 0))
        # Reset QR reference
        self._qr_tk = None
        threading.Thread(target=self._load, daemon=True).start()

    # ══════════════════════════════════════════
    # LOAD
    # ══════════════════════════════════════════
    def _load(self):
        self._info = gather_info()
        self.after(0, self._populate)

    # ══════════════════════════════════════════
    # POPULATE
    # ══════════════════════════════════════════
    def _populate(self):
        # Destroy loading frame (parent of _lbl_load)
        try:
            self._lbl_load.master.destroy()
        except Exception:
            pass
        d = self._info

        # ── Top summary strip ──────────────────
        strip = tk.Frame(self._body, bg=C["card"], height=44)
        strip.pack(fill="x", padx=0, pady=(0, 4))
        strip.pack_propagate(False)
        tk.Frame(strip, bg=C["cyan"], width=3).pack(side="left", fill="y")
        tk.Label(strip,
                 text=f"  HOST:  {d['hostname']}   │   IP:  {d['local_ip']}",
                 font=("Consolas", 9), fg=C["cyan"], bg=C["card"],
                 anchor="w").pack(fill="x", padx=10, pady=12)

        # ── QR Code Panel ─────────────────────
        self._sec("◈   DEVICE   QR CODE", C["cyan"])
        qr_panel = tk.Frame(self._body, bg=C["code"])
        qr_panel.pack(fill="x", padx=20, pady=(0, 4))
        tk.Frame(qr_panel, bg=C["cyan"], width=2).pack(side="left", fill="y")

        qr_inner = tk.Frame(qr_panel, bg=C["code"])
        qr_inner.pack(fill="both", expand=True, padx=14, pady=14)

        # Left: QR image placeholder (generated on save)
        self._qr_lbl = tk.Label(qr_inner, bg=C["code"],
                                text="QR will be generated\nupon saving",
                                font=("Consolas", 9), fg=C["dim"],
                                width=20, height=10, justify="center")
        self._qr_lbl.pack(side="left", anchor="n")

        # Right: QR info text
        qr_text_f = tk.Frame(qr_inner, bg=C["code"])
        qr_text_f.pack(side="left", fill="both", expand=True, padx=(16, 0))

        tk.Label(qr_text_f, text="Scan to view device info",
                 font=("Consolas", 10, "bold"),
                 fg=C["cyan"], bg=C["code"], anchor="w").pack(anchor="w")
        tk.Label(qr_text_f, text=" ", bg=C["code"]).pack()

        for lbl, val in [
            ("Hostname", d["hostname"]),
            ("IP",       d["local_ip"]),
            ("Serial",   d["serial"]),
            ("CPU",      d["cpu"][:36] + ("…" if len(d["cpu"]) > 36 else "")),
            ("RAM",      d["ram"]),
        ]:
            row_f = tk.Frame(qr_text_f, bg=C["code"])
            row_f.pack(anchor="w", pady=1)
            tk.Label(row_f, text=f"{lbl:<10}", font=("Consolas", 9),
                     fg=C["dim"], bg=C["code"]).pack(side="left")
            tk.Label(row_f, text=val, font=("Consolas", 9, "bold"),
                     fg=C["text"], bg=C["code"]).pack(side="left")

        # QR generated on save — see _save_local / _save_server

        # ── System Info ───────────────────────
        self._sec("◈   SYSTEM   INFORMATION", C["cyan"])
        self._row("Whoami",      d["whoami"],      C["yellow"])
        self._row("Hostname",    d["hostname"],    C["cyan"])
        self._row("IP Address",  d["local_ip"],    C["green"])
        self._row("Serial No.",  d["serial"],      C["green"])
        self._row("OS",          d["os"],          C["text"])
        self._row("CPU",         d["cpu"],         C["orange"])
        self._row("RAM",         d["ram"],         C["cyan"])
        self._row("GPU",         d["gpu"],         C["magenta"])
        self._row("Motherboard", d["motherboard"], C["dim"])

        # ── Storage ───────────────────────────
        self._sec("◈   STORAGE   DEVICES", C["orange"])
        if d["disks"]:
            for disk in d["disks"]:
                self._disk_row(disk)
        else:
            self._warn("No storage devices detected")

        # ── Assignment ────────────────────────
        self._sec("◈   ASSIGNMENT   DETAILS", C["gold"])

        req_hint = tk.Label(self._body,
                            text="  ★  Fields marked with  *  are required before saving",
                            font=("Consolas", 9), fg=C["red"], bg=C["bg"], anchor="w")
        req_hint.pack(fill="x", padx=20, pady=(0, 4))

        self._use_assign_name = tk.BooleanVar(value=False)
        self._e_assign     = self._field_with_check(
                                "Assign To",  "Full name ...", "",
                                self._use_assign_name, required=True)
        self._e_department = self._field("Department", "IT / HR / Finance ...", "", required=True)
        self._e_location   = self._field("Location",   "Office / Floor ...", "", required=True)
        self._e_brand      = self._field("Brand",      "Dell / HP / ...",    "", required=True)

        # Note field
        nf = tk.Frame(self._body, bg=C["bg"])
        nf.pack(fill="x", padx=20, pady=(4, 2))
        tk.Label(nf, text="Notes", font=("Consolas", 10),
                 fg=C["dim"], bg=C["bg"], width=12, anchor="nw",
                 pady=8).pack(side="left", anchor="n")
        note_wrap = tk.Frame(nf, bg=C["border"], padx=1, pady=1)
        note_wrap.pack(side="left", fill="x", expand=True)
        self._e_note = tk.Text(note_wrap, height=3, font=("Consolas", 10),
                               bg=C["card2"], fg=C["text"],
                               insertbackground=C["cyan"],
                               relief="flat", bd=0, padx=10, pady=8,
                               wrap="word", selectbackground=C["accent"])
        self._e_note.pack(fill="both")
        # Notes always start empty on open

        # ── Server Path ───────────────────────
        self._sec("◈   NETWORK   / SERVER", C["cyan"])
        hint = tk.Label(self._body,
                        text="  Shared folder path on the server  (e.g.  \\\\SERVER\\Reports)",
                        font=("Consolas", 9), fg=C["dim"], bg=C["bg"], anchor="w")
        hint.pack(fill="x", padx=20, pady=(0, 6))
        self._e_server = self._field("Server Path", r"\\SERVER\Reports",
                                     self._saved.get("server_path",""))

        # ── Divider ───────────────────────────
        div = tk.Frame(self._body, bg=C["bg"])
        div.pack(fill="x", padx=20, pady=(18, 6))
        tk.Frame(div, bg=C["border"], height=1).pack(fill="x")

        # ── Action Buttons ────────────────────
        bf = tk.Frame(self._body, bg=C["bg"])
        bf.pack(fill="x", padx=20, pady=(0, 4))
        for i in range(5):
            bf.grid_columnconfigure(i, weight=1)

        # btn: (label, border_color, text_color, func)
        btns = [
            ("💾  Local",   C["green"],    C["green"],   self._save_local),
            ("🌐  Server",  C["cyan"],     C["cyan"],    self._save_server),
            ("📋  Copy",    C["gold"],     C["gold"],    self._copy),
            ("📂  Folder",  C["orange"],   C["orange"],  self._folder),
            ("🖧  Share",   C["magenta"],  C["magenta"], self._open_share),
        ]
        for i, (txt, border, fg, cmd) in enumerate(btns):
            outer = tk.Frame(bf, bg=border, padx=1, pady=1)
            outer.grid(row=0, column=i, padx=2, sticky="ew")
            b = tk.Label(outer, text=txt,
                         font=("Consolas", 9, "bold"),
                         fg=fg, bg=C["card"],
                         cursor="hand2", padx=4, pady=10)
            b.pack(fill="both")
            b.bind("<Button-1>", lambda e, c=cmd: c())
            b.bind("<Enter>",    lambda e, lb=b, ac=border: lb.configure(bg=ac, fg=C["bg"]))
            b.bind("<Leave>",    lambda e, lb=b, of=fg: lb.configure(fg=of, bg=C["card"]))

        # Close button — dark steel color
        close_outer = tk.Frame(self._body, bg=C["dim"], padx=1, pady=1)
        close_outer.pack(fill="x", padx=20, pady=(6, 0))
        close_btn = tk.Label(close_outer, text="✕   Close Application",
                             font=("Consolas", 10, "bold"),
                             fg=C["text"], bg=C["code"],
                             cursor="hand2", pady=9)
        close_btn.pack(fill="both")
        close_btn.bind("<Button-1>", lambda e: self.destroy())
        close_btn.bind("<Enter>",    lambda e: close_btn.configure(bg=C["dim"], fg=C["bg"]))
        close_btn.bind("<Leave>",    lambda e: close_btn.configure(bg=C["code"], fg=C["text"]))

        # Status bar
        status_bar = tk.Frame(self._body, bg=C["header"], height=32)
        status_bar.pack(fill="x", pady=(10, 0))
        status_bar.pack_propagate(False)
        tk.Frame(status_bar, bg=C["gold"], height=1).pack(side="top", fill="x")
        self._status = tk.Label(status_bar, text="  Ready.",
                                font=("Consolas", 9),
                                fg=C["dim"], bg=C["header"], anchor="w")
        self._status.pack(fill="x", padx=14, pady=6)

    # ══════════════════════════════════════════
    # WIDGET HELPERS
    # ══════════════════════════════════════════
    def _sec(self, text, color):
        f = tk.Frame(self._body, bg=C["bg"])
        f.pack(fill="x", padx=20, pady=(18, 6))
        lbl = tk.Label(f, text=text, font=("Consolas", 11, "bold"),
                       fg=color, bg=C["bg"])
        lbl.pack(side="left")
        tk.Frame(f, bg=color, height=1).pack(
            side="left", fill="x", expand=True, padx=(10, 0), pady=6)

    def _copy_val(self, value):
        """Copy a single value to clipboard and flash status."""
        try:
            self.clipboard_clear()
            self.clipboard_append(value)
            self.update()
            self._status_msg(f"✓  Copied:  {value[:50]}", C["cyan"])
        except Exception:
            pass

    def _row(self, label, value, col):
        f = tk.Frame(self._body, bg=C["code"])
        f.pack(fill="x", padx=20, pady=1)
        # left accent bar
        tk.Frame(f, bg=C["border"], width=2).pack(side="left", fill="y")

        # Copy button on far right — always visible, subtle
        copy_btn = tk.Label(f, text="⎘", font=("Consolas", 11),
                            fg=C["dim"], bg=C["code"],
                            cursor="hand2", padx=8, pady=0)
        copy_btn.pack(side="right", fill="y")
        copy_btn.bind("<Button-1>", lambda e, v=value: self._copy_val(v))
        copy_btn.bind("<Enter>",    lambda e, b=copy_btn: b.configure(fg=C["cyan"], bg=C["card2"]))
        copy_btn.bind("<Leave>",    lambda e, b=copy_btn: b.configure(fg=C["dim"],  bg=C["code"]))

        inner = tk.Frame(f, bg=C["code"])
        inner.pack(fill="x", padx=12, pady=7)
        tk.Label(inner, text=label, font=("Consolas", 10),
                 fg=C["dim"], bg=C["code"],
                 width=13, anchor="w").pack(side="left")
        # separator
        tk.Label(inner, text="│", font=("Consolas", 10),
                 fg=C["border"], bg=C["code"]).pack(side="left", padx=(0, 8))
        tk.Label(inner, text=value, font=("Consolas", 11, "bold"),
                 fg=col, bg=C["code"],
                 anchor="w", wraplength=320, justify="left").pack(side="left")

    def _disk_row(self, d):
        type_style = {
            "NVMe/M.2": (C["cyan"],    "⚡"),
            "SSD":      (C["green"],   "◼"),
            "HDD":      (C["orange"],  "◎"),
            "USB Flash":(C["magenta"], "⬡"),
            "USB HDD":  (C["yellow"],  "⬡"),
        }
        tc, icon = type_style.get(d["type"], (C["dim"], "◎"))
        copy_text = f"{d['size']} [{d['type']}] {d['model']}"

        f = tk.Frame(self._body, bg=C["code"])
        f.pack(fill="x", padx=20, pady=1)
        tk.Frame(f, bg=tc, width=2).pack(side="left", fill="y")

        # Copy button on far right
        copy_btn = tk.Label(f, text="⎘", font=("Consolas", 11),
                            fg=C["dim"], bg=C["code"],
                            cursor="hand2", padx=8, pady=0)
        copy_btn.pack(side="right", fill="y")
        copy_btn.bind("<Button-1>", lambda e, v=copy_text: self._copy_val(v))
        copy_btn.bind("<Enter>",    lambda e, b=copy_btn: b.configure(fg=C["cyan"], bg=C["card2"]))
        copy_btn.bind("<Leave>",    lambda e, b=copy_btn: b.configure(fg=C["dim"],  bg=C["code"]))

        inner = tk.Frame(f, bg=C["code"])
        inner.pack(fill="x", padx=12, pady=9)

        tk.Label(inner, text=f"Disk  {d['idx']}", font=("Consolas", 10),
                 fg=C["dim"], bg=C["code"], width=13, anchor="w").pack(side="left")
        tk.Label(inner, text="│", font=("Consolas", 10),
                 fg=C["border"], bg=C["code"]).pack(side="left", padx=(0, 8))

        badge_f = tk.Frame(inner, bg=tc, padx=1, pady=1)
        badge_f.pack(side="left")
        tk.Label(badge_f, text=f" {icon} {d['type']} ",
                 font=("Consolas", 9, "bold"),
                 fg=C["bg"], bg=tc).pack()

        tk.Label(inner, text=d["size"], font=("Consolas", 12, "bold"),
                 fg=C["text"], bg=C["code"]).pack(side="left", padx=(10, 0))

        if d.get("model"):
            short = d["model"][:38] + ("…" if len(d["model"]) > 38 else "")
            tk.Label(inner, text=f"  {short}", font=("Consolas", 9),
                     fg=C["dim"], bg=C["code"]).pack(side="left")

    def _warn(self, msg):
        f = tk.Frame(self._body, bg=C["code"])
        f.pack(fill="x", padx=20, pady=1)
        tk.Label(f, text=f"  ⚠   {msg}",
                 font=("Consolas", 10), fg=C["red"], bg=C["code"],
                 pady=8).pack(anchor="w", padx=10)

    def _field(self, label, ph, saved, required=False):
        f = tk.Frame(self._body, bg=C["bg"])
        f.pack(fill="x", padx=20, pady=3)

        # Label container — fixed width so all inputs align
        lbl_f = tk.Frame(f, bg=C["bg"])
        lbl_f.pack(side="left")
        lbl_text = tk.Label(lbl_f, text=label, font=("Consolas", 10),
                            fg=C["dim"], bg=C["bg"],
                            width=13, anchor="w")
        lbl_text.pack(side="left")
        if required:
            tk.Label(lbl_f, text="*", font=("Consolas", 10, "bold"),
                     fg=C["red"], bg=C["bg"]).pack(side="left", padx=(0, 4))

        wrap = tk.Frame(f, bg=C["border"], padx=1, pady=1)
        wrap.pack(side="left", fill="x", expand=True)
        e = tk.Entry(wrap, font=("Consolas", 10),
                     bg=C["card2"], fg=C["text"],
                     insertbackground=C["cyan"],
                     relief="flat", bd=0,
                     selectbackground=C["accent"])
        e.pack(fill="x", ipady=8, padx=8)
        e._wrap = wrap
        e._ph   = ph
        e._req  = required
        if saved:
            e.insert(0, saved)
        else:
            e.insert(0, ph)
            e.configure(fg=C["dim"])
            def _fi(ev, en=e, p=ph):
                if en.get() == p:
                    en.delete(0, "end")
                    en.configure(fg=C["text"])
                    en._wrap.configure(bg=C["border"])
            def _fo(ev, en=e, p=ph):
                if not en.get().strip():
                    en.insert(0, p)
                    en.configure(fg=C["dim"])
            e.bind("<FocusIn>",  _fi)
            e.bind("<FocusOut>", _fo)
        return e

    def _field_with_check(self, label, ph, saved, bool_var, required=False):
        """Like _field but with a styled checkbox on the right side."""
        f = tk.Frame(self._body, bg=C["bg"])
        f.pack(fill="x", padx=20, pady=3)

        # Label
        lbl_f = tk.Frame(f, bg=C["bg"])
        lbl_f.pack(side="left")
        tk.Label(lbl_f, text=label, font=("Consolas", 10),
                 fg=C["dim"], bg=C["bg"],
                 width=13, anchor="w").pack(side="left")
        if required:
            tk.Label(lbl_f, text="*", font=("Consolas", 10, "bold"),
                     fg=C["red"], bg=C["bg"]).pack(side="left", padx=(0, 4))

        # Checkbox toggle button on the right (custom styled)
        chk_frame = tk.Frame(f, bg=C["bg"])
        chk_frame.pack(side="right", padx=(6, 0))

        def _refresh_chk():
            if bool_var.get():
                chk_lbl.configure(text="☑  Use as filename", fg=C["cyan"], bg=C["card2"])
            else:
                chk_lbl.configure(text="☐  Use as filename", fg=C["dim"],  bg=C["code"])

        chk_lbl = tk.Label(chk_frame,
                           text="☐  Use as filename",
                           font=("Consolas", 9),
                           fg=C["dim"], bg=C["code"],
                           cursor="hand2", padx=8, pady=4)
        chk_lbl.pack()

        def _toggle(e=None):
            bool_var.set(not bool_var.get())
            _refresh_chk()

        chk_lbl.bind("<Button-1>", _toggle)
        chk_lbl.bind("<Enter>", lambda e: chk_lbl.configure(bg=C["card2"]))
        chk_lbl.bind("<Leave>", lambda e: chk_lbl.configure(
            bg=C["card2"] if bool_var.get() else C["code"]))

        # Entry field
        wrap = tk.Frame(f, bg=C["border"], padx=1, pady=1)
        wrap.pack(side="left", fill="x", expand=True)
        e = tk.Entry(wrap, font=("Consolas", 10),
                     bg=C["card2"], fg=C["text"],
                     insertbackground=C["cyan"],
                     relief="flat", bd=0,
                     selectbackground=C["accent"])
        e.pack(fill="x", ipady=8, padx=8)
        e._wrap = wrap
        e._ph   = ph
        e._req  = required
        if saved:
            e.insert(0, saved)
        else:
            e.insert(0, ph)
            e.configure(fg=C["dim"])
            def _fi(ev, en=e, p=ph):
                if en.get() == p:
                    en.delete(0, "end")
                    en.configure(fg=C["text"])
                    en._wrap.configure(bg=C["border"])
            def _fo(ev, en=e, p=ph):
                if not en.get().strip():
                    en.insert(0, p)
                    en.configure(fg=C["dim"])
            e.bind("<FocusIn>",  _fi)
            e.bind("<FocusOut>", _fo)
        return e

    def _report(self):
        d  = self._info
        av = self._e_assign.get().strip()
        dv = self._e_department.get().strip()
        lv = self._e_location.get().strip()
        bv = self._e_brand.get().strip()
        nv = self._e_note.get("1.0", "end").strip()
        disks = "\n".join(
            f"    Disk {k['idx']}  ─  {k['size']}  [{k['type']}]  {k['model']}"
            for k in d["disks"]) or "    N/A"
        return (
            f"PC INFORMATION REPORT\n"
            f"{'═'*56}\n"
            f"Whoami        : {d['whoami']}\n"
            f"Hostname      : {d['hostname']}\n"
            f"IP Address    : {d['local_ip']}\n"
            f"Serial No.    : {d['serial']}\n"
            f"OS            : {d['os']}\n"
            f"CPU           : {d['cpu']}\n"
            f"RAM           : {d['ram']}\n"
            f"GPU           : {d['gpu']}\n"
            f"Motherboard   : {d['motherboard']}\n"
            f"\nStorage :\n{disks}\n"
            f"{'─'*56}\n"
            f"Assign To     : {av or 'N/A'}\n"
            f"Department    : {dv or 'N/A'}\n"
            f"Location      : {lv or 'N/A'}\n"
            f"Brand         : {bv or 'N/A'}\n"
            f"Notes         : {nv or 'N/A'}\n"
            f"{'═'*56}\n"
        )

    def _qr_data_str(self):
        """Build the compact text embedded in the QR code."""
        d  = self._info
        av = self._e_assign.get().strip()     if hasattr(self, '_e_assign')     else ""
        dv = self._e_department.get().strip() if hasattr(self, '_e_department') else ""
        lv = self._e_location.get().strip()   if hasattr(self, '_e_location')   else ""
        disks = " | ".join(
            f"Disk{k['idx']}:{k['size']} [{k['type']}]"
            for k in d["disks"]) or "N/A"
        return (
            f"HOST:{d['hostname']}\n"
            f"IP:{d['local_ip']}\n"
            f"SERIAL:{d['serial']}\n"
            f"OS:{d['os']}\n"
            f"CPU:{d['cpu']}\n"
            f"RAM:{d['ram']}\n"
            f"GPU:{d['gpu']}\n"
            f"DISKS:{disks}\n"
            f"ASSIGN:{av or 'N/A'}\n"
            f"DEPT:{dv or 'N/A'}\n"
            f"LOCATION:{lv or 'N/A'}"
        )

    def _make_qr_image(self, size=160):
        """Return a styled PIL Image of the QR code (dark bg, cyan modules)."""
        data = self._qr_data_str()
        qr = qrcode.QRCode(
            version=None,
            error_correction=qrcode.constants.ERROR_CORRECT_M,
            box_size=4,
            border=2,
        )
        qr.add_data(data)
        qr.make(fit=True)
        # Custom colors: cyan on dark background
        qr_img = qr.make_image(
            fill_color="#00D4FF",
            back_color="#070D1A"
        ).convert("RGB")
        qr_img = qr_img.resize((size, size), Image.NEAREST)
        # Add thin cyan border around the image
        bordered = Image.new("RGB", (size + 4, size + 4), "#00D4FF")
        bordered.paste(qr_img, (2, 2))
        return bordered

    def _build_qr(self):
        """Generate QR in background thread, then update the QR panel display."""
        try:
            img = self._make_qr_image(size=160)
            def _show():
                try:
                    tk_img = ImageTk.PhotoImage(img)
                    self._qr_tk = tk_img
                    self._qr_lbl.configure(image=tk_img, text="",
                                           width=0, height=0, bg=C["code"])
                except Exception:
                    pass
            self.after(0, _show)
        except Exception as ex:
            self.after(0, lambda: self._qr_lbl.configure(
                text=f"QR Error:\n{ex}", fg=C["red"]))

    def _save_qr(self, folder, filename_base):
        """Save the QR PNG alongside the report. Returns saved path or None."""
        try:
            img = self._make_qr_image(size=300)
            qr_path = os.path.join(folder, filename_base + "_QR.png")
            img.save(qr_path, "PNG")
            return qr_path
        except Exception:
            return None

    def _generate_and_save_qr(self, folder, filename_base):
        """Generate QR PIL images in background thread, then hand off to main thread for display."""
        try:
            # Build both sizes in the background thread (pure PIL, safe)
            img_display = self._make_qr_image(size=160)
            img_save    = self._make_qr_image(size=300)

            # Save PNG file (background thread is fine for file I/O)
            qr_path = os.path.join(folder, filename_base + "_QR.png")
            img_save.save(qr_path, "PNG")

            # Hand the PIL image to the main thread for ImageTk conversion + display
            def _show_on_main():
                try:
                    tk_img = ImageTk.PhotoImage(img_display)
                    self._qr_tk = tk_img          # prevent garbage collection
                    self._qr_lbl.configure(image=tk_img, text="",
                                           width=0, height=0,
                                           bg=C["code"])
                except Exception:
                    pass
                self._status_msg(
                    f"✓  Saved  →  {filename_base}.txt  +  QR.png", C["green"])

            self.after(0, _show_on_main)

        except Exception as ex:
            self.after(0, lambda: self._status_msg(
                f"✗  QR error: {str(ex)[:60]}", C["red"]))

    def _save_qr(self, folder, filename_base):
        """Save QR PNG only (no display update). Returns path or None."""
        try:
            img = self._make_qr_image(size=300)
            qr_path = os.path.join(folder, filename_base + "_QR.png")
            img.save(qr_path, "PNG")
            return qr_path
        except Exception:
            return None

    def _validate_required(self):
        """Check required fields. Flash red border on empty ones. Return True if all OK."""
        PLACEHOLDERS = {
            "Full name ...", "IT / HR / Finance ...",
            "Office / Floor ...", "Dell / HP / ..."
        }
        fields = [
            (self._e_assign,     "Assign To"),
            (self._e_department, "Department"),
            (self._e_location,   "Location"),
            (self._e_brand,      "Brand"),
        ]
        missing = []
        for e, name in fields:
            val = e.get().strip()
            if not val or val in PLACEHOLDERS:
                missing.append(name)
                e._wrap.configure(bg=C["red"])
                self.after(2500, lambda w=e._wrap: w.configure(bg=C["border"]))
        if missing:
            self._status_msg(f"⚠  Required:  {',  '.join(missing)}", C["red"])
            return False
        return True

    def _status_msg(self, msg, col):
        self._status.configure(text=f"  {msg}", fg=col)
        self.after(5000, lambda: self._status.configure(text="  Ready.", fg=C["dim"]))

    def _safe_filename(self):
        d = self._info
        ip_safe = d["local_ip"].replace(":", "_")

        # If checkbox is checked → "AssignName (IPA)"
        if hasattr(self, '_use_assign_name') and self._use_assign_name.get():
            assign = self._e_assign.get().strip()
            # Sanitize assign name for use in filename
            assign_safe = assign
            for ch in r'\/*?"<>|:': assign_safe = assign_safe.replace(ch, "_")
            fn = f"{assign_safe} ({ip_safe}A)"
        else:
            fn = f"{d['hostname']} ({ip_safe})"

        for ch in r'\/*?"<>|': fn = fn.replace(ch, "_")
        return fn

    def _persist_fields(self):
        try:
            with open(SAVE_FILE, "w", encoding="utf-8") as f:
                json.dump({
                    "assign_to":   self._e_assign.get().strip(),
                    "department":  self._e_department.get().strip(),
                    "location":    self._e_location.get().strip(),
                    "brand":       self._e_brand.get().strip(),
                    "note":        self._e_note.get("1.0","end").strip(),
                    "server_path": self._e_server.get().strip(),
                }, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    # ── Save locally ───────────────────────────
    def _save_local(self):
        if not self._info: return
        if not self._validate_required(): return
        self._persist_fields()
        fn   = self._safe_filename()
        path = os.path.join(_app_dir(), fn + ".txt")
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(self._report())
            # Generate + display QR, then save PNG
            threading.Thread(target=self._generate_and_save_qr,
                             args=(_app_dir(), fn), daemon=True).start()
            self._status_msg(f"✓  Saving  →  {fn}.txt  +  generating QR ...", C["green"])
        except Exception as e:
            self._status_msg(f"✗  {e}", C["red"])

    # ── Save to server shared folder ──────────
    def _save_server(self):
        if not self._info: return
        if not self._validate_required(): return
        srv = self._e_server.get().strip()
        PH  = r"\\SERVER\Reports"
        if not srv or srv == PH:
            self._status_msg("⚠  Please enter the server shared folder path first.", C["orange"])
            return

        self._persist_fields()
        fn   = self._safe_filename()
        path = os.path.join(srv, fn + ".txt")

        def _do():
            try:
                os.makedirs(srv, exist_ok=True)
                with open(path, "w", encoding="utf-8") as f:
                    f.write(self._report())
                self._generate_and_save_qr(srv, fn)
                self.after(0, lambda: self._status_msg(
                    f"✓  Saved to server  →  {fn}.txt  +  QR.png", C["cyan"]))
            except PermissionError:
                self.after(0, lambda: self._status_msg(
                    "✗  Access denied — check permissions on the shared folder.", C["red"]))
            except FileNotFoundError:
                self.after(0, lambda: self._status_msg(
                    "✗  Path not found — verify server name and share name.", C["red"]))
            except Exception as ex:
                self.after(0, lambda: self._status_msg(f"✗  {str(ex)[:70]}", C["red"]))

        threading.Thread(target=_do, daemon=True).start()
        self._status_msg("⏳  Connecting to server ...", C["dim"])

    def _copy(self):
        if not self._info: return
        try:
            self.clipboard_clear()
            self.clipboard_append(self._report())
            self.update()
            self._status_msg("✓  Report copied to clipboard.", C["gold"])
        except Exception as e:
            self._status_msg(f"✗  {e}", C["red"])

    def _folder(self):
        subprocess.Popen(f'explorer "{_app_dir()}"', shell=True)

    def _open_share(self):
        srv = self._e_server.get().strip()
        PH  = r"\\SERVER\Reports"
        if not srv or srv == PH:
            self._status_msg("⚠  Please enter the server path first.", C["orange"])
            return
        try:
            subprocess.Popen(f'explorer "{srv}"', shell=True)
            self._status_msg(f"⏳  Opening shared folder ...", C["cyan"])
        except Exception as ex:
            self._status_msg(f"✗  {str(ex)[:70]}", C["red"])


# ══════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════
if __name__ == "__main__":
    app = App()

    #Test data for development without psutil

    app.mainloop()
