import customtkinter as ctk
import uiautomation as auto
import threading
import time
import os
import sys
import comtypes
import ctypes
import tkinter as tk
import json
import psutil
import speech_recognition as sr
import math
import requests
from pydantic import BaseModel
from PIL import Image, ImageTk

# ==========================================
# LISTENING WAVE COMPONENT
# ==========================================
class ListeningWave:
    def __init__(self, parent, colors):
        self.canvas = tk.Canvas(parent, width=150, height=40, bg="#1a1b26", highlightthickness=0)
        self.colors = colors
        self.dots = []
        self.angle = 0
        self.is_animating = False
        
        for i in range(4):
            dot = self.canvas.create_oval(35 + i * 22, 15, 47 + i * 22, 27, fill=self.colors[i], outline="")
            self.dots.append(dot)
    
    def start(self):
        self.is_animating = True
        self.canvas.pack(pady=(0, 10))
        self.animate()
    
    def stop(self):
        self.is_animating = False
        self.canvas.pack_forget()
    
    def animate(self):
        if not self.is_animating:
            return
        
        for i, dot in enumerate(self.dots):
            y_offset = math.sin(self.angle + (i * 0.9)) * 10
            x1 = 35 + i * 22
            y1 = 15 + y_offset
            x2 = 47 + i * 22
            y2 = 27 + y_offset
            self.canvas.coords(dot, x1, y1, x2, y2)
        
        self.angle += 0.2
        self.canvas.after(20, self.animate)

# ==========================================
# SYSTEM OPTIMIZATION
# ==========================================
sys.stdout.reconfigure(encoding='utf-8', line_buffering=True)
try:
    p = psutil.Process(os.getpid())
    p.nice(psutil.HIGH_PRIORITY_CLASS)
except Exception:
    pass

# ==========================================
# CONFIGURATION
# ==========================================
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")
# ---------------------------------------------------------
# ⚠️ YOUR API KEY ⚠️
# ---------------------------------------------------------
MY_API_KEY = "AIzaSyCvOp0Ib4rTAQyDDssN2kkswq703WlElCE"

# ==========================================
# UI COMPONENTS
# ==========================================
class ListeningWave:
    def __init__(self, parent, colors):
        # Reduced height further to bring things up
        self.canvas = tk.Canvas(parent, width=150, height=20, bg="#1a1b26", highlightthickness=0) 
        self.colors = colors
        self.dots = []
        self.angle = 0
        self.is_animating = False
        
        for i in range(4):
            # Adjusted Y position for dots within the smaller canvas
            dot = self.canvas.create_oval(35 + i * 22, 5, 47 + i * 22, 17, fill=self.colors[i], outline="") 
            self.dots.append(dot)
    
    def start(self):
        self.is_animating = True
        self.canvas.pack(pady=(0, 2)) # Minimal padding
        self.animate()
    
    def stop(self):
        self.is_animating = False
        self.canvas.pack_forget()
    
    def animate(self):
        if not self.is_animating: 
            return
        
        for i, dot in enumerate(self.dots):
            y_offset = math.sin(self.angle + (i * 0.9)) * 4 # Slightly reduced amplitude
            x1 = 35 + i * 22
            y1 = 5 + y_offset
            x2 = 47 + i * 22
            y2 = 17 + y_offset
            self.canvas.coords(dot, x1, y1, x2, y2)
        
        self.angle += 0.2
        self.canvas.after(20, self.animate)

# ==========================================
# AUTO-DETECT MODEL
# ==========================================
def get_working_model():
    try:
        print("DEBUG: Detecting available models...")
        url = f"https://generativelanguage.googleapis.com/v1beta/models?key={MY_API_KEY}"
        response = requests.get(url)
        data = response.json()
        candidates = []
        for m in data.get('models', []):
            if 'generateContent' in m.get('supportedGenerationMethods', []):
                name = m['name'].replace("models/", "")
                if "flash" in name and "8b" not in name: 
                    return name
                candidates.append(name)
        if candidates: 
            return candidates[0]
    except: 
        pass
    return "gemini-1.5-flash"

CURRENT_MODEL = get_working_model()
print(f"✅ CONNECTED TO MODEL: {CURRENT_MODEL}")

# ==========================================
# AI CLIENT
# ==========================================

class PPTCommand(BaseModel):
    tab_name: str
    button_name: str
    explanation: str

def get_ai_command(user_text, current_tab="Unknown"):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{CURRENT_MODEL}:generateContent?key={MY_API_KEY}"
    headers = {'Content-Type': 'application/json'}
    
    prompt = f"""You are a PowerPoint Instructor.
CONTEXT: User is at '{current_tab}' tab.
REQ: "{user_text}"
RETURN JSON ONLY: {{ "tab_name": "...", "button_name": "...", "explanation": "..." }}"""

    data = { "contents": [{ "parts": [{"text": prompt}] }] }
    
    try:
        response = requests.post(url, headers=headers, json=data)
        if response.status_code == 200:
            result = response.json()
            try:
                raw = result['candidates'][0]['content']['parts'][0]['text']
                clean = raw.replace("```json", "").replace("```", "").strip()
                d = json.loads(clean)
                return PPTCommand(tab_name=d["tab_name"], button_name=d["button_name"], explanation="AI"), None
            except: 
                return None, "Parsing Error"
        elif response.status_code == 429: 
            return None, "⚠️ QUOTA FULL"
        elif response.status_code == 404: 
            return None, "⚠️ Model Error"
        else: 
            return None, f"Error {response.status_code}"
    except: 
        return None, "Connection Failed"

def translate_text_api(text, target_lang):
    if not text.strip(): 
        return ""
    
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{CURRENT_MODEL}:generateContent?key={MY_API_KEY}"
    headers = {'Content-Type': 'application/json'}
    
    prompt = f"Translate the following technical instruction to {target_lang}. Output ONLY the translated text. Do not add any introductory phrases like 'Here is the translation'. Text: '{text}'"
    data = { "contents": [{ "parts": [{"text": prompt}] }] }
    
    try:
        response = requests.post(url, headers=headers, json=data)
        if response.status_code == 200:
            return response.json()['candidates'][0]['content']['parts'][0]['text'].strip()
    except: 
        pass
    return text
# ==========================================
# GHOST CURSOR
# ==========================================
class GhostCursor:
    def __init__(self):
        self.stop_signal = threading.Event()
    
    def point_at(self, x, y, text):
        self.stop_signal.clear()
        t = threading.Thread(target=self._run_overlay, args=(x, y, text))
        t.start()
    
    def hide(self):
        self.stop_signal.set()
    
    def _draw_tooltip(self, canvas, tx, ty, text):
        text_x, text_y = tx + 25, ty + 25
        txt_item = canvas.create_text(text_x, text_y, text=text, fill="white", font=("Segoe UI", 12, "bold"), anchor="nw")
        bbox = canvas.bbox(txt_item)
        if bbox:
            x1, y1, x2, y2 = bbox
            bg_item = canvas.create_rectangle(x1-5, y1-5, x2+5, y2+5, fill="black", outline="black")
            canvas.tag_lower(bg_item, txt_item)
    
    def _run_overlay(self, tx, ty, text):
        comtypes.CoInitialize()
        try:
            root = tk.Tk()
            root.overrideredirect(True); root.attributes('-topmost', True); root.attributes('-alpha', 1.0)
            root.configure(bg="#010101"); root.wm_attributes("-transparentcolor", "#010101")
            sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
            root.geometry(f"{sw}x{sh}+0+0")
            
            canvas = tk.Canvas(root, width=sw, height=sh, bg="#010101", highlightthickness=0); canvas.pack()
            mx, my = auto.GetCursorPos()
            
            steps = 40
            for i in range(steps):
                if self.stop_signal.is_set(): 
                    break
                t_val = 1 - (1 - (i+1)/steps)**2
                cx = int(mx + (tx - mx) * t_val); cy = int(my + (ty - my) * t_val)
                canvas.delete("all")
                pts = [cx, cy, cx, cy+26, cx+7, cy+20, cx+11, cy+29, cx+16, cy+27, cx+12, cy+18, cx+21, cy+18]
                canvas.create_polygon(pts, outline="#2c2987", fill="#2c2987", width=2)
                self._draw_tooltip(canvas, cx, cy, text)
                root.update(); time.sleep(0.01)
            
            while not self.stop_signal.is_set():
                canvas.delete("all")
                pts = [tx, ty, tx, ty+26, tx+7, ty+20, tx+11, ty+29, tx+16, ty+27, tx+12, ty+18, tx+21, ty+18]
                canvas.create_polygon(pts, outline="#2c2987", fill="#2c2987", width=2)
                self._draw_tooltip(canvas, tx, ty, text)
                root.update(); time.sleep(0.05)
            
            root.destroy()
        except: 
            pass
        finally: 
            comtypes.CoUninitialize()

# ==========================================
# AUTOMATOR
# ==========================================
class Automator:
    def __init__(self, status_cb, script_cb):
        self.update_status = status_cb
        self.update_script = script_cb
        self.stop_event = threading.Event()
    
    def get_ppt_window(self):
        ppt = auto.WindowControl(searchDepth=1, ClassName="PPTFrameClass")
        if ppt.Exists(0, 0): 
            return ppt
        return None
    
    def get_current_tab_name(self, ppt_window): 
        return "Unknown" 
    
    def _wait_for_click(self, target_control):
        rect = target_control.BoundingRectangle
        while not self.stop_event.is_set():
            if ctypes.windll.user32.GetKeyState(0x01) < 0:
                mx, my = auto.GetCursorPos()
                if rect.left <= mx <= rect.right and rect.top <= my <= rect.bottom: 
                    return True
            time.sleep(0.1)
        return False

    def execute(self, cmd, ghost):
        self.stop_event.clear()
        ppt = self.get_ppt_window()
        if not ppt: 
            return "❌ PowerPoint not found!", False
        
        try:
            if not ppt.HasKeyboardFocus: 
                ppt.SetFocus()
        except: 
            pass
        
        self.update_script(f"Step 1: Open '{cmd.tab_name}' tab.\nStep 2: Click '{cmd.button_name}'.")
        
        # 1. Tab
        self.update_status(f"Step 1: Go to '{cmd.tab_name}'")
        tab = ppt.TabItemControl(Name=cmd.tab_name, searchDepth=10)
        if not tab.Exists(0, 1): 
            tab = ppt.TabItemControl(RegexName=f"(?i).*{cmd.tab_name}.*", searchDepth=10)
        
        if tab.Exists(0, 1):
            try:
                if not tab.GetSelectionItemPattern().IsSelected:
                    r = tab.BoundingRectangle
                    ghost.point_at((r.left+r.right)//2, (r.top+r.bottom)//2, f"Click '{cmd.tab_name}'")
                    self._wait_for_click(tab)
                    ghost.hide()
                    if self.stop_event.is_set(): 
                        return "Stopped", False
                    time.sleep(0.5)
            except: 
                pass
        
        # 2. Button
        self.update_status(f"Step 2: Click '{cmd.button_name}'")
        btn = ppt.Control(Name=cmd.button_name, searchDepth=15)
        if not btn.Exists(0, 0): 
            btn = ppt.Control(RegexName=f"(?i).*{cmd.button_name}.*", searchDepth=15)
        if not btn.Exists(0, 0): 
            btn = ppt.PaneControl(Name="Ribbon").Control(RegexName=f"(?i).*{cmd.button_name}.*", searchDepth=5)
        
        if btn.Exists(0, 1):
            r = btn.BoundingRectangle
            ghost.point_at((r.left+r.right)//2, (r.top+r.bottom)//2, f"Click '{cmd.button_name}'")
            self._wait_for_click(btn)
            ghost.hide()
            if self.stop_event.is_set(): 
                return "Stopped", False
            return "✅ Excellent!", True
        
        ghost.hide()
        return f"❌ Cannot find '{cmd.button_name}'", False

# ==========================================
# MAIN UI APPLICATION
# ==========================================
class WorkerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.TRANSPARENT_COLOR = "#ffff01"; self.BG_COLOR = "#1a1b26"; self.HEADER_COLOR = "#1f2335"; self.ACCENT_COLOR = "#24283b"
        
        self.geometry("60x60+100+100"); self.overrideredirect(True); self.attributes('-topmost', True)
        self.configure(fg_color=self.TRANSPARENT_COLOR); self.wm_attributes("-transparentcolor", self.TRANSPARENT_COLOR)
        
        self.is_expanded = False; self.is_listening = False; self.last_input = ""; self.recognizer = sr.Recognizer()
        self.ghost = GhostCursor(); self.automator = Automator(self.update_status_safe, self.update_script_safe)
        
        script_dir = os.path.dirname(os.path.abspath(__file__)); self.images = {}
        
        try: 
            self.logo_img = ctk.CTkImage(Image.open(os.path.join(script_dir, "logo.ico")), size=(20, 20)) 
        except: 
            self.logo_img = None
        
        files = {"idle": "robo_idle.png", "listening": "robo_listening.png", "thinking": "robo_thinking.png", "success": "robo_success.png", "error": "robo_error.png"}
        self.mini_icon_img = None
        for k, v in files.items():
            path = os.path.join(script_dir, v)
            if os.path.exists(path):
                img = Image.open(path); self.images[k] = ctk.CTkImage(img, size=(130, 130)) # Slightly smaller image
                if k == "idle": 
                    self.mini_icon_img = ctk.CTkImage(img, size=(60, 60))
        
        try: 
            self.mic_icon = ctk.CTkImage(Image.open(os.path.join(script_dir, "mic.png")), size=(18, 18))
        except: 
            self.mic_icon = None
        
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent"); self.main_frame.pack(fill="both", expand=True)
        
        self.btn_mini = ctk.CTkButton(self.main_frame, text="🤖", image=self.mini_icon_img, width=60, height=60, fg_color="transparent", command=self.animate_expansion)
        self.btn_mini.place(relx=0.5, rely=0.5, anchor="center")
        
        self.full_ui_frame = ctk.CTkFrame(self.main_frame, fg_color=self.BG_COLOR, corner_radius=15)
        
        # Header
        self.header = ctk.CTkFrame(self.full_ui_frame, fg_color=self.HEADER_COLOR, height=35, corner_radius=15); self.header.pack(fill="x", padx=2, pady=2) # Slightly shorter header
        if self.logo_img: 
            self.icon_lbl = ctk.CTkLabel(self.header, text="", image=self.logo_img); self.icon_lbl.pack(side="left", padx=(15, 5))
        ctk.CTkLabel(self.header, text="Instructly", font=("Segoe UI", 12, "bold")).pack(side="left", padx=0)
        ctk.CTkButton(self.header, text="✕", width=30, height=30, fg_color="transparent", hover_color="#ff4444", command=self.close_app).pack(side="right", padx=5)
        ctk.CTkButton(self.header, text="—", width=30, height=30, fg_color="transparent", command=self.animate_contraction).pack(side="right", padx=2)
        
        # Body (Reduced Padding)
        # Reduced pady significantly to move avatar UP
        self.avatar_display = ctk.CTkLabel(self.full_ui_frame, text="", image=self.images.get("idle")); self.avatar_display.pack(pady=(2, 0)) 
        self.dot_wave = ListeningWave(self.full_ui_frame, ["#4285F4", "#EA4335", "#FBBC05", "#34A853"])
        self.lbl_status = ctk.CTkLabel(self.full_ui_frame, text="Ready.", font=("Segoe UI", 14), text_color="#a9b1d6"); self.lbl_status.pack(pady=(0, 2))
        
        # Footer (Moved UP by reducing padding in pack)
        self.footer = ctk.CTkFrame(self.full_ui_frame, fg_color="transparent"); self.footer.pack(fill="x", padx=15, pady=2, side="bottom")
        
        # 1. Button Bar (Redo/Stop)
        self.btn_bar = ctk.CTkFrame(self.footer, fg_color="transparent"); self.btn_bar.pack(fill="x", pady=(0, 2))
        ctk.CTkButton(self.btn_bar, text="🔄 Redo Last", height=28, fg_color=self.ACCENT_COLOR, command=self.redo_action).pack(side="left", expand=True, padx=(0, 5), fill="x")
        ctk.CTkButton(self.btn_bar, text="🛑 Stop Action", height=28, fg_color="#f7768e", text_color="black", command=self.stop_action).pack(side="left", expand=True, padx=(5, 0), fill="x")
        
        # 2. Suggestions (Try it)
        self.suggestion_frame = ctk.CTkFrame(self.footer, fg_color="transparent"); self.suggestion_frame.pack(fill="x", pady=(0, 2))
        self.lbl_suggest = ctk.CTkLabel(self.suggestion_frame, text="Try it:", font=("Segoe UI", 10), text_color="#565f89"); self.lbl_suggest.pack(side="left")
        self.btn_try_cmd = ctk.CTkButton(self.suggestion_frame, text="Add Text", font=("Segoe UI", 10, "underline"), fg_color="transparent", hover_color=self.BG_COLOR, text_color="#7aa2f7", width=10, command=lambda: self.quick_command("Add Text"))
        self.btn_try_cmd.pack(side="left", padx=2)
        
        # 3. Search Bar (Command Input)
        self.input_bg = ctk.CTkFrame(self.footer, fg_color="black", corner_radius=18, height=35); self.input_bg.pack(fill="x", pady=(0, 5))
        self.entry = ctk.CTkEntry(self.input_bg, placeholder_text="Command...", border_width=0, fg_color="transparent", font=("Segoe UI", 12)); self.entry.pack(side="left", fill="both", expand=True, padx=(10, 5))
        self.entry.bind("<Return>", self.start_execution)
        ctk.CTkButton(self.input_bg, text="", image=self.mic_icon, width=30, height=30, fg_color="transparent", command=self.start_voice).pack(side="right")
        ctk.CTkButton(self.input_bg, text="➤", width=30, height=30, fg_color="transparent", command=lambda: self.start_execution(None)).pack(side="right", padx=(0, 5))
        
        # 4. Script Box & Translate
        self.script_frame = ctk.CTkFrame(self.footer, fg_color="#15161e", corner_radius=10)
        self.script_frame.pack(fill="x", pady=(0, 0))
        # Height remains 90 to be readable
        self.script_box = ctk.CTkTextbox(self.script_frame, height=90, fg_color="transparent", font=("Consolas", 11), text_color="#ffffff")
        self.script_box.pack(fill="both", padx=5, pady=2)
        self.script_box.insert("0.0", "Script will appear here...")
        self.script_box.configure(state="disabled")
        
        self.trans_frame = ctk.CTkFrame(self.script_frame, fg_color="transparent", height=20)
        self.trans_frame.pack(fill="x", padx=5, pady=(0,2))
        self.lang_menu = ctk.CTkOptionMenu(self.trans_frame, values=["English", "Hindi", "Telugu", "Tamil", "Spanish", "French", "German"], width=70, height=18, font=("Segoe UI", 9))
        self.lang_menu.pack(side="left")
        self.btn_trans = ctk.CTkButton(self.trans_frame, text="Translate", width=50, height=18, font=("Segoe UI", 9), fg_color=self.ACCENT_COLOR, command=self.do_translate)
        self.btn_trans.pack(side="right")

    def update_script_safe(self, text):
        self.script_box.configure(state="normal")
        self.script_box.delete("0.0", "end")
        self.script_box.insert("0.0", text)
        self.script_box.configure(state="disabled")

    def do_translate(self):
        target = self.lang_menu.get()
        original = self.script_box.get("0.0", "end").strip()
        if not original or "Script will" in original: 
            return
        self.btn_trans.configure(text="...", state="disabled")
        
        def run_trans():
            translated = translate_text_api(original, target)
            self.update_script_safe(translated)
            self.btn_trans.configure(text="Translate", state="normal")
        threading.Thread(target=run_trans).start()

    def quick_command(self, cmd_text): 
        self.entry.delete(0, 'end'); self.entry.insert(0, cmd_text); self.start_execution(None)

    def animate_expansion(self):
        cx, cy = self.winfo_x(), self.winfo_y(); self.btn_mini.place_forget(); self.configure(fg_color=self.BG_COLOR); self.wm_attributes("-transparentcolor", "")
        # Reduced overall height to 500 to tighten the UI
        self.geometry(f"340x500+{cx}+{cy}"); self.full_ui_frame.pack(fill="both", expand=True); self.entry.focus_set() 

    def animate_contraction(self):
        cx, cy = self.winfo_x(), self.winfo_y(); self.full_ui_frame.pack_forget(); self.configure(fg_color=self.TRANSPARENT_COLOR); self.wm_attributes("-transparentcolor", self.TRANSPARENT_COLOR)
        self.geometry(f"60x60+{cx}+{cy}"); self.btn_mini.place(relx=0.5, rely=0.5, anchor="center")

    def close_app(self): 
        self.quit(); sys.exit()

    def update_status_safe(self, msg): 
        self.lbl_status.configure(text=msg)

    def stop_action(self): 
        self.automator.stop_event.set(); self.ghost.hide(); self.update_status_safe("Stopped."); self.set_expression("error"); self.after(1500, self.reset_to_idle)

    def redo_action(self): 
        if self.last_input: 
            self.execute_thread(self.last_input)

    def start_voice(self):
        if self.is_listening: 
            return
        self.is_listening = True; self.avatar_display.configure(image=self.images.get("listening")); self.update_status_safe("Listening...")
        self.entry.delete(0, 'end'); self.dot_wave.start(); threading.Thread(target=self.bg_listen).start()

    def bg_listen(self):
        try:
            with sr.Microphone() as source: 
                self.recognizer.adjust_for_ambient_noise(source, 0.5); audio = self.recognizer.listen(source, timeout=5)
            text = self.recognizer.recognize_google(audio)
            self.after(0, lambda: self.entry.delete(0, 'end')); self.after(0, lambda: self.entry.insert(0, text)); self.after(700, lambda: self.start_execution(None))
        except: 
            self.update_status_safe("Voice Error"); self.after(0, lambda: self.set_expression("error")); time.sleep(1.5); self.after(0, self.reset_to_idle)
        finally: 
            self.is_listening = False; self.after(0, self.dot_wave.stop)

    def reset_to_idle(self): 
        self.set_expression("idle"); self.update_status_safe("Ready.")

    def start_execution(self, event):
        txt = self.entry.get(); if not txt: return
        self.last_input = txt; self.entry.delete(0, 'end'); self.execute_thread(txt)

    def set_expression(self, expr): 
        if expr in self.images: 
            self.avatar_display.configure(image=self.images[expr])

    def execute_thread(self, txt): 
        self.set_expression("thinking"); self.lbl_status.configure(text="Analyzing..."); threading.Thread(target=self.bg_execute, args=(txt,), daemon=True).start()

    def bg_execute(self, txt):
        comtypes.CoInitialize()
        try:
            ppt_win = self.automator.get_ppt_window(); current_tab = "Home" 
            if ppt_win: 
                current_tab = self.automator.get_current_tab_name(ppt_win)
            cmd, error = get_ai_command(txt, current_tab)
            if cmd: 
                msg, success = self.automator.execute(cmd, self.ghost); self.update_status_safe(msg); self.set_expression("success" if success else "error")
            else: 
                self.update_status_safe(error); self.set_expression("error")
        except: 
            self.update_status_safe("Error")
        finally: 
            comtypes.CoUninitialize(); time.sleep(1.5); self.after(0, self.reset_to_idle)

if __name__ == "__main__":
    app = WorkerApp()

    def start_move(e): 
        app.x = e.x; app.y = e.y

    def do_move(e): 
        try: 
            app.geometry(f"+{app.winfo_x() + e.x - app.x}+{app.winfo_y() + e.y - app.y}")
        except: 
            pass

    bind_list = [app.header, app.btn_mini, app.avatar_display]
    if hasattr(app, 'icon_lbl'): 
        bind_list.append(app.icon_lbl)
    
    for w in bind_list: 
        w.bind("<ButtonPress-1>", start_move); w.bind("<B1-Motion>", do_move)

    app.mainloop()