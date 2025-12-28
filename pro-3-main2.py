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
MY_API_KEY = "AIzaSyBQY4nZUriHn2xwO-UvC_eva4YqHQjCLN0"


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
    def __init__(self, status_cb):
        self.update_status = status_cb

    def execute(self, cmd, ghost):
        ppt = auto.WindowControl(searchDepth=1, ClassName="PPTFrameClass")
        if not ppt.Exists(0, 1): ppt = auto.WindowControl(searchDepth=1, RegexName=".*PowerPoint.*")
        if not ppt.Exists(0, 1): return "PowerPoint not found!", False

        try:
            if ppt.IsMinimized: ppt.Restore()
            ppt.SetFocus()
        except:
            pass

        self.update_status(f"Opening {cmd.tab_name}...")
        tab = ppt.TabItemControl(Name=cmd.tab_name, searchDepth=10)
        if not tab.Exists(0, 0): tab = ppt.TabItemControl(RegexName=f".*{cmd.tab_name}.*", searchDepth=10)

        if tab.Exists(0, 1):
            if not tab.GetSelectionItemPattern().IsSelected:
                r = tab.BoundingRectangle
                ghost.show_and_move((r.left + r.right) // 2, (r.top + r.bottom) // 2, f"Click {cmd.tab_name}")
                time.sleep(0.5)
                tab.Click(simulateMove=False)
            time.sleep(0.5)
        else:
            return f"Tab {cmd.tab_name} missing", False

        self.update_status(f"Finding {cmd.button_name}...")
        
        # First attempt: Try exact name match with shallow search for performance
        btn = ppt.Control(Name=cmd.button_name, searchDepth=15)
        
        # Second attempt: Use case-insensitive regex if exact match fails
        if not btn.Exists(0, 0):
            btn = ppt.Control(RegexName=f"(?i).*{cmd.button_name}.*", searchDepth=50)
        
        # Third attempt: Search within ribbon control for better button detection
        if not btn.Exists(0, 0):
            ribbon = ppt.PaneControl(RegexName=".*Ribbon.*", searchDepth=5)
            found_btn, found_name = self.find_best_match_button(ribbon, cmd.button_name)

        if btn.Exists(0, 1):
            r = btn.BoundingRectangle
            ghost.show_and_move((r.left + r.right) // 2, (r.top + r.bottom) // 2, f"Click {cmd.button_name}")
            return "Done!", True
        return f"Button {cmd.button_name} missing", False


# ==========================================
# UI (TRANSPARENT MINI MODE)
# ==========================================
class WorkerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- WINDOW SETUP ---
        # "Chroma Key" color for transparency (this color becomes invisible)
        self.TRANSPARENT_COLOR = "#ffff01"
        self.DARK_BG = "#1a1a1a"

        self.geometry("60x60+100+100")
        self.overrideredirect(True)
        self.attributes('-topmost', True)

        # Start in transparent mode (Mini)
        self.configure(fg_color=self.TRANSPARENT_COLOR)
        self.wm_attributes("-transparentcolor", self.TRANSPARENT_COLOR)

        self.is_expanded = False
        self.is_listening = False
        self.recognizer = sr.Recognizer()
        self.ghost = GhostCursor()
        self.automator = Automator(self.update_status_safe)

        # --- LOAD ASSETS ---
        script_dir = os.path.dirname(os.path.abspath(__file__))
        self.images = {}
        # ADDED new constants:
        self.BG_COLOR = "#1a1b26"      # NEW
        self.ACCENT_COLOR = "#24283b"  # NEW (unused)


        # 1. Start Icon (The Robot Face only)
        start_icon_path = os.path.join(script_dir, "start_icon.png")
        if os.path.exists(start_icon_path):
            self.start_icon_img = ctk.CTkImage(Image.open(start_icon_path), size=(50, 50))
        else:
            self.start_icon_img = None

        # 2. Robo Avatars
        files = {
            "idle": "robo_idle.png", "listening": "robo_listening.png",
            "thinking": "robo_thinking.png", "success": "robo_success.png",
            "error": "robo_error.png"
        }
        try:
            for k, v in files.items():
                path = os.path.join(script_dir, v)
                if os.path.exists(path):
                    self.images[k] = ctk.CTkImage(Image.open(path), size=(130, 130))
        except:
            pass

        # 3. Mic Icon
        mic_path = os.path.join(script_dir, "mic.png")
        self.mic_icon = None
        if os.path.exists(mic_path):
            self.mic_icon = ctk.CTkImage(Image.open(mic_path), size=(20, 20))

        # --- LAYOUT ---

        # Main Container
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent", corner_radius=0)
        self.main_frame.pack(fill="both", expand=True)

        # A. MINI MODE BUTTON (Transparent Background)
        self.btn_mini = ctk.CTkButton(
            self.main_frame,
            text="" if self.start_icon_img else "🤖",
            image=self.start_icon_img,
            width=60, height=60,
            fg_color="transparent",  # No frame color
            hover_color=None,  # No hover effect
            command=self.animate_expansion
        )
        self.btn_mini.place(relx=0.5, rely=0.5, anchor="center")

        # B. FULL MODE CONTAINER (Hidden)
        self.full_ui_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")

        # 1. Avatar Button
        self.avatar_btn = ctk.CTkButton(
            self.full_ui_frame,
            text="",
            image=self.images.get("idle"),
            fg_color="transparent",
            hover_color="#2a2a2a",
            width=130, height=130,
            command=self.animate_contraction
        )
        self.avatar_btn.pack(pady=(10, 0))

        # 2. Controls
        self.controls_frame = ctk.CTkFrame(self.full_ui_frame, fg_color="transparent")
        self.controls_frame.pack(pady=5)

        self.entry = ctk.CTkEntry(self.controls_frame, placeholder_text="Tell me what to do...", width=180)
        self.entry.pack(pady=5)
        self.entry.bind("<Return>", self.start_execution)

        self.btn_mic = ctk.CTkButton(
            self.controls_frame,
            text=" Voice" if self.mic_icon else "🎤 Voice",
            image=self.mic_icon,
            width=160,
            fg_color="#444444",
            command=self.start_voice_listening
        )
        self.btn_mic.pack(pady=5)

        self.btn_run = ctk.CTkButton(self.controls_frame, text="GO", width=160,
                                     command=lambda: self.start_execution(None))
        self.btn_run.pack(pady=5)

        self.lbl_status = ctk.CTkLabel(self.controls_frame, text="Ready.", text_color="gray", wraplength=180)
        self.lbl_status.pack(pady=5)

    def set_expression(self, expression):
        if self.images and expression in self.images:
            self.avatar_btn.configure(image=self.images[expression])

    def animate_expansion(self):
        self.btn_mini.place_forget()

        # Switch background to OPAQUE DARK for the main UI
        self.configure(fg_color=self.DARK_BG)

        steps = [("80x100", 0.02), ("120x150", 0.02), ("160x220", 0.02), ("220x380", 0.02)]
        for geo, delay in steps:
            self.geometry(geo)
            self.update()
            time.sleep(delay)

        self.full_ui_frame.pack(fill="both", expand=True)
        self.is_expanded = True

    def animate_contraction(self):
        self.full_ui_frame.pack_forget()

        # Switch background back to TRANSPARENT for floating robot
        self.configure(fg_color=self.TRANSPARENT_COLOR)

        steps = [("160x220", 0.02), ("120x150", 0.02), ("80x100", 0.02), ("60x60", 0.02)]
        for geo, delay in steps:
            self.geometry(geo)
            self.update()
            time.sleep(delay)

        self.btn_mini.place(relx=0.5, rely=0.5, anchor="center")
        self.is_expanded = False

    def update_status_safe(self, msg):
        self.lbl_status.configure(text=msg)

    def start_voice_listening(self):
        if self.is_listening: return
        self.is_listening = True
        self.lbl_status.configure(text="Listening...", text_color="#E53935")
        self.set_expression("listening")
        threading.Thread(target=self.bg_listen).start()

    def bg_listen(self):
        try:
            with sr.Microphone() as source:
                self.recognizer.adjust_for_ambient_noise(source, 0.5)
                audio = self.recognizer.listen(source, timeout=5)
            text = self.recognizer.recognize_google(audio)
            self.entry.delete(0, 'end');
            self.entry.insert(0, text)
            self.after(0, lambda: self.start_execution(None))
        except:
            self.set_expression("error")
        finally:
            self.is_listening = False

    def start_execution(self, event):
        txt = self.entry.get()
        if not txt: return
        self.set_expression("thinking")
        self.lbl_status.configure(text="Thinking...", text_color="#3B8ED0")
        threading.Thread(target=self.bg_execute, args=(txt,)).start()

    def bg_execute(self, txt):
        comtypes.CoInitialize()
        try:
            cmd, err = get_ai_command(txt)
            if cmd:
                self.update_status_safe(f"On it: {cmd.button_name}")
                msg, success = self.automator.execute(cmd, self.ghost)
                self.update_status_safe(msg)
                self.set_expression("success" if success else "error")
            else:
                self.update_status_safe("Confused.")
                self.set_expression("error")
        finally:
            comtypes.CoUninitialize()
            time.sleep(2)
            self.set_expression("idle")


if __name__ == "__main__":
    app = WorkerApp()


    # Drag Logic
    def start_move(event):
        app.x = event.x; app.y = event.y


    def stop_move(event):
        app.x = None; app.y = None


    def do_move(event):
        try:
            deltax = event.x - app.x
            deltay = event.y - app.y
            x = app.winfo_x() + deltax
            y = app.winfo_y() + deltay
            app.geometry(f"+{x}+{y}")
        except:
            pass


    app.bind("<ButtonPress-1>", start_move)
    app.bind("<ButtonRelease-1>", stop_move)
    app.bind("<B1-Motion>", do_move)

    app.mainloop()
