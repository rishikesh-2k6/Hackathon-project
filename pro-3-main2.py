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
import difflib
import math
from pydantic import BaseModel
from PIL import Image, ImageTk
import google.generativeai as genai

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
MY_API_KEY = "AIzaSyChcf8B7dBZneaGihSgLrxgGKLl-OGtJwA"

# ==========================================
# AI SETUP
# ==========================================
model = None
try:
    genai.configure(api_key=MY_API_KEY)
    available_models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
    selected_model_name = next((m for m in available_models if "gemini-1.5-flash" in m), None)
    if not selected_model_name:
        selected_model_name = available_models[0] if available_models else None
    if selected_model_name:
        model = genai.GenerativeModel(selected_model_name)
except Exception as e:
    print(f"❌ AI Init Error: {e}")

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

# ==========================================
# OFFLINE COMMAND CACHE
# ==========================================
OFFLINE_COMMANDS = {
    "font": {"tab": "Home", "btn": "Font"},
    "bold": {"tab": "Home", "btn": "Bold"},
    "italic": {"tab": "Home", "btn": "Italic"},
    "underline": {"tab": "Home", "btn": "Underline"},
    "new slide": {"tab": "Home", "btn": "New Slide"},
    "layout": {"tab": "Home", "btn": "Layout"},
    "text": {"tab": "Insert", "btn": "Text Box"},
    "textbox": {"tab": "Insert", "btn": "Text Box"},
    "picture": {"tab": "Insert", "btn": "Pictures"},
    "image": {"tab": "Insert", "btn": "Pictures"},
    "shape": {"tab": "Insert", "btn": "Shapes"},
    "icon": {"tab": "Insert", "btn": "Icons"},
    "chart": {"tab": "Insert", "btn": "Chart"},
    "animation": {"tab": "Animations", "btn": "Add Animation"},
    "animate": {"tab": "Animations", "btn": "Add Animation"},
    "fade": {"tab": "Animations", "btn": "Fade"},
    "fly in": {"tab": "Animations", "btn": "Fly In"},
    "play": {"tab": "Slide Show", "btn": "From Beginning"},
}

class PPTCommand(BaseModel):
    tab_name: str
    button_name: str
    explanation: str

def get_ai_command(user_text):
    clean = user_text.lower().strip()
    for key, val in OFFLINE_COMMANDS.items():
        if key in clean:
            return PPTCommand(tab_name=val["tab"], button_name=val["btn"], explanation="Offline"), "Offline"

    if not model: 
        return None, "AI Config Error"

    try:
        prompt = f"""Map this PowerPoint request to a Ribbon Tab and Button Name. 
Request: "{user_text}" 
Return JSON: {{"tab_name": "Insert", "button_name": "Shapes", "explanation": "..."}}"""
        response = model.generate_content(prompt)
        text_resp = response.text.replace("```json", "").replace("```", "").strip()
        data = json.loads(text_resp)
        return PPTCommand(tab_name=data["tab_name"], button_name=data["button_name"], explanation="AI"), None
    except Exception:
        return None, "AI Error"

# ==========================================
# GHOST CURSOR
# ==========================================
class GhostCursor:
    def show_and_move(self, target_x, target_y, text, stop_event=None):
        t = threading.Thread(target=self._run_overlay, args=(target_x, target_y, text, stop_event))
        t.start()
        t.join()

    def _run_overlay(self, tx, ty, text, stop_event):
        comtypes.CoInitialize()
        try:
            root = tk.Tk()
            root.overrideredirect(True)
            root.attributes('-topmost', True)
            root.attributes('-alpha', 1.0)
            root.configure(bg="#010101")
            root.wm_attributes("-transparentcolor", "#010101")
            sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
            root.geometry(f"{sw}x{sh}+0+0")

            canvas = tk.Canvas(root, width=sw, height=sh, bg="#010101", highlightthickness=0)
            canvas.pack()
            mx, my = auto.GetCursorPos()

            steps = 120
            for i in range(steps):
                if stop_event and stop_event.is_set(): 
                    break
                t_val = (i + 1) / steps
                t_val = 1 - (1 - t_val) * (1 - t_val)
                cx = int(mx + (tx - mx) * t_val)
                cy = int(my + (ty - my) * t_val)
                canvas.delete("all")
                pts = [cx, cy, cx, cy + 26, cx + 7, cy + 20, cx + 11, cy + 29, cx + 16, cy + 27, cx + 12, cy + 18,
                       cx + 21, cy + 18]
                canvas.create_polygon(pts, outline="#2c2987", fill="#2c2987", width=2)
                root.update()
                time.sleep(0.005)

            if not (stop_event and stop_event.is_set()):
                canvas.delete("all")
                padding = 10
                t_obj = canvas.create_text(0, 0, text=text, font=("Arial", 12, "bold"), anchor="nw")
                bbox = canvas.bbox(t_obj)
                canvas.delete(t_obj)
                w, h = bbox[2] - bbox[0] + 20, bbox[3] - bbox[1] + 20
                bx, by = tx + 15, ty + 40
                canvas.create_rectangle(bx, by, bx + w, by + h, fill="black", outline="black")
                canvas.create_text(bx + 10, by + 10, text=text, fill="white", font=("Arial", 12, "bold"), anchor="nw")
                pts = [tx, ty, tx, ty + 26, tx + 7, ty + 20, tx + 11, ty + 29, tx + 16, ty + 27, tx + 12, ty + 18,
                       tx + 21, ty + 18]
                canvas.create_polygon(pts, outline="#2c2987", fill="#2c2987", width=2)
                root.update()
                time.sleep(2.5)

            root.destroy()
        except Exception:
            pass
        finally:
            comtypes.CoUninitialize()

# ==========================================
# AUTOMATOR
# ==========================================
class Automator:
    def __init__(self, status_cb):
        self.update_status = status_cb
        self.stop_event = threading.Event()

    def get_topmost_ppt(self):
        foreground_hwnd = ctypes.windll.user32.GetForegroundWindow()
        ppt = auto.WindowControl(searchDepth=1, ClassName="PPTFrameClass")
        if ppt.Exists(0, 0):
            if ppt.NativeWindowHandle == foreground_hwnd:
                return ppt
            else:
                rect = ppt.BoundingRectangle
                if rect.width() > 0 and rect.height() > 0:
                    return ppt
        return None

    def execute(self, cmd, ghost):
        self.stop_event.clear()
        ppt = self.get_topmost_ppt()
        if not ppt:
            return "❌ PowerPoint not active!", False

        try:
            ppt.SetFocus()
            time.sleep(0.2)
        except Exception:
            pass

        if self.stop_event.is_set(): 
            return "Stopped", False

        self.update_status(f"Scanning {cmd.tab_name}...")
        tab = ppt.TabItemControl(Name=cmd.tab_name, searchDepth=10)
        if not tab.Exists(0, 1):
            return f"❌ Tab '{cmd.tab_name}' not found", False

        try:
            already_selected = tab.GetSelectionItemPattern().IsSelected
        except:
            already_selected = False

        if not already_selected:
            r = tab.BoundingRectangle
            ghost.show_and_move((r.left + r.right) // 2, (r.top + r.bottom) // 2, f"Click {tab.Name}", self.stop_event)
            if self.stop_event.is_set(): 
                return "Stopped", False
            tab.Click(simulateMove=False)
            time.sleep(0.4)
        else:
            self.update_status(f"{cmd.tab_name} is already open.")

        self.update_status(f"Directing to {cmd.button_name}...")
        btn = ppt.Control(Name=cmd.button_name, searchDepth=15)
        if not btn.Exists(0, 0):
            btn = ppt.Control(RegexName=f"(?i).*{cmd.button_name}.*", searchDepth=30)

        if btn.Exists(0, 1):
            r = btn.BoundingRectangle
            ghost.show_and_move((r.left + r.right) // 2, (r.top + r.bottom) // 2, f"Click {btn.Name}", self.stop_event)
            if self.stop_event.is_set(): 
                return "Stopped", False
            try:
                btn.Click(simulateMove=False)
                return "✅ Done!", True
            except Exception:
                auto.Click((r.left + r.right) // 2, (r.top + r.bottom) // 2)
                return "✅ Done!", True

        return f"❌ Cannot find '{cmd.button_name}'", False

# ==========================================
# MAIN UI APPLICATION
# ==========================================
class WorkerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.TRANSPARENT_COLOR = "#ffff01"
        self.BG_COLOR = "#1a1b26"
        self.HEADER_COLOR = "#1f2335"  # Slightly different shade for header
        self.ACCENT_COLOR = "#24283b"
        
        self.geometry("60x60+100+100")
        self.overrideredirect(True)
        self.attributes('-topmost', True)
        self.configure(fg_color=self.TRANSPARENT_COLOR)
        self.wm_attributes("-transparentcolor", self.TRANSPARENT_COLOR)
        
        self.is_expanded = False
        self.is_listening = False
        self.last_input = ""
        self.recognizer = sr.Recognizer()
        self.ghost = GhostCursor()
        self.automator = Automator(self.update_status_safe)
        
        script_dir = os.path.dirname(os.path.abspath(__file__))
        self.images = {}
        
        # Load Icon
        logo_path = os.path.join(script_dir, "logo.ico")
        self.logo_img = None
        try:
            if os.path.exists(logo_path):
                self.logo_img = ctk.CTkImage(Image.open(logo_path), size=(20, 20))
        except Exception:
            self.logo_img = None
        
        files = {"idle": "robo_idle.png", "listening": "robo_listening.png", "thinking": "robo_thinking.png",
                 "success": "robo_success.png", "error": "robo_error.png"}
        self.mini_icon_img = None
        for k, v in files.items():
            path = os.path.join(script_dir, v)
            if os.path.exists(path):
                try:
                    img = Image.open(path)
                    self.images[k] = ctk.CTkImage(img, size=(150, 150))
                    if k == "idle": 
                        self.mini_icon_img = ctk.CTkImage(img, size=(60, 60))
                except Exception:
                    self.images[k] = None
        
        mic_path = os.path.join(script_dir, "mic.png")
        self.mic_icon = None
        try:
            if os.path.exists(mic_path):
                self.mic_icon = ctk.CTkImage(Image.open(mic_path), size=(18, 18))
        except Exception:
            self.mic_icon = None
        
        # --- UI LAYOUT ---
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.pack(fill="both", expand=True)
        
        self.btn_mini = ctk.CTkButton(self.main_frame, text="🤖", image=self.mini_icon_img, width=60, height=60,
                                     fg_color="transparent", command=self.animate_expansion)
        self.btn_mini.place(relx=0.5, rely=0.5, anchor="center")
        
        self.full_ui_frame = ctk.CTkFrame(self.main_frame, fg_color=self.BG_COLOR, corner_radius=15)
        
        # --- MODIFIED HEADER WITH DIFFERENT COLOR AND ICON ---
        self.header = ctk.CTkFrame(self.full_ui_frame, fg_color=self.HEADER_COLOR, height=40, corner_radius=15)
        self.header.pack(fill="x", padx=2, pady=2)
        
        # App Icon in Header
        if self.logo_img:
            self.icon_lbl = ctk.CTkLabel(self.header, text="", image=self.logo_img)
            self.icon_lbl.pack(side="left", padx=(15, 5))
        
        ctk.CTkLabel(self.header, text="Instructly", font=("Segoe UI", 14, "bold")).pack(side="left", padx=0)
        ctk.CTkButton(self.header, text="✕", width=30, height=30, fg_color="transparent", hover_color="#ff4444",
                     command=self.close_app).pack(side="right", padx=5)
        ctk.CTkButton(self.header, text="—", width=30, height=30, fg_color="transparent",
                     command=self.animate_contraction).pack(side="right", padx=2)
        
        # Body Content
        self.avatar_display = ctk.CTkLabel(self.full_ui_frame, text="", image=self.images.get("idle"))
        self.avatar_display.pack(pady=(20, 10))
        
        self.dot_wave = ListeningWave(self.full_ui_frame, ["#4285F4", "#EA4335", "#FBBC05", "#34A853"])
        
        self.lbl_status = ctk.CTkLabel(self.full_ui_frame, text="Ready.", font=("Segoe UI", 16), text_color="#a9b1d6")
        self.lbl_status.pack()
        
        self.footer = ctk.CTkFrame(self.full_ui_frame, fg_color="transparent")
        self.footer.pack(fill="x", padx=20, pady=20, side="bottom")
        
        # Command Buttons
        self.btn_bar = ctk.CTkFrame(self.footer, fg_color="transparent")
        self.btn_bar.pack(fill="x", pady=(0, 5))
        
        ctk.CTkButton(self.btn_bar, text="🔄 Redo Last", height=32, fg_color=self.ACCENT_COLOR,
                     command=self.redo_action).pack(side="left", expand=True, padx=(0, 5), fill="x")
        ctk.CTkButton(self.btn_bar, text="🛑 Stop Action", height=32, fg_color="#f7768e", text_color="black",
                     command=self.stop_action).pack(side="left", expand=True, padx=(5, 0), fill="x")
        
        # Suggestion Area
        self.suggestion_frame = ctk.CTkFrame(self.footer, fg_color="transparent")
        self.suggestion_frame.pack(fill="x", pady=(0, 5))
        
        self.lbl_suggest = ctk.CTkLabel(self.suggestion_frame, text="Try it:", font=("Segoe UI", 11),
                                       text_color="#565f89")
        self.lbl_suggest.pack(side="left")
        
        self.btn_try_cmd = ctk.CTkButton(self.suggestion_frame, text="Add Text",
                                        font=("Segoe UI", 11, "underline"),
                                        fg_color="transparent", hover_color=self.BG_COLOR,
                                        text_color="#7aa2f7", width=10,
                                        command=lambda: self.quick_command("Add Text"))
        self.btn_try_cmd.pack(side="left", padx=2)
        
        # Input Box
        self.input_bg = ctk.CTkFrame(self.footer, fg_color="black", corner_radius=20, height=45)
        self.input_bg.pack(fill="x")
        
        self.entry = ctk.CTkEntry(self.input_bg, placeholder_text="Command...", border_width=0, fg_color="transparent")
        self.entry.pack(side="left", fill="both", expand=True, padx=(15, 5))
        self.entry.bind("<Return>", self.start_execution)
        
        ctk.CTkButton(self.input_bg, text="", image=self.mic_icon, width=35, height=35, fg_color="transparent",
                     command=self.start_voice).pack(side="right")
        ctk.CTkButton(self.input_bg, text="➤", width=35, height=35, fg_color="transparent",
                     command=lambda: self.start_execution(None)).pack(side="right", padx=(0, 5))

    def quick_command(self, cmd_text):
        self.entry.delete(0, 'end')
        self.entry.insert(0, cmd_text)
        self.start_execution(None)

    def animate_expansion(self):
        cx, cy = self.winfo_x(), self.winfo_y()
        self.btn_mini.place_forget()
        self.configure(fg_color=self.BG_COLOR)
        self.wm_attributes("-transparentcolor", "")
        self.geometry(f"320x480+{cx}+{cy}")
        self.full_ui_frame.pack(fill="both", expand=True)
        self.entry.focus_set()

    def animate_contraction(self):
        cx, cy = self.winfo_x(), self.winfo_y()
        self.full_ui_frame.pack_forget()
        self.configure(fg_color=self.TRANSPARENT_COLOR)
        self.wm_attributes("-transparentcolor", self.TRANSPARENT_COLOR)
        self.geometry(f"60x60+{cx}+{cy}")
        self.btn_mini.place(relx=0.5, rely=0.5, anchor="center")

    def close_app(self):
        self.quit()
        sys.exit()

    def update_status_safe(self, msg):
        self.lbl_status.configure(text=msg)

    def stop_action(self):
        self.automator.stop_event.set()
        self.update_status_safe("Stopped.")
        self.set_expression("error")
        self.after(1500, self.reset_to_idle)

    def redo_action(self):
        if self.last_input: 
            self.execute_thread(self.last_input)

    def start_voice(self):
        if self.is_listening: 
            return
        self.is_listening = True
        self.avatar_display.configure(image=self.images["listening"])
        self.update_status_safe("Listening...")
        self.entry.delete(0, 'end')
        self.dot_wave.start()
        threading.Thread(target=self.bg_listen).start()

    def bg_listen(self):
        try:
            with sr.Microphone() as source:
                self.recognizer.adjust_for_ambient_noise(source, 0.5)
                audio = self.recognizer.listen(source, timeout=5)
            text = self.recognizer.recognize_google(audio)
            self.after(0, lambda: self.entry.delete(0, 'end'))
            self.after(0, lambda: self.entry.insert(0, text))
            self.after(700, lambda: self.start_execution(None))
        except Exception:
            self.update_status_safe("Voice Error")
            self.after(0, lambda: self.set_expression("error"))
            time.sleep(1.5)
            self.after(0, self.reset_to_idle)
        finally:
            self.is_listening = False
            self.after(0, self.dot_wave.stop)

    def reset_to_idle(self):
        self.set_expression("idle")
        self.update_status_safe("Ready.")

    def start_execution(self, event):
        txt = self.entry.get()
        if not txt: 
            return
        self.last_input = txt
        self.entry.delete(0, 'end')
        self.execute_thread(txt)

    def set_expression(self, expr):
        if expr in self.images: 
            self.avatar_display.configure(image=self.images[expr])

    def execute_thread(self, txt):
        self.set_expression("thinking")
        self.lbl_status.configure(text="Searching...")
        threading.Thread(target=self.bg_execute, args=(txt,), daemon=True).start()

    def bg_execute(self, txt):
        comtypes.CoInitialize()
        try:
            cmd, _ = get_ai_command(txt)
            if cmd:
                msg, success = self.automator.execute(cmd, self.ghost)
                self.update_status_safe(msg)
                self.set_expression("success" if success else "error")
            else:
                self.update_status_safe("Command Error")
                self.set_expression("error")
        finally:
            comtypes.CoUninitialize()
            time.sleep(1.5)
            self.after(0, self.reset_to_idle)

if __name__ == "__main__":
    app = WorkerApp()

    def start_move(e):
        app.x = e.x;
        app.y = e.y

    def stop_move(e):
        app.x = None;
        app.y = None

    def do_move(e):
        try:
            app.geometry(f"+{app.winfo_x() + e.x - app.x}+{app.winfo_y() + e.y - app.y}")
        except Exception:
            pass

    # Updated binding to include icon label for dragging
    bind_list = [app.header, app.btn_mini, app.avatar_display]
    if hasattr(app, 'icon_lbl'): 
        bind_list.append(app.icon_lbl)
    
    for w in bind_list:
        w.bind("<ButtonPress-1>", start_move)
        w.bind("<ButtonRelease-1>", stop_move)
        w.bind("<B1-Motion>", do_move)

    app.mainloop()