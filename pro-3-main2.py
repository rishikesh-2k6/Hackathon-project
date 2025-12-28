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
    def auto_click(self, control):
        # 1️⃣ InvokePattern (best)
        try:
            control.GetInvokePattern().Invoke()
            return True
        except:
            pass
        
        # 2️⃣ TogglePattern
        try:
            control.GetTogglePattern().Toggle()
            return True
        except:
            pass
        
        # 3️⃣ SelectionItemPattern
        try:
            control.GetSelectionItemPattern().Select()
            return True
        except:
            pass
        
        # 4️⃣ ExpandCollapsePattern
        try:
            control.GetExpandCollapsePattern().Expand()
            return True
        except:
            pass
        
        # 5️⃣ Mouse fallback
        try:
            r = control.BoundingRectangle
            x = (r.left + r.right) // 2
            y = (r.top + r.bottom) // 2
            auto.MoveTo(x, y)
            time.sleep(0.05)
            auto.Click(x, y)
            return True
        except:
            return False
    
    def point_control(self, ghost, control, text):
        try:
            r = control.BoundingRectangle
            x = (r.left + r.right) // 2
            y = (r.top + r.bottom) // 2
            ghost.point_at(x, y, text)
        except:
            pass
    
    def warm_up_control(self, control):
        try:
            r = control.BoundingRectangle
            x = (r.left + r.right) // 2
            y = (r.top + r.bottom) // 2
            auto.MoveTo(x, y)
            time.sleep(0.05)
        except:
            pass
    
    def bring_ppt_to_foreground(self, ppt_window):
        try:
            hwnd = ppt_window.NativeWindowHandle
            # Do NOT restore / maximize
            # Just allow interaction
            fg = ctypes.windll.user32.GetForegroundWindow()
            tid1 = ctypes.windll.user32.GetWindowThreadProcessId(fg, 0)
            tid2 = ctypes.windll.user32.GetWindowThreadProcessId(hwnd, 0)
            ctypes.windll.user32.AttachThreadInput(tid1, tid2, True)
            ctypes.windll.user32.SetForegroundWindow(hwnd)
            ctypes.windll.user32.AttachThreadInput(tid1, tid2, False)
        except Exception as e:
            print("Foreground attach failed:", e)

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

    def execute(self, cmd, ghost, auto_mode=False):
        ppt = self.get_ppt_window()
        if not ppt:
            return "❌ PowerPoint not found!", False
        
        self.bring_ppt_to_foreground(ppt)
        time.sleep(0.3)
        
        self.update_script(f"Step 1: Open '{cmd.tab_name}' tab.\n"
                          f"Step 2: Click '{cmd.button_name}'.")
        
        # ===============================
        # STEP 1 : TAB
        # ===============================
        self.update_status(f"Step 1: Go to '{cmd.tab_name}'")
        tab = ppt.TabItemControl(Name=cmd.tab_name, searchDepth=10)
        if not tab.Exists(0, 1):
            tab = ppt.TabItemControl(RegexName=f"(?i).*{cmd.tab_name}.*",
                                   searchDepth=10)
        
        if tab.Exists(0, 1):
            try:
                if not tab.GetSelectionItemPattern().IsSelected:
                    r = tab.BoundingRectangle
                    ghost.point_at((r.left + r.right) // 2,
                                 (r.top + r.bottom) // 2,
                                 f"{'Auto opening' if auto_mode else 'Click'} '{cmd.tab_name}'")
                    time.sleep(0.25)  # 🔑 allow ghost/UI to settle
                    
                    if auto_mode:
                        self.auto_click(tab)
                    else:
                        self._wait_for_click(tab)
                    
                    ghost.hide()
                    if self.stop_event.is_set():
                        return "Stopped", False
                    time.sleep(0.3)
            except:
                pass
        
        # ===============================
        # STEP 2 : BUTTON
        # ===============================
        self.update_status(f"Step 2: Click '{cmd.button_name}'")
        ribbon = ppt.PaneControl(Name="Ribbon", searchDepth=3)
        btn = ribbon.Control(Name=cmd.button_name, searchDepth=10)
        
        if not btn.Exists(0, 0):
            btn = ribbon.Control(RegexName=f"(?i).*{cmd.button_name}.*",
                               searchDepth=10)
        
        if not btn.Exists(0, 0):
            btn = ppt.PaneControl(Name="Ribbon").Control(RegexName=f"(?i).*{cmd.button_name}.*",
                                                       searchDepth=5)
        
        if not btn.Exists(0, 1):
            ghost.hide()
            return f"❌ Cannot find '{cmd.button_name}'", False
        
        # 🔹 Show ghost at LIVE position
        self.point_control(ghost,
                          btn,
                          f"{'Auto clicking' if auto_mode else 'Click'} '{cmd.button_name}'")
        
        # 🔹 Warm-up hover (CRITICAL for Ribbon)
        self.warm_up_control(btn)
        time.sleep(0.25)
        
        if auto_mode:
            self.auto_click(btn)
        else:
            self._wait_for_click(btn)
        
        ghost.hide()
        ghost.hide()
        if self.stop_event.is_set():
            return "Stopped", False
        
        return "✅ Excellent!", True


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
