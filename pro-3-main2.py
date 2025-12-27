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
from pydantic import BaseModel
from PIL import Image, ImageTk
import google.generativeai as genai


# SYSTEM OPTIMIZATION
# ==========================================
sys.stdout.reconfigure(encoding='utf-8', line_buffering=True)
try:
    p = psutil.Process(os.getpid())
    p.nice(psutil.HIGH_PRIORITY_CLASS)
except:
    pass

# ==========================================
# CONFIGURATION
# ==========================================
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")
# CHANGE: Uncommented and added real key
MY_API_KEY = "AIzaSyChcf8B7dBZneaGihSgLrxgGKLl-OGtJwA"


# ==========================================
# AI SETUP
# ==========================================
model = None
try:
    genai.configure(api_key=MY_API_KEY)
    available_models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
    selected_model_name = None
    for m in available_models:
        if "gemini-1.5-flash" in m:
            selected_model_name = m
            break
    if not selected_model_name:
        for m in available_models:
            if "gemini-1.5-pro" in m:
                selected_model_name = m
                break
    if not selected_model_name and available_models:
        selected_model_name = available_models[0]
    if selected_model_name:
        model = genai.GenerativeModel(selected_model_name)
    else:
        print("❌ No AI models found.")
except Exception as e:
    print(f"❌ API Error: {e}")

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except:
    pass

# ==========================================
# OFFLINE CACHE
# ==========================================
OFFLINE_COMMANDS = {
    "font": {"tab": "Home", "btn": "Font"},
    "change font": {"tab": "Home", "btn": "Font"},
    "typeface": {"tab": "Home", "btn": "Font"},
    "font size": {"tab": "Home", "btn": "Font Size"},
    "text size": {"tab": "Home", "btn": "Font Size"},
    "color": {"tab": "Home", "btn": "Font Color"},
    "bold": {"tab": "Home", "btn": "Bold"},
    "italic": {"tab": "Home", "btn": "Italic"},
    "new slide": {"tab": "Home", "btn": "New Slide"},
    "layout": {"tab": "Home", "btn": "Layout"},
    "text": {"tab": "Insert", "btn": "Text Box"},
    "textbox": {"tab": "Insert", "btn": "Text Box"},
    "wordart": {"tab": "Insert", "btn": "WordArt"},
    "stylish": {"tab": "Insert", "btn": "WordArt"},
    "picture": {"tab": "Insert", "btn": "Pictures"},
    "image": {"tab": "Insert", "btn": "Pictures"},
    "shape": {"tab": "Insert", "btn": "Shapes"},
    "icon": {"tab": "Insert", "btn": "Icons"},
    "chart": {"tab": "Insert", "btn": "Chart"},
    "play": {"tab": "Slide Show", "btn": "From Beginning"},
    # adding new offline commands:
"underline": {"tab": "Home", "btn": "Underline"},
"arrange": {"tab": "Home", "btn": "Arrange"},
# Insert Tab
"icon": {"tab": "Insert", "btn": "Icons"},
# Animations Tab (NEW SECTION)
"animation": {"tab": "Animations", "btn": "Add Animation"},
"animate": {"tab": "Animations", "btn": "Add Animation"},
"fade": {"tab": "Animations", "btn": "Fade"},
"fly in": {"tab": "Animations", "btn": "Fly In"},
# ... 6 more animation commands

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

    if not model: return None, "AI Config Error"

    try:
        prompt = f"""
        Map user request to PowerPoint Ribbon Tab and Button.
        Rules: "Change font"->"Home"|"Font"; "Stylish"->"Insert"|"WordArt".
        Request: "{user_text}"
        Return JSON: {{"tab_name": "Insert", "button_name": "Shapes", "explanation": "..."}}
        """
        response = model.generate_content(prompt)
        text_resp = response.text.replace("```json", "").replace("```", "").strip()
        data = json.loads(text_resp)
        return PPTCommand(tab_name=data["tab_name"], button_name=data["button_name"], explanation="AI"), None
    except Exception as e:
        return None, str(e)


# ==========================================
# GHOST CURSOR
# ==========================================
class GhostCursor:
    def show_and_move(self, target_x, target_y, text):
        t = threading.Thread(target=self._run_overlay, args=(target_x, target_y, text))
        t.start()
        t.join()

    def _run_overlay(self, tx, ty, text):
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

            steps = 150
            for i in range(steps):
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

            canvas.delete("all")
            padding = 10
            t_obj = canvas.create_text(0, 0, text=text, font=("Arial", 12, "bold"), anchor="nw")
            bbox = canvas.bbox(t_obj)
            canvas.delete(t_obj)
            w, h = bbox[2] - bbox[0] + 20, bbox[3] - bbox[1] + 20
            bx, by = tx + 15, ty + 40
            canvas.create_rectangle(bx, by, bx + w, by + h, fill="black", outline="black")
            canvas.create_text(bx + 10, by + 10, text=text, fill="white", font=("Arial", 12, "bold"), anchor="nw")
            pts = [tx, ty, tx, ty + 26, tx + 7, ty + 20, tx + 11, ty + 29, tx + 16, ty + 27, tx + 12, ty + 18, tx + 21,
                   ty + 18]
            canvas.create_polygon(pts, outline="#2c2987", fill="#2c2987", width=2)
            root.update()

            time.sleep(3)
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
        btn = ppt.Control(Name=cmd.button_name, searchDepth=50)
        if not btn.Exists(0, 0): btn = ppt.Control(RegexName=f".*{cmd.button_name}.*", searchDepth=50)

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
