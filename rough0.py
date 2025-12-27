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
from pydantic import BaseModel
from PIL import Image, ImageTk
import google.generativeai as genai

# ==========================================
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

    # 1. Offline Check
    for key, val in OFFLINE_COMMANDS.items():
        if key in clean:
            return PPTCommand(tab_name=val["tab"], button_name=val["btn"], explanation="Offline"), "Offline"

    if not model: return None, "AI Config Error"

    # 2. AI Check
    try:
        prompt = f"""
        Map this PowerPoint request to a Ribbon Tab and Button Name.
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

            steps = 120
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

            time.sleep(2.5)
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

    def find_best_match_button(self, parent, target_name):
        all_buttons = []
        try:
            children = parent.GetChildren()
            for child in children:
                if child.Name: all_buttons.append((child.Name, child))
                for sub in child.GetChildren():
                    if sub.Name: all_buttons.append((sub.Name, sub))
        except:
            pass

        names = [x[0] for x in all_buttons]
        matches = difflib.get_close_matches(target_name, names, n=1, cutoff=0.4)
        if matches:
            best = matches[0]
            for name, ctrl in all_buttons:
                if name == best: return ctrl, best
        return None, None

    def execute(self, cmd, ghost):
        ppt = auto.WindowControl(searchDepth=1, ClassName="PPTFrameClass")
        if not ppt.Exists(0, 1): ppt = auto.WindowControl(searchDepth=1, RegexName=".*PowerPoint.*")
        if not ppt.Exists(0, 1): return "PowerPoint not found!", False

        try:
            if ppt.IsMinimized: ppt.Restore()
            ppt.SetFocus()
        except:
            pass

        self.update_status(f"Tab: {cmd.tab_name}...")
        tab = ppt.TabItemControl(Name=cmd.tab_name, searchDepth=10)

        if not tab.Exists(0, 0):
            tabs = ppt.TabControl(searchDepth=5).GetChildren()
            for t in tabs:
                if cmd.tab_name.lower() in t.Name.lower() or t.Name.lower() in cmd.tab_name.lower():
                    tab = t;
                    break

        if tab.Exists(0, 1):
            if not tab.GetSelectionItemPattern().IsSelected:
                r = tab.BoundingRectangle
                ghost.show_and_move((r.left + r.right) // 2, (r.top + r.bottom) // 2, f"Click {tab.Name}")
                tab.Click(simulateMove=False)
                time.sleep(0.5)
        else:
            return f"Tab '{cmd.tab_name}' not found", False

        self.update_status(f"Scan: {cmd.button_name}")
        btn = ppt.Control(Name=cmd.button_name, searchDepth=15)

        if not btn.Exists(0, 0):
            btn = ppt.Control(RegexName=f"(?i).*{cmd.button_name}.*", searchDepth=50)

        if not btn.Exists(0, 0):
            self.update_status("Deep searching...")
            ribbon = ppt.PaneControl(RegexName=".*Ribbon.*", searchDepth=5)
            if ribbon.Exists(0, 0):
                found_btn, found_name = self.find_best_match_button(ribbon, cmd.button_name)
                if found_btn:
                    btn = found_btn
                    self.update_status(f"Found: {found_name}")

        if btn.Exists(0, 1):
            r = btn.BoundingRectangle
            ghost.show_and_move((r.left + r.right) // 2, (r.top + r.bottom) // 2, f"Click {btn.Name}")
            try:
                btn.Click(simulateMove=False)
                return "Done!", True
            except:
                return "Click blocked.", False

        return f"Cannot find '{cmd.button_name}'", False


# ==========================================
# UI
# ==========================================
class WorkerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.TRANSPARENT_COLOR = "#ffff01";
        self.BG_COLOR = "#1a1b26";
        self.ACCENT_COLOR = "#24283b"
        self.geometry("60x60+100+100");
        self.overrideredirect(True);
        self.attributes('-topmost', True)
        self.configure(fg_color=self.TRANSPARENT_COLOR);
        self.wm_attributes("-transparentcolor", self.TRANSPARENT_COLOR)

        self.is_expanded = False;
        self.is_listening = False;
        self.recognizer = sr.Recognizer()
        self.ghost = GhostCursor();
        self.automator = Automator(self.update_status_safe)

        # Assets
        script_dir = os.path.dirname(os.path.abspath(__file__))
        self.images = {}

        logo_path = os.path.join(script_dir, "logo.ico")
        if os.path.exists(logo_path):
            self.logo_img = ctk.CTkImage(Image.open(logo_path), size=(20, 20))
            try:
                self.iconbitmap(logo_path)
            except:
                pass
        else:
            self.logo_img = None

        files = {"idle": "robo_idle.png", "listening": "robo_listening.png", "thinking": "robo_thinking.png",
                 "success": "robo_success.png", "error": "robo_error.png"}
        self.mini_icon_img = None
        for k, v in files.items():
            path = os.path.join(script_dir, v)
            if os.path.exists(path):
                img = Image.open(path)
                self.images[k] = ctk.CTkImage(img, size=(150, 150))
                if k == "idle": self.mini_icon_img = ctk.CTkImage(img, size=(60, 60))
            else:
                self.images[k] = None

        mic_path = os.path.join(script_dir, "mic.png")
        self.mic_icon = ctk.CTkImage(Image.open(mic_path), size=(18, 18)) if os.path.exists(mic_path) else None

        # --- UI LAYOUT ---
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent", corner_radius=0);
        self.main_frame.pack(fill="both", expand=True)
        self.btn_mini = ctk.CTkButton(self.main_frame, text="🤖", image=self.mini_icon_img, width=60, height=60,
                                      fg_color="transparent", hover_color=None, command=self.animate_expansion)
        self.btn_mini.place(relx=0.5, rely=0.5, anchor="center")

        self.full_ui_frame = ctk.CTkFrame(self.main_frame, fg_color=self.BG_COLOR, corner_radius=15)
        self.header = ctk.CTkFrame(self.full_ui_frame, fg_color="transparent", height=35);
        self.header.pack(fill="x", padx=15, pady=(10, 0))
        self.lbl_logo = ctk.CTkLabel(self.header, text="", image=self.logo_img);
        self.lbl_logo.pack(side="left")
        self.lbl_title = ctk.CTkLabel(self.header, text="Instructly", font=("Segoe UI", 14, "bold"),
                                      text_color="white");
        self.lbl_title.pack(side="left", padx=8)
        self.btn_close = ctk.CTkButton(self.header, text="✕", width=25, height=25, fg_color="transparent",
                                       hover_color="#ff4444", font=("Arial", 12), command=self.close_app);
        self.btn_close.pack(side="right")
        self.btn_minimize = ctk.CTkButton(self.header, text="—", width=25, height=25, fg_color="transparent",
                                          hover_color="#333333", font=("Arial", 12, "bold"),
                                          command=self.animate_contraction);
        self.btn_minimize.pack(side="right", padx=2)

        self.body = ctk.CTkFrame(self.full_ui_frame, fg_color="transparent");
        self.body.pack(fill="both", expand=True, pady=10)
        self.avatar_display = ctk.CTkLabel(self.body, text="", image=self.images.get("idle"));
        self.avatar_display.pack(pady=(30, 10))
        self.lbl_status = ctk.CTkLabel(self.body, text="Ready.", font=("Segoe UI", 16), text_color="#a9b1d6");
        self.lbl_status.pack()

        # --- INPUT AREA (FIXED) ---
        self.footer = ctk.CTkFrame(self.full_ui_frame, fg_color="transparent")
        self.footer.pack(fill="x", padx=20, pady=20, side="bottom")

        self.input_bg = ctk.CTkFrame(self.footer, fg_color="black", corner_radius=20, height=45)
        self.input_bg.pack(fill="x")
        self.input_bg.pack_propagate(False)

        # 1. Text Entry
        self.entry = ctk.CTkEntry(self.input_bg, placeholder_text="Command...", border_width=0, fg_color="transparent",
                                  font=("Segoe UI", 13), text_color="white", placeholder_text_color="#565f89")
        self.entry.pack(side="left", fill="both", expand=True, padx=(15, 5))
        self.entry.bind("<Return>", self.start_execution)

        # 2. Mic Button
        self.btn_mic = ctk.CTkButton(self.input_bg, text="", image=self.mic_icon, width=35, height=35,
                                     fg_color="transparent", hover_color="#222222", corner_radius=18,
                                     command=self.start_voice_listening)
        self.btn_mic.pack(side="right", padx=0)

        # 3. Send Button (New) - Ensures visual feedback for sending
        self.btn_send = ctk.CTkButton(self.input_bg, text="➤", width=35, height=35, fg_color="transparent",
                                      hover_color="#222222", corner_radius=18, font=("Arial", 14),
                                      command=lambda: self.start_execution(None))
        self.btn_send.pack(side="right", padx=(0, 5))

    def animate_expansion(self):
        cx, cy = self.winfo_x(), self.winfo_y();
        self.btn_mini.place_forget();
        self.configure(fg_color=self.BG_COLOR);
        self.wm_attributes("-transparentcolor", "")
        self.geometry(f"320x450+{cx}+{cy}");
        self.full_ui_frame.pack(fill="both", expand=True);
        self.is_expanded = True;
        self.entry.focus_set()

    def animate_contraction(self):
        cx, cy = self.winfo_x(), self.winfo_y();
        self.full_ui_frame.pack_forget();
        self.configure(fg_color=self.TRANSPARENT_COLOR);
        self.wm_attributes("-transparentcolor", self.TRANSPARENT_COLOR)
        self.geometry(f"60x60+{cx}+{cy}");
        self.btn_mini.place(relx=0.5, rely=0.5, anchor="center");
        self.is_expanded = False

    def close_app(self):
        self.quit(); sys.exit()

    def set_expression(self, expr):
        if expr in self.images and self.images[expr]: self.avatar_display.configure(image=self.images[expr])

    def update_status_safe(self, msg):
        self.lbl_status.configure(text=msg)

    def start_voice_listening(self):
        if self.is_listening: return
        self.is_listening = True
        self.entry.delete(0, 'end')
        self.lbl_status.configure(text="Listening...", text_color="#f7768e")
        self.set_expression("listening")
        self.entry.configure(placeholder_text="Listening...")
        threading.Thread(target=self.bg_listen).start()

    def bg_listen(self):
        try:
            with sr.Microphone() as source:
                self.recognizer.adjust_for_ambient_noise(source, 0.5)
                audio = self.recognizer.listen(source, timeout=5)
            text = self.recognizer.recognize_google(audio)

            # Update UI safely
            self.after(0, lambda: self.entry.insert(0, text))
            self.after(0, lambda: self.start_execution(None))
        except:
            self.set_expression("error");
            self.update_status_safe("Didn't catch that.")
        finally:
            self.is_listening = False;
            self.after(0, lambda: self.entry.configure(placeholder_text="Command..."))

    def start_execution(self, event):
        txt = self.entry.get()
        if not txt: return

        # INSTANT CLEAR & RE-FOCUS
        self.entry.delete(0, 'end')
        self.entry.focus_set()

        self.set_expression("thinking")
        self.lbl_status.configure(text="Thinking...", text_color="#7aa2f7")
        threading.Thread(target=self.bg_execute, args=(txt,)).start()

    def bg_execute(self, txt):
        comtypes.CoInitialize()
        try:
            cmd, err = get_ai_command(txt)
            if cmd:
                self.update_status_safe(f"Seek: {cmd.button_name}")
                msg, success = self.automator.execute(cmd, self.ghost)
                self.update_status_safe(msg)
                self.set_expression("success" if success else "error")
            else:
                self.update_status_safe("Confused.")
                self.set_expression("error")
        finally:
            comtypes.CoUninitialize()
            time.sleep(1.5)
            self.set_expression("idle")
            self.update_status_safe("Ready.")


if __name__ == "__main__":
    app = WorkerApp()


    def start_move(e):
        app.x = e.x; app.y = e.y


    def stop_move(e):
        app.x = None; app.y = None


    def do_move(e):
        try:
            app.geometry(f"+{app.winfo_x() + e.x - app.x}+{app.winfo_y() + e.y - app.y}")
        except:
            pass


    for w in [app.header, app.lbl_logo, app.lbl_title, app.btn_mini, app.body, app.avatar_display]:
        w.bind("<ButtonPress-1>", start_move);
        w.bind("<ButtonRelease-1>", stop_move);
        w.bind("<B1-Motion>", do_move)
    app.mainloop()