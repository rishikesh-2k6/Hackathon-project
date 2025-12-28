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
from typing import List

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
MY_API_KEY = "AIzaSyCwzywKhx91pyv8YxBjh2OYsDTAQZtYeVc"

# ==========================================
# UI COMPONENTS
# ==========================================
class ListeningWave:
    def __init__(self, parent, colors):
        self.canvas = tk.Canvas(parent, width=150, height=20, bg="#1a1b26", highlightthickness=0)
        self.colors = colors
        self.dots = []
        self.angle = 0
        self.is_animating = False
        for i in range(4):
            dot = self.canvas.create_oval(35 + i * 22, 5, 47 + i * 22, 17, fill=self.colors[i], outline="")
            self.dots.append(dot)

    def start(self):
        self.is_animating = True
        self.canvas.pack(pady=(0, 2))
        self.animate()

    def stop(self):
        self.is_animating = False
        self.canvas.pack_forget()

    def animate(self):
        if not self.is_animating: return
        for i, dot in enumerate(self.dots):
            y_offset = math.sin(self.angle + (i * 0.9)) * 4
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
        url = f"https://generativelanguage.googleapis.com/v1beta/models?key={MY_API_KEY}"
        response = requests.get(url)
        data = response.json()
        for m in data.get('models', []):
            if 'generateContent' in m.get('supportedGenerationMethods', []):
                name = m['name'].replace("models/", "")
                if "flash" in name and "8b" not in name: return name
    except: pass
    return "gemini-1.5-flash"

CURRENT_MODEL = get_working_model()

# ==========================================
# AI CLIENT (SEQUENTIAL MULTI-TASKING)
# ==========================================
class PPTCommand(BaseModel):
    tab_name: str
    button_name: str
    explanation: str

def get_ai_commands(user_text, current_tab="Unknown"):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{CURRENT_MODEL}:generateContent?key={MY_API_KEY}"
    headers = {'Content-Type': 'application/json'}
    prompt = f"""You are a PowerPoint Instructor. User may provide multiple tasks (e.g., 'add a slide and change background').
Break these into a sequence of Ribbon button clicks.

CONTEXT: User is at '{current_tab}' tab.
REQ: "{user_text}"

RETURN JSON LIST ONLY: [{{ "tab_name": "Home", "button_name": "New Slide", "explanation": "Create a slide" }},{{ "tab_name": "Design", "button_name": "Format Background", "explanation": "Change the background" }}]"""
    
    data = { "contents": [{ "parts": [{"text": prompt}] }] }
    try:
        response = requests.post(url, headers=headers, json=data)
        if response.status_code == 200:
            raw = response.json()['candidates'][0]['content']['parts'][0]['text']
            clean = raw.replace("```json", "").replace("```", "").strip()
            parsed = json.loads(clean)
            if isinstance(parsed, dict): parsed = [parsed]
            return [PPTCommand(**d) for d in parsed], None
        return None, "API Error"
    except: return None, "Connection Failed"

def translate_text_api(text, target_lang):
    if not text.strip(): return ""
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{CURRENT_MODEL}:generateContent?key={MY_API_KEY}"
    headers = {'Content-Type': 'application/json'}
    prompt = f"Translate to {target_lang}. Return ONLY translated text: '{text}'"
    data = { "contents": [{ "parts": [{"text": prompt}] }] }
    try:
        response = requests.post(url, headers=headers, json=data)
        if response.status_code == 200:
            return response.json()['candidates'][0]['content']['parts'][0]['text'].strip()
    except: pass
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
            steps = 80
            for i in range(steps):
                if self.stop_signal.is_set(): break
                t_val = 1 - (1 - (i+1)/steps)**2
                cx = int(mx + (tx - mx) * t_val); cy = int(my + (ty - my) * t_val)
                canvas.delete("all")
                pts = [cx, cy, cx, cy+26, cx+7, cy+20, cx+11, cy+29, cx+16, cy+27, cx+12, cy+18, cx+21, cy+18]
                canvas.create_polygon(pts, outline="#2c2987", fill="#2c2987", width=2)
                self._draw_tooltip(canvas, cx, cy, text)
                root.update(); time.sleep(0.02)
            while not self.stop_signal.is_set():
                canvas.delete("all")
                pts = [tx, ty, tx, ty+26, tx+7, ty+20, tx+11, ty+29, tx+16, ty+27, tx+12, ty+18, tx+21, ty+18]
                canvas.create_polygon(pts, outline="#2c2987", fill="#2c2987", width=2)
                self._draw_tooltip(canvas, tx, ty, text)
                root.update(); time.sleep(0.01)
            root.destroy()
        except: pass
        finally: comtypes.CoUninitialize()

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
        if ppt.Exists(0, 0): return ppt
        return None

    def _wait_for_click(self, target_control):
        rect = target_control.BoundingRectangle
        while not self.stop_event.is_set():
            if ctypes.windll.user32.GetKeyState(0x01) < 0:
                mx, my = auto.GetCursorPos()
                if rect.left <= mx <= rect.right and rect.top <= my <= rect.bottom: return True
            time.sleep(0.1)
        return False

    def execute_single_step(self, cmd, ghost, ppt):
        if self.stop_event.is_set(): return "Stopped", False
        # 1. Tab
        self.update_status(f"Go to: {cmd.tab_name}")
        tab = ppt.TabItemControl(Name=cmd.tab_name, searchDepth=10)
        if not tab.Exists(0, 0): tab = ppt.TabItemControl(RegexName=f"(?i).*{cmd.tab_name}.*", searchDepth=10)
        if tab.Exists(0, 1):
            r = tab.BoundingRectangle
            ghost.point_at((r.left+r.right)//2, (r.top+r.bottom)//2, f"Click '{cmd.tab_name}'")
            if not self._wait_for_click(tab): return "Stopped", False
            ghost.hide()
            time.sleep(0.5)
        # 2. Button
        self.update_status(f"Target: {cmd.button_name}")
        btn = ppt.Control(Name=cmd.button_name, searchDepth=15)
        if not btn.Exists(0, 0): btn = ppt.Control(RegexName=f"(?i).*{cmd.button_name}.*", searchDepth=15)
        if btn.Exists(0, 1):
            r = btn.BoundingRectangle
            ghost.point_at((r.left+r.right)//2, (r.top+r.bottom)//2, f"Click '{cmd.button_name}'")
            if not self._wait_for_click(btn): return "Stopped", False
            ghost.hide()
            return f"Step '{cmd.button_name}' Done", True
        ghost.hide()
        return f"❌ Missing: {cmd.button_name}", False

if __name__ == "__main__":
    print("Instructly AI - Automator Added")