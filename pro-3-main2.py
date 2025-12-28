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

if __name__ == "__main__":
    print("Instructly AI - AI Client Added")