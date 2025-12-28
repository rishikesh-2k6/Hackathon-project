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

if __name__ == "__main__":
    print("Instructly AI - UI Components Added")