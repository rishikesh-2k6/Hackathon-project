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

if __name__ == "__main__":
    print("Instructly AI - Initial Setup Complete")