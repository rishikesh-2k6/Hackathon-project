
## ✨ Features
- **Real-time on-screen guidance** with animated cursor
- **Dynamic pointers** for precise navigation
- **Voice commands** - speak your requests naturally
- **AI-powered command interpretation** using Google Gemini
- **PowerPoint automation** - works with Microsoft PowerPoint
- **Transparent floating UI** - stays out of your way
- **No need to switch** between tutorial videos and applications
- **Interactive learning experience** with visual feedback

## 🚀 Getting Started

### Prerequisites
- Python 3.8 or higher
- Microsoft PowerPoint installed
- Microphone (for voice commands)
- Windows OS (required for UI automation)

### Installation

1. **Clone this repository**
   ```bash
   git clone https://github.com/rishikesh-2k6/Hackathon-project.git
   cd Hackathon-project
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Set up environment variables** (optional)
   ```bash
   cp .env.example .env
   # Edit .env with your Google AI API key if needed
   ```

4. **Run the application**
   ```bash
   python pro-3-main2.py
   ```

## 🎯 How to Use

1. **Launch the app** - A small robot icon will appear on your screen
2. **Click the robot** to expand the interface
3. **Open PowerPoint** in the background
4. **Give commands** either by:
   - Typing in the text box (e.g., "change font", "add new slide")
   - Using voice commands (click the microphone button)
5. **Watch the magic** - The app will guide you with animated pointers

## 🎤 Voice Commands Examples
- "Change font"
- "Make it bold"
- "Add new slide"
- "Insert picture"
- "Start slideshow"
- "Add text box"

## 🛠️ Technical Details

### Built With
- **CustomTkinter** - Modern GUI framework
- **Google Gemini AI** - Natural language processing
- **Windows UI Automation** - PowerPoint integration
- **Speech Recognition** - Voice command processing
- **PIL/Pillow** - Image processing for UI elements

### Architecture
- **Transparent overlay system** for visual guidance
- **Multi-threaded design** for responsive UI
- **AI-powered command mapping** with offline fallbacks
- **Ghost cursor animation** for precise pointing

## 🤝 Contributing

We welcome contributions! Here's how to get started:

1. **Fork the repository**
2. **Create a feature branch** (`git checkout -b feature/amazing-feature`)
3. **Make your changes** and commit (`git commit -m 'Add amazing feature'`)
4. **Push to your branch** (`git push origin feature/amazing-feature`)
5. **Open a Pull Request**

### Development Setup
```bash
# Clone your fork
git clone https://github.com/YOUR-USERNAME/Hackathon-project.git

# Create virtual environment (recommended)
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

## 📁 Project Structure
```
Hackathon-project/
├── pro-3-main2.py          # Main application file
├── requirements.txt        # Python dependencies
├── README.md              # This file
├── .env.example           # Environment variables template
├── commands.json          # Command mappings
├── robo_*.png            # Robot avatar images
├── mic.png               # Microphone icon
├── start_icon.png        # App launcher icon
└── logo.ico              # Application icon
```

## 🐛 Troubleshooting

**PowerPoint not detected?**
- Make sure PowerPoint is running and visible
- Try clicking on PowerPoint window first

**Voice recognition not working?**
- Check microphone permissions
- Ensure you have a stable internet connection
- Try typing commands instead

**App not responding?**
- Close and restart the application
- Check if all dependencies are installed correctly

## 📄 License
This project is open source and available under the [MIT License](LICENSE).

## 👥 Team
- **Rishikesh** - Project Lead & Developer

---
*Built with ❤️ for making software learning more interactive and intuitive!*
