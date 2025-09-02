#!/bin/bash

# Teams Message Fetcher - Streamlit App Launcher
echo "🚀 Starting Teams Message Fetcher..."

# Check if virtual environment exists
if [ ! -d "venv" ]; then
    echo "📦 Creating virtual environment..."
    python3 -m venv venv
fi

# Activate virtual environment
echo "🔧 Activating virtual environment..."
source venv/bin/activate

# Install/upgrade requirements
echo "📥 Installing requirements..."
pip install -r requirements.txt --quiet

# Check if streamlit is available
if ! command -v streamlit &> /dev/null; then
    echo "❌ Streamlit installation failed. Please check your Python setup."
    exit 1
fi

# Launch the app
echo "🌐 Launching Streamlit app..."
echo "📍 Open your browser to: http://localhost:8501"
echo "🛑 Press Ctrl+C to stop the server"
echo ""

streamlit run streamlit_app.py