# AI Scene Detection & Thumbnail Extractor

A high-performance scene detection tool powered by **TransNetV2** (Deep Learning). It automatically detects cuts, fades, and dissolves, generating a professional shot list with visual references.

## Features
- **AI-Powered:** Uses a 3D Convolutional Neural Network for state-of-the-art accuracy.
- **GPU Accelerated:** Automatically uses NVIDIA GPUs (via CUDA) for lightning-fast processing.
- **Stable UHD Workflow:** Automatically creates a low-res proxy for analysis to prevent crashes on 4K/UHD footage.
- **Visual Deliverables:**
  - **Interactive HTML Shot Sheet:** Checkboxes for VFX tracking and description fields for notes.
  - **Visual Excel (XLSX):** Embedded thumbnails, shot codes, and frame counts.
  - **CSV Data:** Clean structured data for any spreadsheet app.
  - **UHD Thumbnails:** Extracted directly from the source at the start of every shot.

## Installation

### 1. Prerequisites
- **Python 3.10 or 3.11** (Recommended for best CUDA compatibility)
- **FFmpeg** installed and added to your system PATH.
- **NVIDIA GPU** (Optional, but highly recommended for 10x speed).

### 2. Setup Environment
```bash
# Create a virtual environment
python -m venv venv

# Activate it (Windows)
.\venv\Scripts\activate

# Activate it (Mac/Linux)
source venv/bin/activate
```

### 3. Install Dependencies
```bash
# Install PyTorch with CUDA support (for Windows/Linux with NVIDIA GPU)
pip install torch torchvision --index-url https://download.pytorch.org/whl/cu121

# Install other requirements
pip install -r requirements.txt
```

## Usage

Run the script by pointing it to your video file:

```bash
python detect_cuts.py "C:/path/to/your/video.mp4"
```

### Options
- `--output`: Specify where to save the results (defaults to the current folder).

## Output Structure
The tool creates a folder named `{video_name}_scenedetect` containing:
- `thumbnails/`: Folder containing all high-res shot references (`sh_0010.jpg`, etc.).
- `Shot_Sheet.html`: Interactive visual guide for browsers.
- `Shot_List_Visual.xlsx`: Portable Excel sheet with images.
- `shot_list.csv`: Raw data.

## Credits
This tool uses the **TransNetV2** model for shot boundary detection.
