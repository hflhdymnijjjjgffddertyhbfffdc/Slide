PPT Generator with Audio Narration
Overview
This project consists of two main Python scripts that work together to:

Convert PDF research papers into structured PowerPoint presentations

Add audio narration to existing PowerPoint presentations
1. Research Paper to PPT Converter (main.py)
Extracts structured information from research papers (title, authors, abstract, etc.)

Generates professional PowerPoint slides with:

Clean, consistent formatting

Properly organized content sections

Images extracted from the paper with captions

Speaker notes in a news broadcast style

Handles long titles by automatically splitting them

Processes images with proper wrapping and formatting
2. PPT to Video with Narration (ppt_presenter.py)
Converts PowerPoint slides to video format

Extracts text from slides (both notes and slide content)

Generates audio narration using a local TTS service

Combines slides and audio into a final MP4 video

Includes error handling for missing content
Installation
Clone this repository

Install dependencies:
pip install -r requirements.txt
Usage
1. For Research Paper to PPT Conversion
python main.py
Configuration:

Set input PDF directory in dir_pdf variable

Output will be saved to dir_ppt directory
2. For PPT to Video with Narration
python ppt_presenter.py
Configuration
Edit these variables in the scripts:

API keys for GPT services

Directory paths for input/output files

TTS service URL (default: http://127.0.0.1:9880)

PDF conversion service URL (default: http://127.0.0.1:8000/upload-pdf/)
