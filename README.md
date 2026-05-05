# AutoDeck AI — Presentation Analyzer

Built for the INICAI AI Multi-Domain Hackathon.

## What it does
Upload any PowerPoint file (.ppt, .pptx, .pps, .pot) and instantly get:
- **Complexity** — Low / Medium / High
- **Density** — Light / Medium / Dense  
- **Category** — Legacy / Modern / Slideshow / Template
- **Confidence score** from ML model

## How it works
- Extracts features from PowerPoint files (slide count, word count, file size)
- Random Forest ML model trained on 698 real government presentations
- Falls back to calibrated size-based estimation for legacy .ppt files

## Tech Stack
- Python, Gradio, scikit-learn, python-pptx
- Trained on US Government PowerPoint archive dataset