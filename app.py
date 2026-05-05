import os
import pickle
import gradio as gr
from pptx import Presentation
import tempfile

# Load models
with open('models.pkl', 'rb') as f:
    models = pickle.load(f)

clf_complexity = models['clf_complexity']
clf_density = models['clf_density']
le_complexity = models['le_complexity']
le_density = models['le_density']

def extract_features(filepath, file_extension):
    try:
        if file_extension.lower() in ['.pptx', '.ppsx', '.pptm']:
            prs = Presentation(filepath)
            slide_count = len(prs.slides)
            word_count = 0
            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text"):
                        word_count += len(shape.text.split())
            return slide_count, word_count, True
        else:
            size_kb = os.path.getsize(filepath) / 1024
            if size_kb < 297:
                estimated_slides = max(1, int(size_kb / 30))
                words_per_slide = 20
            elif size_kb < 944:
                estimated_slides = max(1, int(size_kb / 80))
                words_per_slide = 45
            elif size_kb < 3136:
                estimated_slides = max(1, int(size_kb / 120))
                words_per_slide = 65
            else:
                estimated_slides = max(1, int(size_kb / 200))
                words_per_slide = 90
            return estimated_slides, estimated_slides * words_per_slide, False
    except Exception:
        size_kb = os.path.getsize(filepath) / 1024
        estimated_slides = max(1, int(size_kb / 100))
        return estimated_slides, estimated_slides * 45, False

def get_category(ext):
    ext = ext.lower()
    if ext == '.pptx': return "Modern"
    elif ext == '.ppt': return "Legacy"
    elif ext in ['.pps', '.ppsx']: return "Slideshow"
    elif ext in ['.pot', '.potx']: return "Template"
    else: return "Modern"

def analyze_presentation(file):
    if file is None:
        return "Please upload a file.", "", "", "", "", ""

    filepath = file.name
    filename = os.path.basename(filepath)
    ext = os.path.splitext(filename)[1]
    size_kb = os.path.getsize(filepath) / 1024
    size_mb = round(size_kb / 1024, 2)

    slide_count, word_count, exact = extract_features(filepath, ext)
    avg_words = word_count / max(slide_count, 1)

    X = [[size_kb, size_mb, slide_count, word_count, avg_words]]

    complexity = le_complexity.inverse_transform(clf_complexity.predict(X))[0]
    density = le_density.inverse_transform(clf_density.predict(X))[0]
    category = get_category(ext)

    complexity_proba = clf_complexity.predict_proba(X).max()
    density_proba = clf_density.predict_proba(X).max()
    confidence = round((complexity_proba + density_proba) / 2 * 100, 1)

    method = "Exact extraction" if exact else "Estimated from file size"

    summary = f"""
📊 File Analysis Summary
━━━━━━━━━━━━━━━━━━━━━━━━
📁 Filename     : {filename}
📦 File Size    : {size_mb} MB ({size_kb:.1f} KB)
🔖 Extension    : {ext}
📑 Slides       : {slide_count}
📝 Words        : {word_count}
📐 Avg Words/Slide: {avg_words:.1f}
🔍 Method       : {method}
━━━━━━━━━━━━━━━━━━━━━━━━
🧠 Predictions
━━━━━━━━━━━━━━━━━━━━━━━━
⚡ Complexity   : {complexity}
🌊 Density      : {density}
🗂️  Category     : {category}
✅ Confidence   : {confidence}%
━━━━━━━━━━━━━━━━━━━━━━━━
    """

    return summary, complexity, density, category, f"{confidence}%", f"{slide_count} slides / {word_count} words"

# Build Gradio UI
with gr.Blocks(title="AutoDeck AI Analyzer") as demo:
    
    gr.Markdown("""
    # 🎯 AutoDeck AI — Presentation Analyzer
    ### Powered by Random Forest ML Model
    Upload any PowerPoint file (.ppt, .pptx, .pps, .pot) to instantly analyze its complexity, density, and category.
    """)

    with gr.Row():
        with gr.Column(scale=1):
            file_input = gr.File(
                label="📂 Upload PowerPoint File",
                file_types=[".ppt", ".pptx", ".pps", ".ppsx", ".pot", ".pptm"]
            )
            analyze_btn = gr.Button("🔍 Analyze", variant="primary", size="lg")

        with gr.Column(scale=2):
            summary_output = gr.Textbox(
                label="📋 Full Analysis",
                lines=18,
                interactive=False
            )

    with gr.Row():
        complexity_out = gr.Label(label="⚡ Complexity")
        density_out = gr.Label(label="🌊 Density")
        category_out = gr.Label(label="🗂️ Category")
        confidence_out = gr.Textbox(label="✅ Confidence", interactive=False)
        stats_out = gr.Textbox(label="📊 Stats", interactive=False)

    analyze_btn.click(
        fn=analyze_presentation,
        inputs=[file_input],
        outputs=[summary_output, complexity_out, density_out, 
                 category_out, confidence_out, stats_out]
    )

    gr.Markdown("""
    ---
    **Built for INICAI AI Multi-Domain Hackathon** | AutoDeck AI Track
    """)

demo.launch(share=False)