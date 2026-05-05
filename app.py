import os
import pickle
import gradio as gr
from pptx import Presentation

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
        return """
<div style='text-align:center; padding: 60px 20px; color: #4a5a4a; font-family: monospace; font-size: 15px;'>
    [ AWAITING FILE INPUT ]<br><br>
    <span style='font-size:12px; color: #3a4a3a;'>Upload a PowerPoint file to initialize analysis</span>
</div>
"""

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

    method = "DIRECT EXTRACTION" if exact else "SIZE-BASED ESTIMATION"

    complexity_colors = {"Low": "#22c55e", "Medium": "#f59e0b", "High": "#ef4444"}
    density_colors = {"Light": "#22c55e", "Medium": "#f59e0b", "Dense": "#ef4444"}
    bar_color = "#22c55e" if confidence >= 80 else "#f59e0b" if confidence >= 60 else "#ef4444"

    cc = complexity_colors.get(complexity, "#22c55e")
    dc = density_colors.get(density, "#22c55e")
    bar_fill = int(confidence)

    html = f"""
<div style='font-family: "Courier New", monospace; background: #0a0f0a; 
            border: 1px solid #1a2e1a; border-radius: 12px; padding: 24px; color: #a0c4a0;'>

    <!-- Header -->
    <div style='border-bottom: 1px solid #1a3a1a; padding-bottom: 14px; margin-bottom: 20px;'>
        <div style='display: flex; justify-content: space-between; align-items: center;'>
            <div>
                <span style='color: #4a7a4a; font-size: 11px; letter-spacing: 0.15em;'>
                    ▶ AUTODECK INTEL SYSTEM v1.0
                </span><br>
                <span style='color: #22c55e; font-size: 16px; font-weight: bold;'>
                    ANALYSIS COMPLETE
                </span>
            </div>
            <div style='text-align: right;'>
                <span style='background: #22c55e; color: #000; font-size: 10px; 
                             font-weight: bold; padding: 3px 10px; border-radius: 4px;
                             letter-spacing: 0.1em;'>● ONLINE</span>
            </div>
        </div>
    </div>

    <!-- File Intel -->
    <div style='background: #0d150d; border: 1px solid #1a3a1a; border-radius: 8px; 
                padding: 16px; margin-bottom: 16px;'>
        <div style='color: #4a7a4a; font-size: 10px; letter-spacing: 0.15em; margin-bottom: 10px;'>
            ◈ FILE INTELLIGENCE
        </div>
        <div style='color: #e0ffe0; font-size: 15px; font-weight: bold; 
                    margin-bottom: 12px; word-break: break-all;'>
            {filename}
        </div>
        <div style='display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px;'>
            <div style='background: #111a11; border: 1px solid #1a3a1a; 
                        border-radius: 6px; padding: 10px; text-align: center;'>
                <div style='color: #4a7a4a; font-size: 9px; letter-spacing: 0.1em;'>SIZE</div>
                <div style='color: #22c55e; font-size: 16px; font-weight: bold;'>{size_mb}</div>
                <div style='color: #4a7a4a; font-size: 9px;'>MB</div>
            </div>
            <div style='background: #111a11; border: 1px solid #1a3a1a; 
                        border-radius: 6px; padding: 10px; text-align: center;'>
                <div style='color: #4a7a4a; font-size: 9px; letter-spacing: 0.1em;'>SLIDES</div>
                <div style='color: #22c55e; font-size: 16px; font-weight: bold;'>{slide_count}</div>
                <div style='color: #4a7a4a; font-size: 9px;'>COUNT</div>
            </div>
            <div style='background: #111a11; border: 1px solid #1a3a1a; 
                        border-radius: 6px; padding: 10px; text-align: center;'>
                <div style='color: #4a7a4a; font-size: 9px; letter-spacing: 0.1em;'>WORDS</div>
                <div style='color: #22c55e; font-size: 16px; font-weight: bold;'>{word_count}</div>
                <div style='color: #4a7a4a; font-size: 9px;'>TOTAL</div>
            </div>
            <div style='background: #111a11; border: 1px solid #1a3a1a; 
                        border-radius: 6px; padding: 10px; text-align: center;'>
                <div style='color: #4a7a4a; font-size: 9px; letter-spacing: 0.1em;'>AVG W/S</div>
                <div style='color: #22c55e; font-size: 16px; font-weight: bold;'>{avg_words:.0f}</div>
                <div style='color: #4a7a4a; font-size: 9px;'>PER SLIDE</div>
            </div>
        </div>
        <div style='margin-top: 10px; font-size: 10px; color: #3a6a3a; letter-spacing: 0.1em;'>
            ◈ EXTRACTION METHOD: {method}
        </div>
    </div>

    <!-- Threat Assessment / Predictions -->
    <div style='display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; margin-bottom: 16px;'>

        <div style='background: #0d150d; border: 1px solid {cc}44; 
                    border-radius: 8px; padding: 16px; text-align: center;
                    box-shadow: 0 0 12px {cc}22;'>
            <div style='color: #4a7a4a; font-size: 9px; letter-spacing: 0.15em; margin-bottom: 8px;'>
                ◈ COMPLEXITY LEVEL
            </div>
            <div style='font-size: 26px; font-weight: 900; color: {cc}; 
                        letter-spacing: 0.05em;'>{complexity.upper()}</div>
            <div style='margin-top: 8px; height: 3px; background: #1a2e1a; border-radius: 2px;'>
                <div style='height: 100%; border-radius: 2px; background: {cc};
                    width: {"33%" if complexity=="Low" else "66%" if complexity=="Medium" else "100%"};'>
                </div>
            </div>
        </div>

        <div style='background: #0d150d; border: 1px solid {dc}44; 
                    border-radius: 8px; padding: 16px; text-align: center;
                    box-shadow: 0 0 12px {dc}22;'>
            <div style='color: #4a7a4a; font-size: 9px; letter-spacing: 0.15em; margin-bottom: 8px;'>
                ◈ CONTENT DENSITY
            </div>
            <div style='font-size: 26px; font-weight: 900; color: {dc}; 
                        letter-spacing: 0.05em;'>{density.upper()}</div>
            <div style='margin-top: 8px; height: 3px; background: #1a2e1a; border-radius: 2px;'>
                <div style='height: 100%; border-radius: 2px; background: {dc};
                    width: {"33%" if density=="Light" else "66%" if density=="Medium" else "100%"};'>
                </div>
            </div>
        </div>

        <div style='background: #0d150d; border: 1px solid #22c55e44; 
                    border-radius: 8px; padding: 16px; text-align: center;
                    box-shadow: 0 0 12px #22c55e22;'>
            <div style='color: #4a7a4a; font-size: 9px; letter-spacing: 0.15em; margin-bottom: 8px;'>
                ◈ FILE CATEGORY
            </div>
            <div style='font-size: 26px; font-weight: 900; color: #22c55e; 
                        letter-spacing: 0.05em;'>{category.upper()}</div>
            <div style='margin-top: 8px; font-size: 10px; color: #3a6a3a;'>{ext.upper()}</div>
        </div>

    </div>

    <!-- Confidence -->
    <div style='background: #0d150d; border: 1px solid #1a3a1a; 
                border-radius: 8px; padding: 16px;'>
        <div style='display: flex; justify-content: space-between; margin-bottom: 10px;'>
            <span style='color: #4a7a4a; font-size: 10px; letter-spacing: 0.15em;'>
                ◈ MODEL CONFIDENCE RATING
            </span>
            <span style='color: {bar_color}; font-size: 18px; font-weight: 900;'>
                {confidence}%
            </span>
        </div>
        <div style='background: #111a11; border-radius: 999px; height: 8px; 
                    overflow: hidden; border: 1px solid #1a3a1a;'>
            <div style='width: {bar_fill}%; height: 100%; background: {bar_color}; 
                        border-radius: 999px;'></div>
        </div>
        <div style='display: flex; justify-content: space-between; 
                    margin-top: 6px; font-size: 9px; color: #3a5a3a;'>
            <span>0%</span>
            <span>RANDOM FOREST · 100 ESTIMATORS · TRAINED ON 698 FILES</span>
            <span>100%</span>
        </div>
    </div>

</div>
"""
    return html


CSS = """
.gradio-container {
    background: #060d06 !important;
    max-width: 900px !important;
    margin: 0 auto !important;
}
footer { display: none !important; }
.gr-button-primary {
    background: #22c55e !important;
    color: #000 !important;
    font-weight: 800 !important;
    font-family: 'Courier New', monospace !important;
    letter-spacing: 0.1em !important;
    border: none !important;
}
.gr-button-primary:hover {
    background: #16a34a !important;
}
"""

with gr.Blocks(title="AutoDeck AI", css=CSS) as demo:

    gr.HTML("""
    <div style='font-family: "Courier New", monospace; text-align: center; 
                padding: 36px 0 28px; background: #060d06;'>
        <div style='font-size: 11px; color: #3a6a3a; letter-spacing: 0.3em; margin-bottom: 8px;'>
            ▶ INICAI HACKATHON · AUTODECK TRACK
        </div>
        <div style='font-size: 36px; font-weight: 900; color: #22c55e; 
                    letter-spacing: 0.1em; text-shadow: 0 0 30px #22c55e66;'>
            AUTODECK AI
        </div>
        <div style='font-size: 13px; color: #4a7a4a; letter-spacing: 0.2em; margin-top: 6px;'>
            PRESENTATION INTELLIGENCE SYSTEM
        </div>
        <div style='width: 60px; height: 2px; background: #22c55e; 
                    margin: 16px auto 0; border-radius: 2px;'></div>
    </div>
    """)

    with gr.Row():
        with gr.Column(scale=1):
            gr.HTML("""
            <div style='font-family: "Courier New", monospace; background: #0a0f0a;
                        border: 1px solid #1a3a1a; border-radius: 10px; padding: 16px;
                        margin-bottom: 12px;'>
                <div style='color: #4a7a4a; font-size: 10px; letter-spacing: 0.15em; margin-bottom: 10px;'>
                    ◈ MISSION BRIEF
                </div>
                <div style='color: #a0c4a0; font-size: 12px; line-height: 1.8;'>
                    Upload any PowerPoint file for instant AI-powered analysis.<br><br>
                    The system will classify:<br>
                    <span style='color: #22c55e;'>▸</span> Complexity Level<br>
                    <span style='color: #22c55e;'>▸</span> Content Density<br>
                    <span style='color: #22c55e;'>▸</span> File Category<br>
                    <span style='color: #22c55e;'>▸</span> Confidence Score
                </div>
            </div>
            """)

            file_input = gr.File(
                label="◈ UPLOAD TARGET FILE",
                file_types=[".ppt", ".pptx", ".pps", ".ppsx", ".pot", ".pptm"]
            )
            analyze_btn = gr.Button(
                "▶ INITIALIZE ANALYSIS",
                variant="primary",
                size="lg"
            )
            gr.HTML("""
            <div style='font-family: "Courier New", monospace; font-size: 10px; 
                        color: #3a5a3a; text-align: center; margin-top: 8px; letter-spacing: 0.1em;'>
                SUPPORTED: .PPT .PPTX .PPS .PPSX .POT .PPTM
            </div>
            """)

        with gr.Column(scale=2):
            output = gr.HTML("""
            <div style='font-family: "Courier New", monospace; background: #0a0f0a;
                        border: 1px solid #1a3a1a; border-radius: 12px; padding: 40px 20px;
                        text-align: center; color: #3a6a3a; min-height: 300px;
                        display: flex; flex-direction: column; 
                        align-items: center; justify-content: center;'>
                <div style='font-size: 32px; margin-bottom: 16px;'>⬡</div>
                <div style='font-size: 13px; letter-spacing: 0.2em; color: #2a5a2a;'>
                    SYSTEM STANDBY
                </div>
                <div style='font-size: 11px; color: #1a3a1a; margin-top: 8px; letter-spacing: 0.1em;'>
                    AWAITING FILE INPUT
                </div>
            </div>
            """)

    analyze_btn.click(
        fn=analyze_presentation,
        inputs=[file_input],
        outputs=[output]
    )

    gr.HTML("""
    <div style='font-family: "Courier New", monospace; text-align: center; 
                padding: 20px 0 8px; font-size: 10px; color: #2a4a2a; 
                letter-spacing: 0.15em; border-top: 1px solid #0d1a0d; margin-top: 20px;'>
        AUTODECK AI · RANDOM FOREST ENGINE · INICAI HACKATHON 2026 · ALL SYSTEMS NOMINAL
    </div>
    """)

demo.launch()