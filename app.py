import os, json, re, io
from typing import Optional
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from openai import OpenAI

APP_DIR = os.path.dirname(os.path.abspath(__file__))
app = FastAPI(title="SlideCraft AI", version="0.1.0")
app.mount("/static", StaticFiles(directory=os.path.join(APP_DIR, "static")), name="static")

@app.get("/", response_class=HTMLResponse)
def home():
    return FileResponse(os.path.join(APP_DIR, "static", "index.html"))

@app.get("/api/status")
def status():
    return {"ok": True, "ai_configured": bool(os.getenv("OPENAI_API_KEY"))}

def extract_text_from_pptx(prs):
    slides = []
    for idx, slide in enumerate(prs.slides, start=1):
        bits = []
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False):
                t = shape.text.strip()
                if t:
                    bits.append(t[:300])
        slides.append({"slide": idx, "texts": bits})
    return slides

def fallback_outline(topic, audience, n, source_text):
    # Explicitly label this as a draft: no external facts or citations are fabricated.
    slides = []
    base_titles = [
        "Pengantar dan konteks", "Tujuan pembahasan", "Konsep utama",
        "Penjelasan inti", "Contoh atau penerapan", "Analisis dan pembahasan",
        "Hal yang perlu diperhatikan", "Ringkasan", "Kesimpulan", "Referensi"
    ]
    for i in range(n):
        title = base_titles[i] if i < len(base_titles) else f"Pembahasan {i+1}"
        body = (
            "Lengkapi bagian ini menggunakan materi atau sumber yang sudah diverifikasi."
            if not source_text else
            "Susun poin inti berdasarkan materi sumber yang diunggah. Periksa kembali kesesuaian isi dengan sumber asli."
        )
        slides.append({"title": title, "body": body, "notes": "Draf awal — verifikasi isi sebelum digunakan."})
    slides[0]["title"] = topic[:90] if topic.strip() else "Presentasi"
    return {"title": topic or "Presentasi", "slides": slides, "references": []}

def make_ai_outline(topic, audience, n, source_text, style, extra_instructions):
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        return fallback_outline(topic, audience, n, source_text), "fallback"
    client = OpenAI(api_key=api_key)
    system = """You are an academic presentation editor. Return only valid JSON with keys:
title, slides (array of objects with title, body, visual_suggestion, speaker_note), references (array).
Each slide body should be concise, natural Indonesian, usually 2-4 bullets separated by newline.
Do not invent facts, data, quotations, or references. Use only claims supported by supplied source text.
If the source does not support a claim, write [PERLU SUMBER] or [PERLU DIISI], rather than guessing.
References must contain only bibliographic details explicitly present in source text; otherwise return [].
Avoid generic AI-sounding phrasing and repetitive slide structures. Vary the rhetorical purpose of slides.
For formulas, numbers, and technical claims, preserve source wording and values exactly."""
    user = f"""Topik: {topic}
Audiens: {audience}
Jumlah slide: {n}
Gaya visual: {style}
Instruksi tambahan: {extra_instructions}
Materi sumber (anggap sebagai sumber utama, jangan melampaui isinya):
{source_text[:24000]}

Buat alur presentasi yang logis dan cocok untuk mahasiswa/dosen. Jika sumber tidak cukup, tandai bagian yang perlu dilengkapi.
Format JSON:
{{"title":"...", "slides":[{{"title":"...", "body":"poin 1\\npoin 2", "visual_suggestion":"...", "speaker_note":"..."}}], "references":["..."]}}"""
    try:
        response = client.chat.completions.create(
            model=os.getenv("OPENAI_MODEL", "gpt-4o-mini"),
            temperature=0.45,
            response_format={"type": "json_object"},
            messages=[{"role":"system","content":system},{"role":"user","content":user}],
        )
        data = json.loads(response.choices[0].message.content)
        slides = data.get("slides", [])
        if not slides:
            raise ValueError("Model returned no slides")
        data["slides"] = slides[:n]
        while len(data["slides"]) < n:
            data["slides"].append({"title":"Bagian tambahan", "body":"[PERLU DIISI]", "visual_suggestion":"", "speaker_note":""})
        return data, "openai"
    except Exception as e:
        # Do not silently imply AI ran successfully.
        raise HTTPException(status_code=502, detail=f"AI generation failed: {str(e)[:250]}")

def replace_tokens_in_shape(shape, mapping):
    if not getattr(shape, "has_text_frame", False):
        return
    # Replace tokens in each paragraph/run while preserving as much formatting as possible.
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            text = run.text
            for token, value in mapping.items():
                text = text.replace(token, value)
            run.text = text

def apply_content_to_template(prs, outline):
    slides = outline.get("slides", [])
    if not prs.slides:
        raise ValueError("Template has no slides")
    # Replace named tokens wherever they occur. Tokens: {{TITLE}}, {{BODY}}, {{NOTES}}, {{VISUAL}}, {{REFERENCES}}.
    # Slides beyond the template are not fabricated/cloned; report the mismatch clearly.
    count = min(len(prs.slides), len(slides))
    for i in range(count):
        slide = prs.slides[i]
        item = slides[i]
        body = item.get("body", "")
        mapping = {
            "{{TITLE}}": str(item.get("title", "")),
            "{{BODY}}": str(body),
            "{{NOTES}}": str(item.get("speaker_note", item.get("notes", ""))),
            "{{VISUAL}}": str(item.get("visual_suggestion", "")),
            "{{REFERENCES}}": "\n".join(outline.get("references", [])),
        }
        for shape in slide.shapes:
            replace_tokens_in_shape(shape, mapping)
        # If no placeholders were used, replace text in the first title-like shape and first body-like shape.
        all_text = "\n".join(getattr(s, "text", "") for s in slide.shapes if getattr(s, "has_text_frame", False))
        if "{{TITLE}}" not in all_text and "{{BODY}}" not in all_text:
            text_shapes = [s for s in slide.shapes if getattr(s, "has_text_frame", False)]
            if text_shapes:
                # Replace the first non-empty text shape with slide title; second with body when available.
                nonempty = [s for s in text_shapes if s.text.strip()]
                if nonempty:
                    nonempty[0].text = str(item.get("title", ""))
                if len(nonempty) > 1:
                    nonempty[1].text = str(body)
    return count

@app.post("/api/generate")
async def generate(
    topic: str = Form(...),
    audience: str = Form("Mahasiswa dan dosen"),
    slide_count: int = Form(8),
    style: str = Form("Akademik modern"),
    extra_instructions: str = Form(""),
    source_file: Optional[UploadFile] = File(None),
    template_file: Optional[UploadFile] = File(None),
):
    if slide_count < 3 or slide_count > 20:
        raise HTTPException(status_code=400, detail="Jumlah slide harus 3–20.")
    source_text = ""
    if source_file and source_file.filename:
        name = source_file.filename.lower()
        raw = await source_file.read()
        if name.endswith(".txt") or name.endswith(".md"):
            source_text = raw.decode("utf-8", errors="replace")
        elif name.endswith(".pptx"):
            try:
                src_prs = Presentation(io.BytesIO(raw))
                source_text = "\n".join(
                    f"Slide {s['slide']}: " + " | ".join(s["texts"])
                    for s in extract_text_from_pptx(src_prs)
                )
            except Exception:
                raise HTTPException(status_code=400, detail="File sumber PPTX tidak dapat dibaca.")
        elif name.endswith(".docx"):
            try:
                from docx import Document
                doc = Document(io.BytesIO(raw))
                source_text = "\n".join(p.text for p in doc.paragraphs)
            except Exception:
                raise HTTPException(status_code=400, detail="DOCX tidak dapat dibaca. Coba TXT atau PPTX.")
        else:
            raise HTTPException(status_code=400, detail="Versi prototipe menerima sumber TXT, MD, DOCX, atau PPTX. PDF belum didukung.")
    outline, mode = make_ai_outline(topic, audience, slide_count, source_text, style, extra_instructions)
    if template_file and template_file.filename:
        if not template_file.filename.lower().endswith(".pptx"):
            raise HTTPException(status_code=400, detail="Template harus berformat PPTX.")
        template_bytes = await template_file.read()
        try:
            prs = Presentation(io.BytesIO(template_bytes))
            used = apply_content_to_template(prs, outline)
            if len(prs.slides) != slide_count:
                warning = f"Template berisi {len(prs.slides)} slide; konten diterapkan pada {used} slide pertama. Samakan jumlah slide template dengan jumlah slide yang dipilih untuk hasil terbaik."
            else:
                warning = "Konten diterapkan ke slide template. Periksa kembali layout dan teks sebelum presentasi."
        except Exception as e:
            raise HTTPException(status_code=400, detail=f"Template tidak dapat diproses: {str(e)[:200]}")
    else:
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        warning = "Tidak ada template diunggah; file dibuat dengan layout dasar. Untuk hasil Canva, unggah PPTX template yang sudah memiliki placeholder."
        for item in outline.get("slides", []):
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            title_box = slide.shapes.add_textbox(Inches(0.7), Inches(0.55), Inches(11.9), Inches(1.0))
            p = title_box.text_frame.paragraphs[0]
            p.text = item.get("title", "")
            p.font.size = Pt(28)
            p.font.bold = True
            p.font.color.rgb = RGBColor(35, 55, 70)
            body_box = slide.shapes.add_textbox(Inches(0.9), Inches(1.8), Inches(11.5), Inches(4.8))
            tf = body_box.text_frame
            tf.word_wrap = True
            tf.text = item.get("body", "")
            for p in tf.paragraphs:
                p.font.size = Pt(20)
                p.font.color.rgb = RGBColor(55, 65, 75)
            if item.get("speaker_note"):
                try:
                    slide.notes_slide.notes_text_frame.text = item["speaker_note"]
                except Exception:
                    pass
    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    headers = {
        "Content-Disposition": 'attachment; filename="slidecraft-presentasi.pptx"',
        "X-Generation-Mode": mode,
        "X-Template-Warning": warning.replace("\n", " ")[:500],
    }
    return StreamingResponse(out, media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation", headers=headers)

