from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from typing import List, Optional
import google.generativeai as genai
import os
from dotenv import load_dotenv
import io
import traceback

from utils.docx_generator import generate_docx
from utils.pptx_generator import generate_pptx

load_dotenv()

# Initialize Gemini
genai.configure(api_key=os.getenv('GEMINI_API_KEY'))
model = genai.GenerativeModel('gemini-2.5-flash-lite')

app = FastAPI()

# Enable CORS
# Enable CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:5173",  # Local development
        "https://*.onrender.com",  # Render frontend
        "*"  # Allow all (for testing)
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ============ MODELS ============
class Section(BaseModel):
    id: int
    title: str
    content: str

class GenerateRequest(BaseModel):
    topic: str
    sections: List[Section]
    docType: str

class GenerateSectionRequest(BaseModel):
    topic: str
    sectionTitle: str
    docType: str

class RefineRequest(BaseModel):
    currentContent: str
    instruction: str

class ExportRequest(BaseModel):
    topic: str
    sections: List[Section]
    docType: str
    theme: Optional[str] = "professional_blue"

# ============ ROUTES ============

@app.api_route("/", methods=["GET", "HEAD"])
def read_root():
    return {"message": "🚀 AI Document Generator API is running!"}

@app.post("/api/generate-section")
async def generate_section(request: GenerateSectionRequest):
    """Generate content for a single section"""
    try:
        print(f"Generating content for: {request.sectionTitle}")
        
        if request.docType == "docx":
            prompt = f"""
You are writing a section for a professional document about: {request.topic}

Section Title: {request.sectionTitle}

Write detailed, well-structured content for this section (3-4 paragraphs).
Make it professional, informative, and engaging.
Use clear language and proper formatting.
Do not include the section title in your response.
"""
        else:  # pptx
            prompt = f"""
You are creating content for a PowerPoint slide about: {request.topic}

Slide Title: {request.sectionTitle}

Write concise, impactful content for this slide (4-6 bullet points).
Keep it brief and presentation-friendly.
Each point should be clear and actionable.
Do not include the slide title in your response.
Format as bullet points using • symbol.
"""
        
        response = model.generate_content(prompt)
        content = response.text.strip()
        
        print(f"Content generated successfully for: {request.sectionTitle}")
        return {"content": content}
    
    except Exception as e:
        print(f"Error generating section: {e}")
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/refine-section")
async def refine_section(request: RefineRequest):
    """Refine existing content based on user instruction"""
    try:
        print(f"Refining content with instruction: {request.instruction}")
        
        prompt = f"""
Current Content:
{request.currentContent}

User Instruction: {request.instruction}

Rewrite the content following the user's instruction.
Maintain professional quality and coherence.
Keep the same general structure unless asked to change it.
Do not add any preamble or explanation, just provide the refined content.
"""
        
        response = model.generate_content(prompt)
        refined_content = response.text.strip()
        
        print("Content refined successfully")
        return {"refinedContent": refined_content}
    
    except Exception as e:
        print(f"Error refining content: {e}")
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/export-document")
async def export_document(request: ExportRequest):
    """Export document as .docx or .pptx"""
    try:
        print("\n===== EXPORT REQUEST RECEIVED =====")
        print("Topic:", request.topic)
        print("DocType:", request.docType)
        print("Theme:", request.theme)
        print("Sections received:", len(request.sections))

        # 🔥 SAFE conversion: works for Pydantic objects & dicts
        sections_data = []
        for s in request.sections:
            if isinstance(s, dict):  # frontend
                sections_data.append({
                    "id": s.get("id"),
                    "title": s.get("title"),
                    "content": s.get("content")
                })
            else:  # Pydantic Section
                sections_data.append({
                    "id": s.id,
                    "title": s.title,
                    "content": s.content
                })

        print("Converted sections (final):", sections_data)

        # Generate file
        if request.docType == "docx":
            file_bytes = generate_docx(request.topic, sections_data)
            filename = f"{request.topic.replace(' ', '_')}.docx"
            media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

        else:  # pptx
            theme = request.theme or "professional_blue"
            file_bytes = generate_pptx(request.topic, sections_data, theme)
            filename = f"{request.topic.replace(' ', '_')}.pptx"
            media_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

        # Send file
        return StreamingResponse(
            io.BytesIO(file_bytes),
            media_type=media_type,
            headers={
                "Content-Disposition": f"attachment; filename={filename}"
            }
        )

    except Exception as e:
        print("❌ EXPORT ERROR:", e)
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))



@app.post("/api/generate-template")
async def generate_template(topic: str, doc_type: str, num_sections: int = 5):
    """Generate suggested outline/template"""
    try:
        if doc_type == "docx":
            prompt = f"""
Generate EXACTLY {num_sections} section titles for a professional document about: {topic}

IMPORTANT RULES:
- Return ONLY the section titles
- One title per line
- NO explanations, NO introductions, NO extra text
- Start each line with just the title
- Do NOT number them

Example format:
Introduction
Market Analysis
Key Findings
Recommendations
Conclusion
"""
        else:
            prompt = f"""
Generate EXACTLY {num_sections} slide titles for a PowerPoint presentation about: {topic}

IMPORTANT RULES:
- Return ONLY the slide titles
- One title per line
- NO explanations, NO introductions, NO extra text
- Start each line with just the title
- Do NOT number them
- Keep titles SHORT (3-7 words max)

Example format:
Introduction to AI Trading
How AI Analyzes Markets
Key Trading Algorithms
Benefits and Risks
Future of AI Trading
"""
        
        response = model.generate_content(prompt)
        text = response.text.strip()
        
        # Clean up the response - remove any explanatory text
        lines = text.split('\n')
        clean_lines = []
        
        for line in lines:
            line = line.strip()
            
            # Skip empty lines
            if not line:
                continue
            
            # Skip lines that look like explanations
            if any(skip_word in line.lower() for skip_word in [
                'here are', 'these are', 'following', 'presentation', 
                'document', 'titles', 'for an', 'about'
            ]):
                continue
            
            # Remove numbering (1., 2., etc.)
            if line[0].isdigit() and ('.' in line or ')' in line):
                # Remove "1. " or "1) " from start
                line = line.split('.', 1)[-1].split(')', 1)[-1].strip()
            
            # Remove asterisks or dashes from markdown
            line = line.lstrip('*-•').strip()
            
            # Only add if it's not too long (likely a description)
            if len(line.split()) <= 15:  # Max 15 words for a title
                clean_lines.append(line)
        
        # Take only the requested number of sections
        clean_lines = clean_lines[:num_sections]
        
        # If we don't have enough, add generic ones
        while len(clean_lines) < num_sections:
            if doc_type == "docx":
                clean_lines.append(f"Section {len(clean_lines) + 1}")
            else:
                clean_lines.append(f"Slide {len(clean_lines) + 1}")
        
        # Create sections array
        sections = []
        for i, title in enumerate(clean_lines, 1):
            sections.append({
                "id": i,
                "title": title,
                "content": ""
            })
        
        return {"sections": sections}
    
    except Exception as e:
        print(f"Error in generate_template: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))
if __name__ == "__main__":
    import uvicorn
    print("\n" + "="*60)
    print("🚀 AI DOCUMENT GENERATOR API")
    print("="*60)
    print(f"📍 Server: http://localhost:8000")
    print(f"📖 Docs: http://localhost:8000/docs")
    print(f"🤖 AI Model: {model.model_name}")
    print("="*60 + "\n")
    uvicorn.run(app, host="localhost", port=8000)
