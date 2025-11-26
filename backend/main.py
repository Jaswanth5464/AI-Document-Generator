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
        print(f"\n{'='*50}")
        print(f"Export Request Received")
        print(f"{'='*50}")
        print(f"Topic: {request.topic}")
        print(f"Doc Type: {request.docType}")
        print(f"Number of sections: {len(request.sections)}")
        
        # Convert Pydantic models to dictionaries
        sections_data = []
        for section in request.sections:
            section_dict = {
                'id': section.id,
                'title': section.title,
                'content': section.content
            }
            sections_data.append(section_dict)
            print(f"  - Section {section.id}: {section.title} ({len(section.content)} chars)")
        
        print(f"\nGenerating {request.docType} file...")
        
        if request.docType == "docx":
            # Generate Word document
            file_bytes = generate_docx(request.topic, sections_data)
            filename = f"{request.topic.replace(' ', '_')}.docx"
            media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            print(f"✅ Word document generated: {len(file_bytes)} bytes")
        
        else:  # pptx
            # Generate PowerPoint
            file_bytes = generate_pptx(request.topic, sections_data)
            filename = f"{request.topic.replace(' ', '_')}.pptx"
            media_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            print(f"✅ PowerPoint generated: {len(file_bytes)} bytes")
        
        print(f"📥 Sending file: {filename}")
        print(f"{'='*50}\n")
        
        # Return file
        return StreamingResponse(
            io.BytesIO(file_bytes),
            media_type=media_type,
            headers={
                "Content-Disposition": f"attachment; filename={filename}",
                "Access-Control-Expose-Headers": "Content-Disposition"
            }
        )
    
    except Exception as e:
        print(f"\n❌ ERROR EXPORTING DOCUMENT:")
        print(f"Error: {e}")
        print(f"Traceback:")
        traceback.print_exc()
        print(f"{'='*50}\n")
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/generate-template")
async def generate_template(topic: str, doc_type: str, num_sections: int = 5):
    """Generate suggested outline/template"""
    try:
        print(f"Generating template for: {topic}")
        
        if doc_type == "docx":
            prompt = f"""
Create an outline for a professional document about: {topic}

Provide {num_sections} section titles that would make a comprehensive document.
Return ONLY the section titles, one per line, numbered.
Example format:
1. Introduction
2. Background
3. Analysis
"""
        else:
            prompt = f"""
Create a PowerPoint presentation outline about: {topic}

Provide {num_sections} slide titles for an effective presentation.
Return ONLY the slide titles, one per line, numbered.
Example format:
1. Title Slide
2. Overview
3. Key Points
"""
        
        response = model.generate_content(prompt)
        
        # Parse response into list
        lines = response.text.strip().split('\n')
        sections = []
        for i, line in enumerate(lines[:num_sections], 1):
            # Remove numbering if present
            title = line.split('.', 1)[-1].strip() if '.' in line else line.strip()
            if title:
                sections.append({
                    "id": i,
                    "title": title,
                    "content": ""
                })
        
        print(f"Generated {len(sections)} template sections")
        return {"sections": sections}
    
    except Exception as e:
        print(f"Error generating template: {e}")
        traceback.print_exc()
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
