from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import io

def generate_pptx(topic: str, sections: list) -> bytes:
    """Generate a beautifully formatted PowerPoint presentation"""
    
    # Create presentation
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    
    # Define colors
    PRIMARY_COLOR = RGBColor(102, 126, 234)  # Purple
    DARK_COLOR = RGBColor(51, 51, 51)        # Dark gray
    
    # ========== SLIDE 1: TITLE SLIDE ==========
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    
    # Title
    title = slide.shapes.title
    title.text = topic
    title.text_frame.paragraphs[0].font.size = Pt(44)
    title.text_frame.paragraphs[0].font.bold = True
    title.text_frame.paragraphs[0].font.color.rgb = PRIMARY_COLOR
    title.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    # Subtitle
    subtitle = slide.placeholders[1]
    subtitle.text = "AI-Generated Presentation"
    subtitle.text_frame.paragraphs[0].font.size = Pt(24)
    subtitle.text_frame.paragraphs[0].font.color.rgb = DARK_COLOR
    subtitle.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    # ========== CONTENT SLIDES ==========
    for section in sections:
        # Create new slide with title and content layout
        content_slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(content_slide_layout)
        
        # Set slide title
        title = slide.shapes.title
        title.text = section['title']
        title.text_frame.paragraphs[0].font.size = Pt(32)
        title.text_frame.paragraphs[0].font.bold = True
        title.text_frame.paragraphs[0].font.color.rgb = PRIMARY_COLOR
        
        # Get section content
        content = section.get('content', '')
        
        if content:
            # Get the content placeholder (body)
            body_shape = slide.placeholders[1]
            text_frame = body_shape.text_frame
            text_frame.clear()
            
            # Split content into lines
            lines = content.split('\n')
            
            # Add each line as a bullet point
            for j, line in enumerate(lines):
                line = line.strip()
                
                # Skip empty lines
                if not line:
                    continue
                
                # Remove bullet markers if already present
                if line.startswith('•'):
                    line = line[1:].strip()
                elif line.startswith('-'):
                    line = line[1:].strip()
                elif line.startswith('*'):
                    line = line[1:].strip()
                
                # Create paragraph
                if j == 0:
                    # Use first paragraph
                    p = text_frame.paragraphs[0]
                else:
                    # Add new paragraph
                    p = text_frame.add_paragraph()
                
                # Set text and formatting
                p.text = line
                p.level = 0
                p.font.size = Pt(18)
                p.font.color.rgb = DARK_COLOR
                p.space_before = Pt(12)
                p.space_after = Pt(12)
    
    # ========== SAVE TO BYTES ==========
    file_stream = io.BytesIO()
    prs.save(file_stream)
    file_stream.seek(0)
    
    return file_stream.getvalue()