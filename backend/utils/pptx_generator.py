from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE

import io
import re

# ========== THEME DEFINITIONS ==========
THEMES = {
    'professional_blue': {
        'name': 'Professional Blue',
        'bg_start': (227, 242, 253),  # Light blue
        'bg_end': (187, 222, 251),    # Slightly darker blue
        'title_color': (25, 118, 210),  # Deep blue
        'text_color': (33, 33, 33),     # Dark gray
        'accent_color': (66, 165, 245)  # Bright blue
    },
    'modern_dark': {
        'name': 'Modern Dark',
        'bg_start': (48, 48, 48),      # Dark gray
        'bg_end': (33, 33, 33),        # Darker gray
        'title_color': (255, 255, 255), # White
        'text_color': (220, 220, 220),  # Light gray
        'accent_color': (102, 126, 234) # Purple
    },
    'vibrant_orange': {
        'name': 'Vibrant Orange',
        'bg_start': (255, 243, 224),   # Light orange
        'bg_end': (255, 224, 178),     # Darker orange
        'title_color': (230, 81, 0),   # Deep orange
        'text_color': (62, 39, 35),    # Dark brown
        'accent_color': (255, 152, 0)  # Bright orange
    },
    'nature_green': {
        'name': 'Nature Green',
        'bg_start': (232, 245, 233),   # Light green
        'bg_end': (200, 230, 201),     # Darker green
        'title_color': (27, 94, 32),   # Forest green
        'text_color': (33, 33, 33),    # Dark gray
        'accent_color': (76, 175, 80)  # Bright green
    },
    'elegant_purple': {
        'name': 'Elegant Purple',
        'bg_start': (243, 229, 245),   # Light purple
        'bg_end': (225, 190, 231),     # Darker purple
        'title_color': (74, 20, 140),  # Deep purple
        'text_color': (33, 33, 33),    # Dark gray
        'accent_color': (142, 36, 170) # Bright purple
    }
}

# ========== ICON MAPPING ==========
ICON_MAP = {
    # Introduction & Overview
    'introduction': '📊',
    'overview': '👁️',
    'agenda': '📋',
    'outline': '📝',
    
    # Business & Strategy
    'strategy': '🎯',
    'business': '💼',
    'market': '📈',
    'trading': '💹',
    'finance': '💰',
    'investment': '💵',
    'sales': '🤝',
    'marketing': '📢',
    
    # Analysis & Data
    'analysis': '🔍',
    'data': '📊',
    'statistics': '📉',
    'metrics': '📏',
    'report': '📄',
    'research': '🔬',
    
    # Technology
    'technology': '💻',
    'ai': '🤖',
    'artificial intelligence': '🤖',
    'machine learning': '🧠',
    'algorithm': '⚙️',
    'automation': '🔄',
    'digital': '📱',
    'software': '💾',
    
    # Risk & Security
    'risk': '⚠️',
    'security': '🔒',
    'protection': '🛡️',
    'safety': '🦺',
    
    # Growth & Success
    'growth': '📈',
    'success': '🏆',
    'achievement': '🎖️',
    'goals': '🎯',
    'target': '🎯',
    
    # Future & Innovation
    'future': '🔮',
    'innovation': '💡',
    'trends': '📊',
    'forecast': '🌤️',
    'prediction': '🔮',
    
    # Communication
    'communication': '💬',
    'team': '👥',
    'collaboration': '🤝',
    'meeting': '🗓️',
    
    # Process & Timeline
    'process': '⚙️',
    'timeline': '📅',
    'roadmap': '🗺️',
    'workflow': '🔄',
    
    # Results & Conclusion
    'results': '✅',
    'conclusion': '🏁',
    'summary': '📝',
    'takeaway': '🎁',
    'recommendation': '👍',
    
    # Problems & Solutions
    'problem': '❗',
    'challenge': '🧗',
    'solution': '💡',
    'benefits': '✨',
    'advantages': '➕',
    
    # Education & Learning
    'education': '🎓',
    'learning': '📚',
    'training': '🏋️',
    'knowledge': '🧠',
}

# ========== HELPER FUNCTIONS ==========

def get_icon_for_title(title: str) -> str:
    """Find best matching icon for slide title"""
    title_lower = title.lower()
    
    # Check for direct matches first
    for keyword, icon in ICON_MAP.items():
        if keyword in title_lower:
            return icon
    
    # Default icons based on position
    if any(word in title_lower for word in ['intro', 'start', 'begin', 'welcome']):
        return '📊'
    elif any(word in title_lower for word in ['end', 'conclude', 'final', 'summary']):
        return '🏁'
    elif any(word in title_lower for word in ['thank', 'questions', 'q&a']):
        return '🙏'
    
    return ''  # No icon

def clean_text_formatting(text: str) -> str:
    """Remove ** markdown and clean text"""
    # Remove ** for bold
    text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
    
    # Remove * for italic
    text = re.sub(r'\*(.+?)\*', r'\1', text)
    
    # Remove leading bullets if present
    text = re.sub(r'^[\*\-•]\s*', '', text, flags=re.MULTILINE)
    
    return text.strip()

def apply_gradient_background(slide, theme_colors):
    """Apply gradient background to slide"""
    try:
        background = slide.background
        fill = background.fill
        fill.gradient()
        fill.gradient_angle = 45.0  # Diagonal gradient
        
        # Set gradient stops
        fill.gradient_stops[0].color.rgb = RGBColor(*theme_colors['bg_start'])
        fill.gradient_stops[1].color.rgb = RGBColor(*theme_colors['bg_end'])
    except Exception as e:
        print(f"Warning: Could not apply gradient: {e}")
        # Fallback to solid color
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(*theme_colors['bg_start'])

def add_decorative_bar(slide, theme_colors):
    """Add decorative accent bar at bottom of slide"""
    try:
        left = Inches(0)
        top = slide.shapes[0].top + slide.shapes[0].height - Inches(0.15)
        width = Inches(10)
        height = Inches(0.15)
        
        shape = slide.shapes.add_shape(
            1,  # Rectangle
            left, top, width, height
        )
        
        # Style the bar
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(*theme_colors['accent_color'])
        shape.line.fill.background()
    except:
        pass  # Skip if error

# ========== MAIN GENERATOR FUNCTION ==========

def generate_pptx(topic: str, sections: list, theme: str = 'professional_blue') -> bytes:
    """Generate a clean, modern, professional PowerPoint presentation"""

    # Validate theme
    if theme not in THEMES:
        theme = "professional_blue"
    theme_colors = THEMES[theme]

    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    # ===== TITLE SLIDE =====
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
    apply_gradient_background(slide, theme_colors)

    # Title Card
    try:
        title_card = slide.shapes.add_shape(
            MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
            Inches(0.8), Inches(2),
            Inches(8.4), Inches(2.5)
        )
        title_card.fill.solid()
        title_card.fill.fore_color.rgb = RGBColor(255, 255, 255)
        title_card.fill.transparency = 0.10
        title_card.line.fill.background()

        shadow = title_card.shadow
        shadow.inherit = False
        shadow.blur_radius = Pt(20)
        shadow.distance = Pt(5)
        shadow.transparency = 0.70
    except:
        pass

    # Main Title
    title_box = slide.shapes.add_textbox(Inches(1), Inches(2.4), Inches(8), Inches(1))
    tf = title_box.text_frame
    tf.text = topic
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.font.size = Pt(48)
    p.font.bold = True
    p.font.color.rgb = RGBColor(*theme_colors["title_color"])

    # Subtitle
    subtitle_box = slide.shapes.add_textbox(Inches(1), Inches(3.4), Inches(8), Inches(1))
    sf = subtitle_box.text_frame
    sf.text = "AI-Generated Presentation"
    sp = sf.paragraphs[0]
    sp.alignment = PP_ALIGN.CENTER
    sp.font.size = Pt(24)
    sp.font.color.rgb = RGBColor(*theme_colors["text_color"])

    # ===== CONTENT SLIDES =====
    for i, section in enumerate(sections, start=1):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        apply_gradient_background(slide, theme_colors)

        # ----- HEADER BAR -----
        header = slide.shapes.add_shape(
            MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
            Inches(0.5), Inches(0.4),
            Inches(9), Inches(1)
        )
        header.fill.solid()
        header.fill.fore_color.rgb = RGBColor(255, 255, 255)
        header.fill.transparency = 0.03
        header.line.fill.solid()
        header.line.fill.fore_color.rgb = RGBColor(*theme_colors["accent_color"])
        header.line.width = Pt(1.4)

        # Title text
        icon = get_icon_for_title(section["title"])
        title_text = f"{icon}  {section['title']}" if icon else section["title"]

        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.55), Inches(8.4), Inches(0.8))
        tf = title_box.text_frame
        tf.text = title_text
        tp = tf.paragraphs[0]
        tp.font.size = Pt(32)
        tp.font.bold = True
        tp.font.color.rgb = RGBColor(*theme_colors["title_color"])

        # Accent underline
        underline = slide.shapes.add_shape(
            MSO_AUTO_SHAPE_TYPE.RECTANGLE,
            Inches(0.8), Inches(1.35),
            Inches(8.4), Inches(0.07)
        )
        underline.fill.solid()
        underline.fill.fore_color.rgb = RGBColor(*theme_colors["accent_color"])
        underline.line.fill.background()

        # ----- CONTENT CARD -----
        content_bg = slide.shapes.add_shape(
            MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
            Inches(0.8), Inches(1.65),
            Inches(8.4), Inches(4.8)
        )
        content_bg.fill.solid()
        content_bg.fill.fore_color.rgb = RGBColor(255, 255, 255)
        content_bg.fill.transparency = 0.04
        content_bg.line.fill.background()

        shadow = content_bg.shadow
        shadow.inherit = False
        shadow.blur_radius = Pt(12)
        shadow.distance = Pt(3)
        shadow.transparency = 0.6

        # ----- CONTENT TEXT -----
        text_box = slide.shapes.add_textbox(
            Inches(1.1), Inches(1.85),
            Inches(7.8), Inches(4.4)
        )
        tf = text_box.text_frame
        tf.word_wrap = True

        cleaned = clean_text_formatting(section.get("content", ""))
        lines = [line.strip() for line in cleaned.split("\n") if line.strip()]

        for j, line in enumerate(lines):
            p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
            p.text = line
            p.bullet = True
            p.font.size = Pt(22)
            p.font.color.rgb = RGBColor(*theme_colors["text_color"])
            p.space_before = Pt(6)
            p.space_after = Pt(6)

        # Slide number
        num = slide.shapes.add_textbox(Inches(9), Inches(7), Inches(0.5), Inches(0.3))
        n = num.text_frame.paragraphs[0]
        n.text = str(i)
        n.font.size = Pt(14)
        n.font.color.rgb = RGBColor(*theme_colors["text_color"])
        n.alignment = PP_ALIGN.RIGHT

    # Output PPTX
    stream = io.BytesIO()
    prs.save(stream)
    stream.seek(0)
    return stream.getvalue()
