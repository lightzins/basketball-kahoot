import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image, ImageDraw, ImageFont

project_dir = r"C:\Users\Enzo\.gemini\antigravity\scratch\basketball-project"
os.makedirs(project_dir, exist_ok=True)
img_dir = os.path.join(project_dir, "images")
os.makedirs(img_dir, exist_ok=True)

# Generate image 1: Basketball ball
def create_basketball_img(path):
    img = Image.new("RGBA", (500, 500), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    # Circle base
    draw.ellipse([20, 20, 480, 480], fill=(234, 88, 12), outline=(15, 23, 42), width=10)
    # Lines
    draw.line([(250, 20), (250, 480)], fill=(15, 23, 42), width=8)
    draw.line([(20, 250), (480, 250)], fill=(15, 23, 42), width=8)
    # Curved arcs
    draw.arc([40, 20, 460, 480], start=45, end=135, fill=(15, 23, 42), width=8)
    draw.arc([40, 20, 460, 480], start=225, end=315, fill=(15, 23, 42), width=8)
    draw.arc([-100, 100, 350, 400], start=290, end=70, fill=(15, 23, 42), width=8)
    draw.arc([150, 100, 600, 400], start=110, end=250, fill=(15, 23, 42), width=8)
    img.save(path, "PNG")

# Generate image 2: Court Diagram
def create_court_img(path):
    img = Image.new("RGBA", (800, 500), (194, 65, 12, 255))
    draw = ImageDraw.Draw(img)
    # Court border
    draw.rectangle([30, 30, 770, 470], outline=(255, 255, 255), width=6)
    # Center line & circle
    draw.line([(400, 30), (400, 470)], fill=(255, 255, 255), width=6)
    draw.ellipse([340, 190, 460, 310], outline=(255, 255, 255), width=6)
    
    # Left key & 3-point
    draw.rectangle([30, 170, 170, 330], fill=(180, 83, 9), outline=(255, 255, 255), width=5)
    draw.ellipse([120, 200, 220, 300], outline=(255, 255, 255), width=5)
    draw.arc([30, 70, 270, 430], start=270, end=90, fill=(255, 255, 255), width=6)
    draw.line([(30, 70), (70, 70)], fill=(255, 255, 255), width=6)
    draw.line([(30, 430), (70, 430)], fill=(255, 255, 255), width=6)
    # Left hoop
    draw.rectangle([45, 225, 50, 275], fill=(255, 255, 255))
    draw.ellipse([50, 238, 74, 262], outline=(234, 88, 12), width=4)

    # Right key & 3-point
    draw.rectangle([630, 170, 770, 330], fill=(180, 83, 9), outline=(255, 255, 255), width=5)
    draw.ellipse([580, 200, 680, 300], outline=(255, 255, 255), width=5)
    draw.arc([530, 70, 770, 430], start=90, end=270, fill=(255, 255, 255), width=6)
    draw.line([(730, 70), (770, 70)], fill=(255, 255, 255), width=6)
    draw.line([(730, 430), (770, 430)], fill=(255, 255, 255), width=6)
    # Right hoop
    draw.rectangle([750, 225, 755, 275], fill=(255, 255, 255))
    draw.ellipse([726, 238, 750, 262], outline=(234, 88, 12), width=4)
    img.save(path, "PNG")

# Generate image 3: Trophy
def create_trophy_img(path):
    img = Image.new("RGBA", (400, 400), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    # Base
    draw.rectangle([120, 320, 280, 360], fill=(51, 65, 85))
    draw.rectangle([140, 290, 260, 320], fill=(245, 158, 11))
    # Stem
    draw.rectangle([185, 210, 215, 290], fill=(217, 119, 6))
    # Cup
    draw.polygon([(130, 90), (270, 90), (240, 210), (160, 210)], fill=(245, 158, 11))
    # Handles
    draw.arc([90, 100, 150, 180], start=90, end=270, fill=(217, 119, 6), width=10)
    draw.arc([250, 100, 310, 180], start=270, end=90, fill=(217, 119, 6), width=10)
    # Ball inside cup
    draw.ellipse([170, 105, 230, 165], fill=(234, 88, 12), outline=(15, 23, 42), width=3)
    img.save(path, "PNG")

ball_path = os.path.join(img_dir, "basketball.png")
court_path = os.path.join(img_dir, "court.png")
trophy_path = os.path.join(img_dir, "trophy.png")

create_basketball_img(ball_path)
create_court_img(court_path)
create_trophy_img(trophy_path)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

DARK_BG = RGBColor(15, 23, 42)      # Slate 900
CARD_BG = RGBColor(30, 41, 59)      # Slate 800
ORANGE = RGBColor(234, 88, 12)      # Primary Orange
LIGHT_ORANGE = RGBColor(254, 215, 170)
WHITE = RGBColor(255, 255, 255)
MUTED = RGBColor(148, 163, 184)     # Slate 400
GOLD = RGBColor(245, 158, 11)

def set_bg(slide):
    bg_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg_shape.fill.solid()
    bg_shape.fill.fore_color.rgb = DARK_BG
    bg_shape.line.fill.background()
    return bg_shape

def add_header(slide, tag_text, title_text):
    # Tag
    tag_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(10), Inches(0.4))
    tf = tag_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = tag_text.upper()
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = ORANGE

    # Title
    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.85), Inches(11), Inches(0.8))
    tf2 = title_box.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.text = title_text
    p2.font.size = Pt(28)
    p2.font.bold = True
    p2.font.color.rgb = WHITE

# ----------------- SLIDE 1: Title Slide -----------------
s1 = prs.slides.add_slide(blank_layout)
set_bg(s1)

# Decorative banner shape
accent = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.8), Inches(0.15), Inches(3.6))
accent.fill.solid()
accent.fill.fore_color.rgb = ORANGE
accent.line.fill.background()

t_box = s1.shapes.add_textbox(Inches(1.2), Inches(1.8), Inches(8), Inches(3.6))
tf = t_box.text_frame
tf.word_wrap = True

p1 = tf.paragraphs[0]
p1.text = "BASKETBALL"
p1.font.size = Pt(56)
p1.font.bold = True
p1.font.color.rgb = ORANGE

p2 = tf.add_paragraph()
p2.text = "The Global Game: Speed, Skill & Passion"
p2.font.size = Pt(28)
p2.font.bold = True
p2.font.color.rgb = WHITE

p3 = tf.add_paragraph()
p3.text = "\nA Comprehensive Presentation in English\nHistory • Rules & Court • Positions • NBA Legends"
p3.font.size = Pt(16)
p3.font.color.rgb = MUTED

s1.shapes.add_picture(ball_path, Inches(8.5), Inches(1.5), Inches(4.2), Inches(4.2))

# ----------------- SLIDE 2: What is Basketball? -----------------
s2 = prs.slides.add_slide(blank_layout)
set_bg(s2)
add_header(s2, "Overview & Essentials", "What is Basketball?")

# Card 1: Core Concept
c1 = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8))
c1.fill.solid()
c1.fill.fore_color.rgb = CARD_BG
c1.line.color.rgb = ORANGE
tf = c1.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "The Core Game\n"
p.font.size = Pt(20)
p.font.bold = True
p.font.color.rgb = ORANGE

bullets = [
    "Two teams of 5 active players on the court.",
    "Objective: score by putting the ball through the opponent's hoop.",
    "Ball movement: dribbling (bouncing) or passing to teammates.",
    "Defense: blocking shots, stealing passes, grabbing rebounds.",
    "Paced by a 24-second shot clock in the NBA, requiring swift strategy."
]
for b in bullets:
    p = tf.add_paragraph()
    p.text = "• " + b
    p.font.size = Pt(14)
    p.font.color.rgb = WHITE

# Card 2: Quick Metrics
metrics = [
    ("450 Million+", "Players worldwide participating at all levels"),
    ("10 Feet (3.05 m)", "Official regulation rim height from the hardwood floor"),
    ("4 Quarters", "12 minutes each in the NBA (10 mins in FIBA / Olympics)")
]

for i, (val, desc) in enumerate(metrics):
    y = 1.8 + (i * 1.65)
    box = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(y), Inches(5.7), Inches(1.4))
    box.fill.solid()
    box.fill.fore_color.rgb = CARD_BG
    box.line.color.rgb = RGBColor(51, 65, 85)
    tf = box.text_frame
    tf.word_wrap = True
    p1 = tf.paragraphs[0]
    p1.text = val
    p1.font.size = Pt(22)
    p1.font.bold = True
    p1.font.color.rgb = GOLD
    p2 = tf.add_paragraph()
    p2.text = desc
    p2.font.size = Pt(13)
    p2.font.color.rgb = MUTED

# ----------------- SLIDE 3: History & Dr. James Naismith -----------------
s3 = prs.slides.add_slide(blank_layout)
set_bg(s3)
add_header(s3, "Origins & Heritage", "The Birth of Basketball (1891)")

events = [
    ("December 1891", "Dr. James Naismith", "Invented at Springfield College, Massachusetts to keep students active indoors during cold New England winters."),
    ("The First Baskets", "Peach Baskets & Ladders", "Peach baskets were nailed to 10-foot gym balcony railings. Every made basket required a janitor with a ladder!"),
    ("13 Original Rules", "Passing Only", "Early rules forbade running with the ball (no dribbling). The game was purely passing, positioning, and set shots."),
    ("1946 - 1949", "Formation of the NBA", "The Basketball Association of America (BAA) merged with the NBL to form the modern NBA.")
]

for i, (date, title, desc) in enumerate(events):
    x = 0.8 + (i * 2.95)
    box = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(2.0), Inches(2.8), Inches(4.6))
    box.fill.solid()
    box.fill.fore_color.rgb = CARD_BG
    box.line.color.rgb = ORANGE if i == 0 else RGBColor(51, 65, 85)
    tf = box.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    p.text = date
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = ORANGE
    
    p2 = tf.add_paragraph()
    p2.text = title + "\n"
    p2.font.size = Pt(17)
    p2.font.bold = True
    p2.font.color.rgb = WHITE
    
    p3 = tf.add_paragraph()
    p3.text = desc
    p3.font.size = Pt(13)
    p3.font.color.rgb = MUTED

# ----------------- SLIDE 4: Court Anatomy & Scoring -----------------
s4 = prs.slides.add_slide(blank_layout)
set_bg(s4)
add_header(s4, "Court & Rules", "Court Anatomy & How Points are Scored")

s4.shapes.add_picture(court_path, Inches(0.8), Inches(2.0), Inches(6.0), Inches(4.5))

rules_box = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.1), Inches(2.0), Inches(5.4), Inches(4.5))
rules_box.fill.solid()
rules_box.fill.fore_color.rgb = CARD_BG
rules_box.line.color.rgb = RGBColor(51, 65, 85)
tf = rules_box.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "Scoring Breakdown\n"
p.font.size = Pt(20)
p.font.bold = True
p.font.color.rgb = ORANGE

scoring_items = [
    ("3-Point Shot:", "Beyond the 3-point arc (23'9\" / 7.24 m in NBA). Revolutionized by modern shooters."),
    ("2-Point Field Goal:", "Any standard basket made inside the arc (layups, floaters, mid-range, dunks)."),
    ("1-Point Free Throw:", "Unopposed shot from 15 ft awarded after fouls. Highly crucial in clutch minutes."),
    ("The Key / Paint:", "Rectangular area under the hoop; offensive players cannot camp there > 3 seconds.")
]

for label, exp in scoring_items:
    p_label = tf.add_paragraph()
    p_label.text = "• " + label + " "
    p_label.font.size = Pt(14)
    p_label.font.bold = True
    p_label.font.color.rgb = GOLD
    
    p_exp = tf.add_paragraph()
    p_exp.text = "  " + exp + "\n"
    p_exp.font.size = Pt(12)
    p_exp.font.color.rgb = WHITE

# ----------------- SLIDE 5: The 5 Positions -----------------
s5 = prs.slides.add_slide(blank_layout)
set_bg(s5)
add_header(s5, "Team Roles", "The 5 Standard Basketball Positions")

positions = [
    ("1 • PG", "Point Guard", "Floor General", "Playmaker, ball handler, sets up the offense with passes.", "Steph Curry\nMagic Johnson"),
    ("2 • SG", "Shooting Guard", "Perimeter Scorer", "Elite jump shooter and perimeter defensive specialist.", "Michael Jordan\nKobe Bryant"),
    ("3 • SF", "Small Forward", "All-Around Threat", "Versatile hybrid who scores inside/outside and defends.", "LeBron James\nKevin Durant"),
    ("4 • PF", "Power Forward", "Paint Enforcer", "Tough rebounder, rim finisher, and mid-range shooter.", "Tim Duncan\nGiannis Anteto."),
    ("5 • C", "Center", "Rim Protector", "Tallest player on floor, blocks shots and controls rebounds.", "Shaquille O'Neal\nNikola Jokić")
]

for i, (num, pos, role, desc, examples) in enumerate(positions):
    x = 0.8 + (i * 2.38)
    box = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(2.0), Inches(2.25), Inches(4.7))
    box.fill.solid()
    box.fill.fore_color.rgb = CARD_BG
    box.line.color.rgb = ORANGE if i == 0 else RGBColor(51, 65, 85)
    tf = box.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    p.text = num
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = ORANGE
    
    p2 = tf.add_paragraph()
    p2.text = pos
    p2.font.size = Pt(16)
    p2.font.bold = True
    p2.font.color.rgb = WHITE
    
    p3 = tf.add_paragraph()
    p3.text = role + "\n"
    p3.font.size = Pt(11)
    p3.font.italic = True
    p3.font.color.rgb = GOLD
    
    p4 = tf.add_paragraph()
    p4.text = desc + "\n"
    p4.font.size = Pt(12)
    p4.font.color.rgb = MUTED
    
    p5 = tf.add_paragraph()
    p5.text = "Notable Players:\n" + examples
    p5.font.size = Pt(11)
    p5.font.bold = True
    p5.font.color.rgb = WHITE

# ----------------- SLIDE 6: NBA Legends -----------------
s6 = prs.slides.add_slide(blank_layout)
set_bg(s6)
add_header(s6, "Hall of Fame", "Greatest Basketball Legends of All Time")

legends = [
    ("Michael Jordan", "#23 • Chicago Bulls", "6× NBA Champion • 6× Finals MVP • 5× MVP", "Widely celebrated as the GOAT for his fierce competitive drive, scoring mastery, and undefeated Finals record.", "30.1 PPG (Record)"),
    ("LeBron James", "#23 / #6 • Lakers / Heat / Cavs", "4× NBA Champion • 4× Finals MVP • 4× MVP", "NBA all-time top scorer (surpassing Kareem Abdul-Jabbar with 40,000+ points), unprecedented longevity and IQ.", "40,000+ Points"),
    ("Kobe Bryant", "#8 / #24 • Los Angeles Lakers", "5× NBA Champion • 2× Finals MVP • 2008 MVP", "The 'Black Mamba' defined relentless work ethic (Mamba Mentality) and recorded the iconic 81-point game in 2006.", "81 Pts in 1 Game")
]

for i, (name, team, honors, desc, stat) in enumerate(legends):
    y = 1.9 + (i * 1.68)
    box = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(y), Inches(11.7), Inches(1.5))
    box.fill.solid()
    box.fill.fore_color.rgb = CARD_BG
    box.line.color.rgb = RGBColor(51, 65, 85)
    tf = box.text_frame
    tf.word_wrap = True
    
    p1 = tf.paragraphs[0]
    p1.text = name + "  "
    p1.font.size = Pt(19)
    p1.font.bold = True
    p1.font.color.rgb = WHITE
    
    run = p1.add_run()
    run.text = team
    run.font.size = Pt(13)
    run.font.color.rgb = ORANGE
    
    p2 = tf.add_paragraph()
    p2.text = honors
    p2.font.size = Pt(12)
    p2.font.bold = True
    p2.font.color.rgb = GOLD
    
    p3 = tf.add_paragraph()
    p3.text = desc
    p3.font.size = Pt(12)
    p3.font.color.rgb = MUTED

# ----------------- SLIDE 7: Conclusion & Thank You -----------------
s7 = prs.slides.add_slide(blank_layout)
set_bg(s7)

s7.shapes.add_picture(trophy_path, Inches(1.5), Inches(1.8), Inches(3.8), Inches(3.8))

t_box = s7.shapes.add_textbox(Inches(5.8), Inches(1.8), Inches(6.8), Inches(4.5))
tf = t_box.text_frame
tf.word_wrap = True

p1 = tf.paragraphs[0]
p1.text = "THANK YOU!"
p1.font.size = Pt(48)
p1.font.bold = True
p1.font.color.rgb = ORANGE

p2 = tf.add_paragraph()
p2.text = "Key Takeaways from the Presentation:"
p2.font.size = Pt(20)
p2.font.bold = True
p2.font.color.rgb = WHITE

recap = [
    "Basketball has united millions globally since Dr. Naismith's 1891 invention.",
    "Fast-paced tactical game requiring teamwork, shooting precision, and athleticism.",
    "5 distinct positions working in harmony to defend and score.",
    "Driven by legends whose work ethic and passion continue to inspire generations."
]
for r in recap:
    p = tf.add_paragraph()
    p.text = "✓ " + r
    p.font.size = Pt(14)
    p.font.color.rgb = MUTED

p_end = tf.add_paragraph()
p_end.text = "\nAny Questions? Let's Play the Quiz Game next!"
p_end.font.size = Pt(16)
p_end.font.bold = True
p_end.font.color.rgb = GOLD

out_path = os.path.join(project_dir, "Basketball_Presentation.pptx")
prs.save(out_path)
print("Saved presentation to:", out_path)
