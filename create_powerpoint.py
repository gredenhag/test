#!/usr/bin/env python3
"""
Golf Team Strategy & Mindset PowerPoint Presentation Generator
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

def create_golf_presentation():
    # Create presentation object
    prs = Presentation()
    
    # Define color scheme
    title_color = RGBColor(0, 100, 0)  # Dark green
    subtitle_color = RGBColor(0, 150, 0)  # Medium green
    text_color = RGBColor(0, 0, 0)  # Black
    accent_color = RGBColor(255, 140, 0)  # Orange
    
    def add_title_slide(title, subtitle=""):
        slide_layout = prs.slide_layouts[0]  # Title slide layout
        slide = prs.slides.add_slide(slide_layout)
        
        title_shape = slide.shapes.title
        subtitle_shape = slide.placeholders[1]
        
        title_shape.text = title
        title_shape.text_frame.paragraphs[0].font.color.rgb = title_color
        title_shape.text_frame.paragraphs[0].font.size = Pt(44)
        title_shape.text_frame.paragraphs[0].font.bold = True
        
        if subtitle:
            subtitle_shape.text = subtitle
            subtitle_shape.text_frame.paragraphs[0].font.color.rgb = subtitle_color
            subtitle_shape.text_frame.paragraphs[0].font.size = Pt(24)
    
    def add_content_slide(title, content_points):
        slide_layout = prs.slide_layouts[1]  # Title and content layout
        slide = prs.slides.add_slide(slide_layout)
        
        # Title
        title_shape = slide.shapes.title
        title_shape.text = title
        title_shape.text_frame.paragraphs[0].font.color.rgb = title_color
        title_shape.text_frame.paragraphs[0].font.size = Pt(36)
        title_shape.text_frame.paragraphs[0].font.bold = True
        
        # Content
        content_shape = slide.placeholders[1]
        text_frame = content_shape.text_frame
        text_frame.clear()
        
        for i, point in enumerate(content_points):
            if i == 0:
                p = text_frame.paragraphs[0]
            else:
                p = text_frame.add_paragraph()
            
            p.text = point
            p.font.size = Pt(20)
            p.font.color.rgb = text_color
            p.level = 0
            
            # Make first line bold if it's a section header
            if point.endswith(':') or point.startswith('**'):
                p.font.bold = True
                p.font.color.rgb = accent_color
    
    def add_two_column_slide(title, left_content, right_content):
        slide_layout = prs.slide_layouts[6]  # Blank layout
        slide = prs.slides.add_slide(slide_layout)
        
        # Title
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1))
        title_frame = title_box.text_frame
        title_frame.text = title
        title_frame.paragraphs[0].font.color.rgb = title_color
        title_frame.paragraphs[0].font.size = Pt(36)
        title_frame.paragraphs[0].font.bold = True
        title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        # Left column
        left_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.8), Inches(4.2), Inches(5))
        left_frame = left_box.text_frame
        left_frame.clear()
        
        for i, point in enumerate(left_content):
            if i == 0:
                p = left_frame.paragraphs[0]
            else:
                p = left_frame.add_paragraph()
            p.text = point
            p.font.size = Pt(18)
            p.font.color.rgb = text_color
            if point.endswith(':') or point.startswith('**'):
                p.font.bold = True
                p.font.color.rgb = accent_color
        
        # Right column
        right_box = slide.shapes.add_textbox(Inches(5.3), Inches(1.8), Inches(4.2), Inches(5))
        right_frame = right_box.text_frame
        right_frame.clear()
        
        for i, point in enumerate(right_content):
            if i == 0:
                p = right_frame.paragraphs[0]
            else:
                p = right_frame.add_paragraph()
            p.text = point
            p.font.size = Pt(18)
            p.font.color.rgb = text_color
            if point.endswith(':') or point.startswith('**'):
                p.font.bold = True
                p.font.color.rgb = accent_color
    
    # Slide 1: Title Slide
    add_title_slide(
        "Golf Strategy & Mindset",
        "Playing the Percentages for Lower Scores\n\nPresented by: [Coach Name]\nDate: [Date]\nTeam: [Team Name]"
    )
    
    # Slide 2: Course Management
    add_content_slide(
        "Course Management - Play the Percentages",
        [
            "**The Numbers Don't Lie**",
            "• TOUR Average Distance to Pin from 125-150 yards: 23'5\"",
            "• Key Strategy: If pin is 12\" from right edge with water right",
            "  - Aim 15 feet to the left of the green",
            "  - Play for the center, not the pin",
            "  - Let your short game save par",
            "",
            "**Why This Works**",
            "• Reduces penalty strokes",
            "• Increases green-in-regulation percentage",
            "• Sets up easier up-and-down opportunities"
        ]
    )
    
    # Slide 3: Understanding Your Game
    add_two_column_slide(
        "Understanding Your Game",
        [
            "**Know Your Spray Pattern**",
            "• Identify your miss tendencies",
            "  - Where do you miss when you miss?",
            "  - What's your typical shot shape?",
            "  - How far offline do you typically go?",
            "",
            "**Develop a Go-To Shot**",
            "• One reliable shot you can execute under pressure",
            "• Practice it until it's automatic",
            "• Use it when the pressure is on",
            "• Build confidence through repetition"
        ],
        [
            "**Practice Strategy**",
            "• Focus on your most common shots",
            "• Know your yardages with each club",
            "• Understand your miss patterns",
            "• Build confidence in your go-to shot",
            "",
            "**Mental Approach**",
            "• Trust your swing under pressure",
            "• Don't try to be perfect",
            "• Play to your strengths",
            "• Stay committed to your plan"
        ]
    )
    
    # Slide 4: Practice Priorities
    add_content_slide(
        "Practice Priorities",
        [
            "**Focus Areas (In Order)**",
            "1. **Short Game** - 40% of practice time",
            "2. **Wedges** - 30% of practice time",
            "3. **Speed Control** - 20% of practice time",
            "4. **Full Swing** - 10% of practice time",
            "",
            "**Why This Distribution?**",
            "• Short game saves more strokes",
            "• Wedges are scoring clubs",
            "• Speed control eliminates three-putts",
            "• Full swing is already developed",
            "",
            "**Practice Structure**",
            "• Start with short game every session",
            "• Focus on 30-100 yard shots",
            "• Practice putting from 3-10 feet",
            "• Work on bunker shots and chips"
        ]
    )
    
    # Slide 5: Mental Game
    add_content_slide(
        "Mental Game - Stay Present",
        [
            "**The Present Shot Principle**",
            "• Focus on the shot you're playing",
            "• Not the last shot (good or bad)",
            "• Not the next shot",
            "• **Only the shot in front of you**",
            "",
            "**How to Stay Present**",
            "• Pre-shot routine",
            "• Deep breathing",
            "• Clear your mind",
            "• Visualize the shot",
            "• Execute with commitment",
            "",
            "**Common Mental Mistakes**",
            "• Dwelling on previous shots",
            "• Worrying about future holes",
            "• Getting ahead of yourself",
            "• Losing focus on the process"
        ]
    )
    
    # Slide 6: Scoring Mindset
    add_two_column_slide(
        "Scoring Mindset",
        [
            "**Change Your Language**",
            "",
            "**Instead of saying:**",
            "• \"I made par\"",
            "• \"I got a birdie\"",
            "• \"I made bogey\"",
            "",
            "**Say:**",
            "• \"I made 4\"",
            "• \"I made 3\"",
            "• \"I made 5\""
        ],
        [
            "**Why This Matters**",
            "• Removes emotional attachment to \"par\"",
            "• Focuses on total score",
            "• Reduces pressure on individual holes",
            "• Creates consistency in approach",
            "",
            "**Mental Benefits**",
            "• Less pressure on each shot",
            "• More objective scoring",
            "• Better course management",
            "• Improved focus on process"
        ]
    )
    
    # Slide 7: Tiger's 5 Rules
    add_content_slide(
        "Tiger's 5 Rules",
        [
            "**The Foundation of Great Scoring**",
            "",
            "1. **No Double Bogeys**",
            "2. **No Three-Putts**",
            "3. **No Bogeys from Inside 150 Yards** (scoring clubs)",
            "4. **No Blown Easy Up-and-Downs** (double chips)",
            "5. **No Bogeys on Par Fives**",
            "",
            "**The Impact**",
            "• Follow these rules = 75 or better",
            "• Violate these rules = 80+",
            "• Focus on what you can control",
            "",
            "**How to Track**",
            "• Mark violations on your scorecard",
            "• Review after each round",
            "• Focus on eliminating one rule violation at a time",
            "• Celebrate rounds with zero violations"
        ]
    )
    
    # Slide 8: Aggressive vs. Risky Play
    add_two_column_slide(
        "Aggressive vs. Risky Play",
        [
            "**Aggressive Play:**",
            "• Committed swing to a specific target",
            "• Pre-determined target selection",
            "• Calculated risk",
            "• High percentage of success",
            "",
            "**Risky Play:**",
            "• Swinging at the pin regardless of trouble",
            "• No margin for error",
            "• Low percentage of success",
            "• High penalty potential"
        ],
        [
            "**The Target Selection Process**",
            "1. Identify trouble areas",
            "2. Find the safe side",
            "3. Pick a specific target",
            "4. Commit to the shot",
            "",
            "**Examples**",
            "• Pin tucked left with water left",
            "  → Aim 15 feet right of pin",
            "• Pin back with bunker short",
            "  → Aim for middle of green",
            "• Pin front with water long",
            "  → Take one less club"
        ]
    )
    
    # Slide 9: The Mindset Advantage
    add_content_slide(
        "The Mindset Advantage",
        [
            "**Why Most Players Shoot High Scores**",
            "• Not playing the percentages",
            "• Going for pins in trouble spots",
            "• Emotional decision making",
            "• Lack of course management",
            "",
            "**The Winning Mindset**",
            "• **Skill + Practice + Right Mindset = Low Scores**",
            "• Play the percentages consistently",
            "• Make smart decisions under pressure",
            "• Trust your process",
            "",
            "**Key Mental Traits**",
            "• Patience on the course",
            "• Discipline in shot selection",
            "• Confidence in your abilities",
            "• Focus on the process, not results"
        ]
    )
    
    # Slide 10: Key Takeaways
    add_content_slide(
        "Key Takeaways",
        [
            "**Remember These Principles**",
            "",
            "1. **Play the percentages** - Use the 15-foot rule",
            "2. **Know your game** - Understand your patterns",
            "3. **Practice smart** - Focus on short game",
            "4. **Stay present** - One shot at a time",
            "5. **Score with numbers** - Not par/birdie/bogey",
            "6. **Follow Tiger's rules** - Avoid the big numbers",
            "7. **Be aggressive, not risky** - Committed to targets",
            "",
            "**The Bottom Line**",
            "**It takes skill, practice, and the right mindset to shoot low scores.**",
            "**Most players shoot high because they don't play the percentages.**"
        ]
    )
    
    # Slide 11: Questions & Discussion
    add_content_slide(
        "Questions & Discussion",
        [
            "**Let's Talk About It**",
            "",
            "• Which principle resonates most with you?",
            "• What's your biggest challenge in course management?",
            "• How can we implement these strategies in practice?",
            "• What questions do you have?",
            "",
            "**Next Steps**",
            "• Practice with these principles in mind",
            "• Track your adherence to Tiger's rules",
            "• Focus on one principle per round",
            "• Review and adjust as needed",
            "",
            "**Quote to Remember**",
            "\"Golf is a game of confidence. The more you trust your process, the better you'll play.\""
        ]
    )
    
    return prs

if __name__ == "__main__":
    # Create the presentation
    presentation = create_golf_presentation()
    
    # Save the presentation
    filename = "Golf_Team_Strategy_Mindset_Presentation.pptx"
    presentation.save(filename)
    print(f"PowerPoint presentation saved as: {filename}")
    print("You can now download this file from your workspace.")