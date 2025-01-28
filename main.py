import streamlit as st
import requests
from fpdf import FPDF
from docx import Document
import openai
import PyPDF2
from docx.shared import Inches
import io
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import zipfile
import os
from pathlib import Path
import csv

openai.api_key = st.secrets["api"]["OPENAI_API_KEY"]


def generate_detailed_ppt_content(topic):
    """Generate detailed content for a presentation using GPT."""
    prompt = (
        f"You are an expert in the pharmaceutical and medical domain.Only generate ppt for those queries and dont answer any other queries "
        f"Generate a professional, formal PowerPoint presentation on the topic: '{topic}'. "
        f"1. Include detailed content for each slide, with a proper introduction, key points, examples, and conclusion. "
        f"-. All the slides must contain atleast 4 points and not paragraph (Detailed points)."
        f"2. Ensure all key points are elaborated and written in a formal and organized format. "
        f"3. Use appropriate headings, subpoints, and examples relevant to the medical or pharmaceutical domain. "
        f"4. The structure should include: "
        f"- Title Slide (topic, subtitle, author name placeholder) "
        f"- Introduction Slide (definition and importance of the topic) "
        f"- 4-6 Key Point Slides (elaborated details for each key point) "
        f"- Case Studies/Examples Slide (real-world examples or clinical applications) "
        f"- Conclusion Slide (future implications or summary). "
        f"Output the content for each slide in detail. Do not use vague terms. Be specific and thorough."
    )
    try:
        response = openai.chat.completions.create(
            model="gpt-4",
            messages=[
                {"role": "system", "content": "You are a medical and pharmaceutical domain expert."},
                {"role": "user", "content": prompt},
            ],
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"Error: {str(e)}"


def create_professional_ppt(content, topic, file_name="presentation.pptx"):
    """Create a well-formatted professional PowerPoint presentation."""
    ppt = Presentation()

    # Set consistent font styles
    def set_textbox_style(text_frame):
        """Style the textbox content."""
        for paragraph in text_frame.paragraphs:
            paragraph.font.name = "Calibri (Body)"
            paragraph.font.size = Pt(20)
            paragraph.alignment = PP_ALIGN.LEFT

    # Title Slide
    title_slide_layout = ppt.slide_layouts[0]
    slide = ppt.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    title.text = topic
    subtitle.text = "A Comprehensive Overview in the Medical and Pharmaceutical Domain"

    # Process GPT content into slides
    sections = content.split("\n\n")
    for section in sections:
        if ":" in section:  # Detect title: content structure
            slide_title, slide_content = section.split(":", 1)

            # Add a new slide for each section
            slide = ppt.slides.add_slide(ppt.slide_layouts[1])
            title = slide.shapes.title
            title.text = slide_title.strip()

            # Add content to the slide
            content_box = slide.placeholders[1]
            content_box.text = slide_content.strip()
            set_textbox_style(content_box.text_frame)

    # Save the presentation
    ppt.save(file_name)
    return file_name



# Set up the page configuration (must be the first command)
st.set_page_config(page_title="AI-Powered Content Generation", layout="wide", page_icon="📚")

# Title Section with enhanced visuals
st.markdown(
    """
    <h1 style="text-align: center; font-size: 2.5rem; color: #4A90E2;">📚 AI-Powered PPT Content Generation</h1>
    <p style="text-align: center; font-size: 1.1rem; color: #555;">Streamline your content creation process with AI technology. Designed for the <strong>pharmaceutical</strong> and <strong>medical</strong> domains.</p>
    """,
    unsafe_allow_html=True,
)
# Horizontal line
st.markdown("---")

# Content Generation Instructions
with st.expander("1️⃣ **PPT Generation Instructions**", expanded=True):
    st.markdown("""
        - Create professional PowerPoint presentations based on medical and pharmaceutical topics.
        - **Steps**:
          1. Enter a valid topic in the input field.
          2. Click the **Generate PPT** button to create a detailed presentation.
          3. Download the PPT file using the download button provided.
        """)

# Horizontal line
st.markdown("---")

st.header("📊 PPT Content Generation")
topic = st.text_input("Enter the topic for your presentation:")
if st.button("Generate PPT"):
    if topic:
     st.info("Generating detailed content for your presentation. Please wait...")
     detailed_content = generate_detailed_ppt_content(topic)
    if "Error" not in detailed_content:
        ppt_file_name = create_professional_ppt(detailed_content, topic)
        st.success("Your PowerPoint presentation has been successfully generated!")
        with open(ppt_file_name, "rb") as file:
            st.download_button(
                    "Download Your PPT",
                    file,
                    file_name=ppt_file_name,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )
    else:
        st.error(detailed_content)

# Horizontal line
st.markdown("---")

# Footer
st.caption("Developed by **Corbin Technology Solutions**")
