import streamlit as st
import openai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from io import BytesIO
import requests

# Function to generate AI-powered content for PPT
def generate_detailed_ppt_content(domain, topic):
    """Generate detailed content for a PowerPoint presentation using GPT-4."""
    prompt = (
        f"You are an expert in the {domain} domain. Generate a professional, formal PowerPoint presentation on '{topic}'. "
        f"The structure should include: "
        f"- Title Slide (topic, subtitle, author name placeholder) "
        f"- Introduction Slide (definition, importance) "
        f"- 4-6 Key Point Slides (elaborated details) "
        f"- Case Studies/Examples Slide (real-world applications) "
        f"- Conclusion Slide (summary, future insights). "
        f"Each slide must have at least 4 bullet points, **NO long paragraphs**. "
        f"Suggest **relevant SmartArt or images** for visualization."
    )

    try:
        response = openai.chat.completions.create(
            model="gpt-4",
            messages=[
                {"role": "system", "content": f"You are a {domain} domain expert."},
                {"role": "user", "content": prompt},
            ],
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"Error: {str(e)}"


# Function to create a well-structured PowerPoint presentation
def create_professional_ppt(content, topic, file_name="presentation.pptx"):
    """Create a PowerPoint presentation with good formatting and SmartArt."""
    ppt = Presentation()

    # Title Slide
    title_slide_layout = ppt.slide_layouts[0]
    slide = ppt.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    title.text = topic
    subtitle.text = "A Comprehensive Overview"

    # Process AI-generated content into slides
    sections = content.split("\n\n")
    for section in sections:
        if ":" in section:  # Detect title: content structure
            slide_title, slide_content = section.split(":", 1)

            # Select slide layout
            layout = 5 if "image" in slide_content.lower() else 1
            slide = ppt.slides.add_slide(ppt.slide_layouts[layout])
            title = slide.shapes.title
            title.text = slide_title.strip()

            # Add slide content
            content_box = slide.placeholders[1]
            content_box.text = slide_content.strip()

            
    # Save the presentation
    ppt.save(file_name)
    return file_name

# Streamlit UI
st.header("📊 AI-Powered PPT Generator")

# User inputs domain and topic
domain = st.text_input("Enter the domain of your presentation (e.g., Medical, Business, Technology):")
topic = st.text_input("Enter the topic for your presentation:")

if st.button("Generate PPT"):
    if domain and topic:
        st.info(f"Generating a detailed PowerPoint presentation for **{topic}** in **{domain}** domain. Please wait...")

        # Generate AI content
        detailed_content = generate_detailed_ppt_content(domain, topic)

        if "Error" not in detailed_content:
            # Create PowerPoint
            ppt_file_name = create_professional_ppt(detailed_content, topic)

            st.success("Your PowerPoint presentation has been successfully generated!")
            with open(ppt_file_name, "rb") as file:
                st.download_button(
                    "📥 Download Your PPT",
                    file,
                    file_name=ppt_file_name,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )
        else:
            st.error(detailed_content)

# Footer
st.markdown("---")
st.caption("Developed by **Corbin Technology Solutions**")
