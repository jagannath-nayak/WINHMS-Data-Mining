import streamlit as st
import pandas as pd
from google import genai
from dotenv import load_dotenv
import os
from reportlab.platypus import SimpleDocTemplate, Paragraph, Image, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A4
from datetime import datetime

# ------------------ CONFIG ------------------

st.set_page_config(
    page_title="Rhythm Hospitality WINHMS Data Insights Application",
    page_icon="🌳",
    layout="centered"
)

load_dotenv()

API_KEY = os.getenv("GOOGLE_API_KEY")
client = genai.Client(api_key=API_KEY)

LOGO_PATH = "assets/rhythm_tree_logo"   
REPORT_DIR = "reports"
os.makedirs(REPORT_DIR, exist_ok=True)

# ------------------ HEADER ------------------

st.image(LOGO_PATH, width=160)
st.title("Rhythm Hospitality – WINHMS Data Insights Application")
st.caption("Prompt-based business intelligence")

st.divider()

# ------------------ FILE UPLOAD ------------------

uploaded_file = st.file_uploader(
    "Upload Hotel Report (CSV / Excel)",
    type=["csv", "xlsx", "xls"]
)

# ------------------ PROMPT BOX ------------------

if uploaded_file:

    prompt = st.text_area(
        "Enter your prompt",
        placeholder="Write your business question here...",
        height=130
    )

    if st.button("Submit"):

        if not prompt.strip():
            st.warning("Please enter a prompt before submitting.")
            st.stop()

        # Read file
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        with st.spinner("Generating insights..."):

            final_prompt = f"""
You are a hospitality analytics expert for Rhythm Hospitality.

User Question:
{prompt}

Analyze the dataset below and provide clear insights and recommendations.

Dataset:
{df.to_csv(index=False)[:12000]}
"""

            try:
                response = client.models.generate_content(
                    model="gemini-2.0-flash",
                    contents=final_prompt
                )
                insights = response.text

            except Exception as e:
                st.error(f"Gemini API Error: {e}")
                st.stop()

        # ------------------ SHOW RESULT ------------------

        st.subheader("AI Insights")
        st.write(insights)

        st.divider()

        # ------------------ PROPERTY SELECTION ------------------

        property_name = st.selectbox(
            "Select Property",
            [
                "Rhythm Gurugram",
                "Rhythm Lonavala",
                "Rhythm Kumarakom"
            ]
        )

        # ------------------ PDF SAVE ------------------

        if st.button("Save Insights as PDF"):

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            pdf_path = f"{REPORT_DIR}/{property_name.replace(' ', '_')}_Insights_{timestamp}.pdf"

            doc = SimpleDocTemplate(pdf_path, pagesize=A4)
            styles = getSampleStyleSheet()
            story = []

            story.append(Image(LOGO_PATH, width=110, height=55))
            story.append(Spacer(1, 10))

            story.append(Paragraph(
                "Rhythm Hospitality – AI Insights Report",
                styles["Title"]
            ))

            story.append(Spacer(1, 8))

            story.append(Paragraph(f"Property: {property_name}", styles["Normal"]))
            story.append(Paragraph(
                f"Date: {datetime.now().strftime('%d %B %Y')}",
                styles["Normal"]
            ))

            story.append(Spacer(1, 15))

            for line in insights.split("\n"):
                if line.strip():
                    story.append(Paragraph(line, styles["Normal"]))
                    story.append(Spacer(1, 6))

            doc.build(story)

            st.success("Insights saved successfully")

            with open(pdf_path, "rb") as f:
                st.download_button(
                    "Download PDF",
                    data=f,
                    file_name=os.path.basename(pdf_path),
                    mime="application/pdf"
                )
