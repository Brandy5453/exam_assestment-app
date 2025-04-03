import streamlit as st
import os
import pandas as pd
from docx import Document
from sklearn.ensemble import RandomForestRegressor
from sklearn.model_selection import train_test_split
from sklearn.metrics import mean_absolute_error

# ----------------------- STEP 1: EXTRACT FEATURES FROM DOCX ------------------------

def extract_docx_features(doc_path):
    """Extracts text, formatting, images, and text boxes from a DOCX file."""
    doc = Document(doc_path)
    extracted_data = {
        "text": [],
        "bold": [],
        "font_size": [],
        "images": 0,
        "text_boxes": 0
    }

    # Extract paragraphs and formatting
    for para in doc.paragraphs:
        extracted_data["text"].append(para.text)
        
        # Formatting analysis
        bold_count = sum(1 for run in para.runs if run.bold)
        font_sizes = [run.font.size.pt if run.font.size else 12 for run in para.runs]
        
        extracted_data["bold"].append(bold_count)
        extracted_data["font_size"].append(sum(font_sizes) / len(font_sizes) if font_sizes else 12)

    # Count images and text boxes
    for shape in doc.inline_shapes:
        extracted_data["images"] += 1

    return extracted_data


# ----------------------- STEP 2: COMPARE WITH EXPECTED ANSWERS ------------------------

EXPECTED_ANSWERS = {
    "text": [
        "Exciting Internship Opportunity for Aspiring Social Media Influencers!",
        "Strong understanding of various social media platforms (Instagram, TikTok, YouTube, etc.)",
        "Creativity and a knack for storytelling through captivating content",
        "Excellent communication skills, both written and verbal",
        "Ability to collaborate effectively with team members and influencers"
    ],
    "bold": [0, 1, 1, 1, 1],  
    "font_size": [44, 11, 11, 11, 11],
    "images": 1,
    "text_boxes": 1
}

def calculate_score(extracted_data):
    """Compares extracted features with expected values and assigns marks."""
    total_score = 0

    # Check text presence
    for expected_text in EXPECTED_ANSWERS["text"]:
        if any(expected_text in text for text in extracted_data["text"]):
            total_score += 5

    # Check formatting
    total_score += sum(
        4 for i in range(len(EXPECTED_ANSWERS["bold"])) 
        if extracted_data["bold"][i] == EXPECTED_ANSWERS["bold"][i]
    )
    
    total_score += sum(
        4 for i in range(len(EXPECTED_ANSWERS["font_size"])) 
        if round(extracted_data["font_size"][i]) == EXPECTED_ANSWERS["font_size"][i]
    )

    # Check images
    if extracted_data["images"] >= EXPECTED_ANSWERS["images"]:
        total_score += 10

    return total_score


# ----------------------- STEP 3: STREAMLIT APP ------------------------

st.title("📄 Word Document Grader")
st.write("Upload a `.docx` file to get an automated assessment score.")

uploaded_file = st.file_uploader("Upload your Word Document", type=["docx"])

if uploaded_file:
    # Save uploaded file
    file_path = f"temp_{uploaded_file.name}"
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    # Extract features
    extracted_data = extract_docx_features(file_path)

    # Calculate score
    score = calculate_score(extracted_data)

    # Display results
    st.subheader("📊 Extracted Features:")
    st.json(extracted_data)

    st.subheader(f"✅ Predicted Score: {score}/80")

    # Delete temp file
    os.remove(file_path)

