'''
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
'''


import streamlit as st
import pandas as pd
from docx import Document
import spacy
from transformers import pipeline
import re
from openpyxl import load_workbook
from difflib import SequenceMatcher

# Initialize NLP models
try:
    nlp = spacy.load("en_core_web_sm")
except:
    st.warning("Couldn't load spaCy model, using simpler text processing")
    nlp = None

try:
    zero_shot_classifier = pipeline("zero-shot-classification", model="facebook/bart-large-mnli")
except:
    st.warning("Couldn't load zero-shot classifier, using simpler matching")
    zero_shot_classifier = None

def load_rubric(file):
    """Load rubric from Excel in standard format"""
    try:
        wb = load_workbook(filename=file, read_only=True)
        sheet = wb.active
        
        rubric_data = []
        for row in sheet.iter_rows(values_only=True):
            if row and row[0] and isinstance(row[0], str) and not row[0].startswith("Criteria"):
                rubric_data.append({
                    'Criteria': row[0],
                    'Total Points': float(row[1]) if isinstance(row[1], (int, float)) else 0.0,
                    'Excellent': row[2] if len(row) > 2 else "",
                    'Good': row[3] if len(row) > 3 else "",
                    'Poor': row[4] if len(row) > 4 else ""
                })
        
        return pd.DataFrame(rubric_data)
    except Exception as e:
        st.error(f"Error loading rubric: {str(e)}")
        return None

def extract_document_features(doc_file):
    """Extract features from Word document that can be matched against rubric"""
    try:
        doc = Document(doc_file)
        features = {
            'text': "\n".join([para.text for para in doc.paragraphs]),
            'has_images': len(doc.inline_shapes) > 0,
            'headings': [para.text for para in doc.paragraphs if para.style.name.startswith('Heading')],
            'lists': detect_lists(doc),
            'formatting': detect_formatting(doc),
            'tables': len(doc.tables),
            'sections': len(doc.sections)
        }
        return features
    except Exception as e:
        st.error(f"Error processing document: {str(e)}")
        return None

def detect_lists(doc):
    """Detect numbered/bulleted lists in document"""
    lists = []
    for para in doc.paragraphs:
        if para.style.name.lower().startswith(('list', 'heading')) or para.text.strip().startswith(('•', '-', '*', '1.', 'a)', 'I.')):
            lists.append(para.text)
    return lists

def detect_formatting(doc):
    """Detect basic formatting features"""
    formatting = {
        'bold': [],
        'italic': [],
        'underlined': [],
        'colored_text': []
    }
    for para in doc.paragraphs:
        for run in para.runs:
            if run.bold:
                formatting['bold'].append(run.text)
            if run.italic:
                formatting['italic'].append(run.text)
            if run.underline:
                formatting['underlined'].append(run.text)
            if hasattr(run.font, 'color') and run.font.color.rgb is not None:
                formatting['colored_text'].append(run.text)
    return formatting

def similarity_score(a, b):
    """Calculate text similarity score"""
    return SequenceMatcher(None, a.lower(), b.lower()).ratio()

def match_criteria(student_features, criteria_text):
    """Use NLP to match student work against rubric criteria"""
    if not student_features or not criteria_text:
        return 0.0
    
    # First try exact matches
    if criteria_text.lower() in student_features['text'].lower():
        return 1.0
    
    # Try matching with headings
    for heading in student_features['headings']:
        if similarity_score(heading, criteria_text) > 0.7:
            return 1.0
    
    # Use zero-shot classification if available
    if zero_shot_classifier:
        try:
            result = zero_shot_classifier(student_features['text'], [criteria_text])
            return result['scores'][0]
        except:
            pass
    
    # Fallback to simple text similarity
    return similarity_score(student_features['text'], criteria_text)

def grade_assignment(student_features, rubric_df):
    """Grade assignment against rubric"""
    if student_features is None:
        return [], 0.0, 0.0
    
    results = []
    total_possible = rubric_df['Total Points'].sum()
    total_earned = 0.0
    
    for _, row in rubric_df.iterrows():
        criteria = row['Criteria']
        max_points = float(row['Total Points'])
        
        # Skip criteria with 0 points
        if max_points <= 0:
            results.append({
                'Criteria': criteria,
                'Points Possible': 0,
                'Points Earned': 0,
                'Feedback': "No points allocated",
                'Match %': "N/A"
            })
            continue
        
        # Special handling for different criteria types
        if "insert" in criteria.lower() or "add" in criteria.lower():
            # Content insertion criteria
            match_score = match_criteria(student_features, criteria)
            points = max_points * match_score
            feedback = f"Content match: {match_score:.1%}"
        
        elif "format" in criteria.lower() or "style" in criteria.lower():
            # Formatting criteria
            points = assess_formatting(student_features, criteria, max_points)
            feedback = "Formatting assessed"
        
        elif "image" in criteria.lower() or "picture" in criteria.lower():
            # Image-related criteria
            points = assess_images(student_features, criteria, max_points)
            feedback = "Image requirements assessed"
        
        elif "table" in criteria.lower():
            # Table-related criteria
            points = assess_tables(student_features, criteria, max_points)
            feedback = "Table requirements assessed"
        
        else:
            # Generic criteria
            match_score = match_criteria(student_features, criteria)
            points = max_points * match_score
            feedback = f"Generic criteria match: {match_score:.1%}"
        
        results.append({
            'Criteria': criteria,
            'Points Possible': max_points,
            'Points Earned': round(points, 1),
            'Feedback': feedback,
            'Match %': f"{min(100, int(points/max_points*100))}%" if max_points > 0 else "N/A"
        })
        total_earned += points
    
    return results, total_earned, total_possible

def assess_formatting(features, criteria, max_points):
    """Assess formatting requirements"""
    # Check for bold
    if "bold" in criteria.lower() and features['formatting']['bold']:
        return max_points * 0.8
    
    # Check for italics
    if "italic" in criteria.lower() and features['formatting']['italic']:
        return max_points * 0.8
    
    # Check for colored text
    if "color" in criteria.lower() and features['formatting']['colored_text']:
        return max_points * 0.5
    
    # Check for underline
    if "underline" in criteria.lower() and features['formatting']['underlined']:
        return max_points * 0.5
    
    # Default partial credit for formatting we can't specifically verify
    return max_points * 0.3

def assess_images(features, criteria, max_points):
    """Assess image-related requirements"""
    if not features['has_images']:
        return 0.0
    
    # Basic image presence
    if "insert" in criteria.lower() or "add" in criteria.lower():
        return max_points
    
    # More specific image requirements get partial credit
    return max_points * 0.7

def assess_tables(features, criteria, max_points):
    """Assess table-related requirements"""
    if features['tables'] == 0:
        return 0.0
    
    # Basic table presence
    if "insert" in criteria.lower() or "add" in criteria.lower():
        return max_points
    
    # More specific table requirements get partial credit
    return max_points * 0.7

def display_results(results, total_earned, total_possible, student_info):
    """Display grading results"""
    st.subheader("Grading Results")
    st.write(f"**Student:** {student_info['name']} | **ID:** {student_info['id']}")
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Total Score", f"{total_earned:.1f}/{total_possible:.1f}")
    with col2:
        percentage = (total_earned/total_possible)*100 if total_possible > 0 else 0
        st.metric("Percentage", f"{percentage:.1f}%")
    
    st.subheader("Detailed Criteria Assessment")
    results_df = pd.DataFrame(results)
    
    # Display the dataframe with formatting
    st.dataframe(results_df.style
                .highlight_max(axis=0, subset=['Points Earned'])
                .format({'Points Possible': '{:.1f}', 'Points Earned': '{:.1f}'}))
    
    # Add visualizations (only for rows with numeric Match %)
    st.subheader("Performance Breakdown")
    if not results_df.empty:
        # Create a copy for visualization
        viz_df = results_df.copy()
        viz_df['Match Numeric'] = viz_df['Match %'].apply(
            lambda x: float(x.replace('%','')) if x != 'N/A' else 0)
        
        col1, col2 = st.columns(2)
        with col1:
            st.bar_chart(viz_df.set_index('Criteria')['Points Earned'])
        with col2:
            st.bar_chart(viz_df.set_index('Criteria')['Match Numeric'])

def main():
    st.set_page_config(page_title="University of South Africa - EUP Tool Grading System", layout="wide")
    st.title("University of South Africa - EUP Tool Grading System")
    
    # st.logo
    # Initialize session state
    if 'graded_assignments' not in st.session_state:
        st.session_state.graded_assignments = []
    
    # File upload
    with st.expander("Upload Files", expanded=True):
        col1, col2 = st.columns(2)
        with col1:
            rubric_file = st.file_uploader("Upload Rubric (Excel)", type=["xlsx"])
        with col2:
            assignment_file = st.file_uploader("Upload Student Assignment (Word)", type=["docx"])
    
    # Student info
    student_info = {}
    if rubric_file and assignment_file:
        with st.expander("Student Information", expanded=True):
            student_info['name'] = st.text_input("Student Name", key="student_name")
            student_info['id'] = st.text_input("Student ID", key="student_id")
    
    # Process when ready
    if rubric_file and assignment_file and student_info.get('name') and student_info.get('id'):
        rubric_df = load_rubric(rubric_file)
        if rubric_df is not None:
            with st.spinner("Analyzing assignment..."):
                student_features = extract_document_features(assignment_file)
                if student_features is not None:
                    results, total_earned, total_possible = grade_assignment(student_features, rubric_df)
                    
                    display_results(results, total_earned, total_possible, student_info)
                    
                    # Add to session state
                    st.session_state.graded_assignments.append({
                        'student_info': student_info,
                        'results': results,
                        'total_earned': total_earned,
                        'total_possible': total_possible
                    })
                    
                    # Export options
                    csv_data = pd.DataFrame(results).to_csv(index=False)
                    st.download_button(
                        label="Download Results as CSV",
                        data=csv_data,
                        file_name=f"grading_results_{student_info['id']}.csv",
                        mime="text/csv"
                    )
    
    # Display all graded assignments
    if st.session_state.graded_assignments:
        st.subheader("All Graded Assignments")
        summary_data = []
        for assignment in st.session_state.graded_assignments:
            summary_data.append({
                'Student Name': assignment['student_info']['name'],
                'Student ID': assignment['student_info']['id'],
                'Total Score': f"{assignment['total_earned']:.1f}/{assignment['total_possible']:.1f}",
                'Percentage': f"{(assignment['total_earned']/assignment['total_possible'])*100:.1f}%"
            })
        
        st.dataframe(pd.DataFrame(summary_data))

if __name__ == "__main__":
    main()
