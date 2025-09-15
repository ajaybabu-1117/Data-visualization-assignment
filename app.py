import streamlit as st
import pandas as pd
import plotly.express as px
import pdfplumber
from docx import Document
from wordcloud import WordCloud
import matplotlib.pyplot as plt

# -------------------------------
# App Config
# -------------------------------
st.set_page_config(
    page_title="Data Visualizer App",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("Data Visualizer & Analyzer")

# -------------------------------
# Sidebar
# -------------------------------
st.sidebar.header("Upload Your File")
uploaded_file = st.sidebar.file_uploader(
    "Upload Excel, CSV, PDF, or Word",
    type=["xlsx", "xls", "csv", "pdf", "docx"],
    help="Max size: 200MB"
)

st.sidebar.header("Visualization Options")
chart_type = st.sidebar.selectbox(
    "Choose Chart Type (for tabular data)",
    ["Bar Chart", "Pie Chart", "Line Chart", "Scatter Plot"]
)

# -------------------------------
# Helper Functions
# -------------------------------
@st.cache_data
def get_excel_sheet_names(file, engine="openpyxl"):
    """Return sheet names from an Excel file."""
    try:
        xls = pd.ExcelFile(file, engine=engine)
    except Exception:
        xls = pd.ExcelFile(file, engine="xlrd")
    return xls.sheet_names

@st.cache_data
def read_excel(file, sheet_name, engine="openpyxl"):
    """Read a specific sheet into a DataFrame."""
    try:
        return pd.read_excel(file, sheet_name=sheet_name, engine=engine)
    except Exception:
        return pd.read_excel(file, sheet_name=sheet_name, engine="xlrd")

@st.cache_data
def read_csv(file):
    """Read CSV into DataFrame."""
    return pd.read_csv(file)

@st.cache_data
def extract_text_from_pdf(file, start_page=0, end_page=5):
    """Extract text from PDF within page range."""
    text = ""
    with pdfplumber.open(file) as pdf:
        total_pages = len(pdf.pages)
        end_page = min(end_page, total_pages)
        for i in range(start_page, end_page):
            page = pdf.pages[i]
            text += page.extract_text() or ""
    return text

@st.cache_data
def extract_text_from_docx(file):
    """Extract text from a Word document."""
    doc = Document(file)
    return " ".join([p.text for p in doc.paragraphs if p.text.strip()])

def generate_wordcloud(text):
    """Generate a WordCloud object."""
    wc = WordCloud(width=800, height=400, background_color="white", max_words=200).generate(text)
    return wc

# -------------------------------
# Main Logic
# -------------------------------
if uploaded_file:
    file_type = uploaded_file.name.split(".")[-1].lower()

    if file_type in ["xlsx", "xls"]:
        # Excel handling
        sheet_names = get_excel_sheet_names(uploaded_file)
        selected_sheet = st.sidebar.selectbox("Choose a Sheet", sheet_names)
        df = read_excel(uploaded_file, selected_sheet)

        st.subheader(f"Data Preview: {selected_sheet}")
        st.dataframe(df.head())

        # Column selection
        cols = st.multiselect("Select Columns to Visualize", df.columns)
        if len(cols) >= 2:
            x_axis, y_axis = cols[0], cols[1]

            if chart_type == "Bar Chart":
                fig = px.bar(df, x=x_axis, y=y_axis, title=f"{y_axis} by {x_axis}")
            elif chart_type == "Pie Chart":
                fig = px.pie(df, names=x_axis, values=y_axis, title=f"{y_axis} distribution by {x_axis}")
            elif chart_type == "Line Chart":
                fig = px.line(df, x=x_axis, y=y_axis, title=f"{y_axis} trend by {x_axis}")
            elif chart_type == "Scatter Plot":
                fig = px.scatter(df, x=x_axis, y=y_axis, title=f"{y_axis} vs {x_axis}")

            st.plotly_chart(fig, use_container_width=True)

    elif file_type == "csv":
        # CSV handling
        df = read_csv(uploaded_file)
        st.subheader("Data Preview")
        st.dataframe(df.head())

        cols = st.multiselect("Select Columns to Visualize", df.columns)
        if len(cols) >= 2:
            x_axis, y_axis = cols[0], cols[1]

            if chart_type == "Bar Chart":
                fig = px.bar(df, x=x_axis, y=y_axis)
            elif chart_type == "Pie Chart":
                fig = px.pie(df, names=x_axis, values=y_axis)
            elif chart_type == "Line Chart":
                fig = px.line(df, x=x_axis, y=y_axis)
            elif chart_type == "Scatter Plot":
                fig = px.scatter(df, x=x_axis, y=y_axis)

            st.plotly_chart(fig, use_container_width=True)

    elif file_type == "pdf":
        # PDF text extraction
        st.subheader("Extracted Text from PDF")
        text = extract_text_from_pdf(uploaded_file, 0, 5)
        st.write(text[:2000] + "..." if len(text) > 2000 else text)

        if text:
            wc = generate_wordcloud(text)
            st.subheader("Word Cloud from PDF")
            fig, ax = plt.subplots(figsize=(10, 5))
            ax.imshow(wc, interpolation="bilinear")
            ax.axis("off")
            st.pyplot(fig)

    elif file_type == "docx":
        # Word document handling
        st.subheader("Extracted Text from Word Document")
        text = extract_text_from_docx(uploaded_file)
        st.write(text[:2000] + "..." if len(text) > 2000 else text)

        if text:
            wc = generate_wordcloud(text)
            st.subheader("Word Cloud from Word Document")
            fig, ax = plt.subplots(figsize=(10, 5))
            ax.imshow(wc, interpolation="bilinear")
            ax.axis("off")
            st.pyplot(fig)

else:
    st.info("Please upload a file from the sidebar to get started.")
