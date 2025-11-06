import streamlit as st
import os
import tempfile
import pandas as pd
from backend import run_hr_batch

st.set_page_config(page_title="HR Resume Matcher", page_icon="🤖", layout="centered")

st.title("🤖 HR Resume Matcher Dashboard")
st.write("Upload candidate resumes (PPTX) and job descriptions (DOCX) to generate a detailed match report.")

# =====================================================
# ✅ FILE UPLOADS
# =====================================================
resume_files = st.file_uploader("📄 Upload Resumes (.pptx)", type=["pptx"], accept_multiple_files=True)
jd_files = st.file_uploader("📝 Upload Job Descriptions (.docx)", type=["docx"], accept_multiple_files=True)

# =====================================================
# ✅ HELPER FUNCTION
# =====================================================
def safe_to_str(x):
    """Converts any list/dict/object to a clean string for safe CSV export."""
    if isinstance(x, (list, tuple)):
        parts = []
        for item in x:
            if isinstance(item, dict):
                parts.append(", ".join(f"{k}: {v}" for k, v in item.items()))
            else:
                parts.append(str(item))
        return "; ".join(parts)
    elif isinstance(x, dict):
        return ", ".join(f"{k}: {v}" for k, v in x.items())
    elif x is None:
        return ""
    else:
        return str(x)

# =====================================================
# ✅ RUN BUTTON
# =====================================================
if st.button("🚀 Run Matching Process"):
    if not resume_files or not jd_files:
        st.warning("Please upload at least one Resume (.pptx) and one Job Description (.docx).")
    else:
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                resume_paths, jd_paths = [], []

                # Save resumes
                for resume in resume_files:
                    resume_path = os.path.join(tmpdir, resume.name)
                    with open(resume_path, "wb") as f:
                        f.write(resume.read())
                    resume_paths.append(resume_path)

                # Save job descriptions
                for jd in jd_files:
                    jd_path = os.path.join(tmpdir, jd.name)
                    with open(jd_path, "wb") as f:
                        f.write(jd.read())
                    jd_paths.append(jd_path)

                st.info("🧠 Processing resumes and job descriptions... Please wait ⏳")

                df, csv_file = run_hr_batch(resume_paths, jd_paths)

                # ✅ Normalize all values to string to avoid conversion errors
                df = df.applymap(safe_to_str)

                st.success("✅ Matching completed successfully!")
                st.dataframe(df)

                # ✅ Recreate clean CSV for download
                final_csv = os.path.join(tmpdir, "final_match_report.csv")
                df.to_csv(final_csv, index=False)

                with open(final_csv, "rb") as file:
                    st.download_button(
                        label="📥 Download Final Report (CSV)",
                        data=file,
                        file_name="final_match_report.csv",
                        mime="text/csv"
                    )

        except Exception as e:
            st.error(f"❌ Error Occurred: {str(e)}")

# =====================================================
# ✅ FOOTER
# =====================================================
st.markdown("---")
st.caption("Built by Shivanagouda Patil 🧠 | Powered by LlaMa HR AI Engine")
