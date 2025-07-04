import pandas as pd
from fpdf import FPDF
import os
from fpdf.enums import XPos, YPos
import math
import re

# Config
EXCEL_FILE = 'input.xlsx'  # Change if your file is named differently
OUTPUT_DIR = 'output_cvs'

# Ensure output directory exists
os.makedirs(OUTPUT_DIR, exist_ok=True)

def clean_value(val):
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    if s.lower() in ("nan", "none", "null"):
        return ""
    return s

def clean_date(val):
    s = clean_value(val)
    # Remove time if present (e.g., '2023-01-01 00:00:00' -> '2023-01-01')
    if s:
        match = re.match(r"(\d{4}-\d{2}-\d{2})", s)
        if match:
            return match.group(1)
        # Try splitting on space if not ISO
        if " " in s:
            return s.split(" ")[0]
    return s

def get_experiences(row):
    experiences = []
    for i in range(1, 6):
        company = clean_value(row.get(f'Company Name{i if i > 1 else ""}', ""))
        if not company:
            continue
        experiences.append({
            "company": company,
            "job_title": clean_value(row.get(f'Job Title{i if i > 1 else ""}', "")),
            "location": clean_value(row.get(f'Location{i if i > 1 else ""}', "")),
            "start_date": clean_date(row.get(f'Start Date{i if i > 1 else ""}', "")),
            "end_date": clean_date(row.get(f'End Date{i if i > 1 else ""}', "")),
            "responsibility": clean_value(row.get(f'Main Responsibility\xa0{i if i > 1 else ""}', "")),
        })
    return experiences

def get_education(row):
    education = []
    for i in range(1, 6):
        level = clean_value(row.get(f'Education Level{i if i > 1 else ""}', ""))
        if not level:
            continue
        education.append({
            "level": level,
            "institution": clean_value(row.get(f'Institution Name{i if i > 1 else ""}', "")),
            "field": clean_value(row.get(f'Field of study\xa0{i if i > 1 else ""}', "")),
            "start_date": clean_date(row.get(f'Start Date\xa0{i if i > 1 else ""}', "")),
            "end_date": clean_date(row.get(f'End Date{6 + (i-1) if i > 1 else ""}', "")),
            "location": clean_value(row.get(f'Location (City, Country){i if i > 1 else ""}', "")),
        })
    return education

def get_awards(row):
    awards = []
    for i in range(1, 3):
        name = clean_value(row.get(f'Award/Certificate Name{i if i > 1 else ""}', ""))
        if not name:
            continue
        awards.append({
            "name": name,
            "org": clean_value(row.get(f'Issuing Organization{i if i > 1 else ""}', "")),
            "date": clean_date(row.get(f'Date Awarded{i if i > 1 else ""}', "")),
            "desc": clean_value(row.get(f'Award Description (optional){i if i > 1 else ""}', "")),
        })
    return awards

def row_to_context(row):
    return {
        "first_name": clean_value(row.get("First Name", "")),
        "middle_name": clean_value(row.get("Middle Name", "")),
        "last_name": clean_value(row.get("Last Name", "")),
        "email": clean_value(row.get("Personal Email (primary)", "")),
        "phone": clean_value(row.get("Personal Phone Number", "")),
        "address": clean_value(row.get("Full Address", "")),
        "linkedin": clean_value(row.get("LinkedIn Profile\xa0", "")),
        "website": clean_value(row.get("Website / Portfolio (Text)\xa0", "")),
        "dob": clean_date(row.get("Date of Birth", "")),
        "gender": clean_value(row.get("Gender", "")),
        "nationality": clean_value(row.get("Nationality\xa0", "")),
        "summary": clean_value(row.get("About Me / Profile Summary", "")),
        "experiences": get_experiences(row),
        "education": get_education(row),
        "skills": clean_value(row.get("List of Skills and Tools", "")),
        "languages": clean_value(row.get("Language", "")),
        "awards": get_awards(row),
    }

def add_section_header(pdf, text):
    pdf.set_font("DejaVu", "B", 13)
    pdf.set_text_color(30, 30, 60)
    pdf.cell(0, 9, text, align="L", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    # Draw horizontal line
    y = pdf.get_y()
    pdf.set_draw_color(120, 120, 120)
    pdf.set_line_width(0.5)
    pdf.line(10, y, 200, y)
    pdf.ln(2)
    pdf.set_font("DejaVu", "", 9)
    pdf.set_text_color(0, 0, 0)

def split_centered_multiline(pdf, text, font, font_size, max_width):
    # Split text into lines that fit max_width, center each line
    pdf.set_font(font[0], font[1], font_size)
    words = text.split()
    lines = []
    current = ""
    for word in words:
        test = (current + " " + word).strip()
        if pdf.get_string_width(test) > max_width and current:
            lines.append(current)
            current = word
        else:
            current = test
    if current:
        lines.append(current)
    for line in lines:
        pdf.cell(0, 7, line, align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

def get_first_paragraph(text):
    if not text:
        return ""
    # Split on double newline or period+newline or just newline
    for sep in ["\n\n", ".\n", "\n"]:
        if sep in text:
            return text.split(sep)[0].strip()
    return text.strip()

def generate_cv_pdf(context, output_path):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    # Register DejaVu fonts
    pdf.add_font("DejaVu", "", "DejaVuSans.ttf", uni=True)
    pdf.add_font("DejaVu", "B", "DejaVuSans-Bold.ttf", uni=True)
    # Name (centered, bold, large)
    pdf.set_font("DejaVu", "B", 22)
    pdf.set_text_color(30, 30, 60)
    pdf.cell(0, 16, f"{context['first_name']} {context['middle_name']} {context['last_name']}", align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(1)
    # Contact info (centered, thin, small, wrap if too long)
    pdf.set_font("DejaVu", "", 9)
    pdf.set_text_color(80, 80, 80)
    contact_items = [
        context['email'],
        context['phone'],
        context['address'],
        context['linkedin'],
        context['dob']
    ]
    contact_items = [item for item in contact_items if item]
    contact_line = " | ".join(contact_items)
    split_centered_multiline(pdf, contact_line, ("DejaVu", ""), 9, 180)
    pdf.ln(2)
    # Profile Summary (full version)
    add_section_header(pdf, "Profile Summary")
    pdf.set_font("DejaVu", "", 9)
    pdf.set_text_color(0, 0, 0)
    pdf.multi_cell(0, 6, context['summary'])
    pdf.ln(1)
    # Work Experience
    add_section_header(pdf, "Work Experience")
    pdf.set_font("DejaVu", "", 9)
    if context['experiences']:
        for exp in context['experiences']:
            if not exp['job_title'] and not exp['company']:
                continue
            pdf.set_font("DejaVu", "B", 9)
            pdf.cell(0, 6, f"{exp['job_title']} at {exp['company']} ({exp['start_date']} - {exp['end_date']})", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.set_font("DejaVu", "", 9)
            if exp['location']:
                pdf.cell(0, 6, f"Location: {exp['location']}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            if exp['responsibility']:
                pdf.multi_cell(0, 6, f"Responsibilities: {exp['responsibility']}")
            pdf.ln(1)
    else:
        pdf.cell(0, 6, "No work experience provided.", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(1)
    # Education
    add_section_header(pdf, "Education")
    pdf.set_font("DejaVu", "", 9)
    if context['education']:
        for edu in context['education']:
            if not edu['level'] and not edu['institution']:
                continue
            pdf.set_font("DejaVu", "B", 9)
            pdf.cell(0, 6, f"{edu['level']} in {edu['field']}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.set_font("DejaVu", "", 9)
            if edu['institution'] or edu['start_date'] or edu['end_date']:
                pdf.cell(0, 6, f"Institution: {edu['institution']} | Dates: {edu['start_date']} - {edu['end_date']}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            if edu['location']:
                pdf.cell(0, 6, f"Location: {edu['location']}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.ln(1)
    else:
        pdf.cell(0, 6, "No education information provided.", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(1)
    # Skills (comma separated, inline)
    add_section_header(pdf, "Skills & Tools")
    pdf.set_font("DejaVu", "", 9)
    skills = context['skills']
    if skills:
        # Try to split on comma or newline, then join with comma
        if "," in skills:
            skill_list = [s.strip() for s in skills.split(",") if s.strip()]
        else:
            skill_list = [s.strip() for s in skills.split("\n") if s.strip()]
        pdf.multi_cell(0, 6, ", ".join(skill_list))
    pdf.ln(1)
    # Languages
    add_section_header(pdf, "Languages")
    pdf.set_font("DejaVu", "", 9)
    if context['languages']:
        pdf.multi_cell(0, 6, context['languages'])
    pdf.ln(1)
    # Awards
    if context['awards']:
        add_section_header(pdf, "Awards & Certificates")
        pdf.set_font("DejaVu", "", 9)
        for award in context['awards']:
            if not award['name']:
                continue
            pdf.set_font("DejaVu", "B", 9)
            pdf.cell(0, 6, f"{award['name']} from {award['org']} ({award['date']})", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.set_font("DejaVu", "", 9)
            if award['desc']:
                pdf.multi_cell(0, 6, f"Description: {award['desc']}")
            pdf.ln(1)
    # Save PDF
    pdf.output(output_path)

def main():
    df = pd.read_excel(EXCEL_FILE, engine='openpyxl')
    for idx, row in df.iterrows():
        context = row_to_context(row)
        output_pdf_path = os.path.join(
            OUTPUT_DIR,
            f"{context['first_name']}_{context['last_name']}_CV.pdf".replace(" ", "_")
        )
        generate_cv_pdf(context, output_pdf_path)
        print(f"Generated: {output_pdf_path}")
    print("All CVs generated in the output_cvs/ folder.")

if __name__ == "__main__":
    main() 