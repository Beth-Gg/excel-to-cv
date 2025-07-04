import pandas as pd
from fpdf import FPDF
import os
from fpdf.enums import XPos, YPos

# Config
EXCEL_FILE = 'input.xlsx'  # Change if your file is named differently
OUTPUT_DIR = 'output_cvs'

# Ensure output directory exists
os.makedirs(OUTPUT_DIR, exist_ok=True)

def get_experiences(row):
    experiences = []
    for i in range(1, 6):
        company = row.get(f'Company Name{i if i > 1 else ""}', "")
        if pd.isna(company) or company == "":
            continue
        experiences.append({
            "company": company,
            "job_title": row.get(f'Job Title{i if i > 1 else ""}', ""),
            "location": row.get(f'Location{i if i > 1 else ""}', ""),
            "start_date": row.get(f'Start Date{i if i > 1 else ""}', ""),
            "end_date": row.get(f'End Date{i if i > 1 else ""}', ""),
            "responsibility": row.get(f'Main Responsibility\xa0{i if i > 1 else ""}', ""),
        })
    return experiences

def get_education(row):
    education = []
    for i in range(1, 6):
        level = row.get(f'Education Level{i if i > 1 else ""}', "")
        if pd.isna(level) or level == "":
            continue
        education.append({
            "level": level,
            "institution": row.get(f'Institution Name{i if i > 1 else ""}', ""),
            "field": row.get(f'Field of study\xa0{i if i > 1 else ""}', ""),
            "start_date": row.get(f'Start Date\xa0{i if i > 1 else ""}', ""),
            "end_date": row.get(f'End Date{6 + (i-1) if i > 1 else ""}', ""),
            "location": row.get(f'Location (City, Country){i if i > 1 else ""}', ""),
        })
    return education

def get_awards(row):
    awards = []
    for i in range(1, 3):
        name = row.get(f'Award/Certificate Name{i if i > 1 else ""}', "")
        if pd.isna(name) or name == "":
            continue
        awards.append({
            "name": name,
            "org": row.get(f'Issuing Organization{i if i > 1 else ""}', ""),
            "date": row.get(f'Date Awarded{i if i > 1 else ""}', ""),
            "desc": row.get(f'Award Description (optional){i if i > 1 else ""}', ""),
        })
    return awards

def row_to_context(row):
    return {
        "first_name": row.get("First Name", ""),
        "middle_name": row.get("Middle Name", ""),
        "last_name": row.get("Last Name", ""),
        "email": row.get("Personal Email (primary)", ""),
        "phone": row.get("Personal Phone Number", ""),
        "address": row.get("Full Address", ""),
        "linkedin": row.get("LinkedIn Profile\xa0", ""),
        "website": row.get("Website / Portfolio (Text)\xa0", ""),
        "dob": row.get("Date of Birth", ""),
        "gender": row.get("Gender", ""),
        "nationality": row.get("Nationality\xa0", ""),
        "summary": row.get("About Me / Profile Summary", ""),
        "experiences": get_experiences(row),
        "education": get_education(row),
        "skills": row.get("List of Skills and Tools", ""),
        "languages": row.get("Language", ""),
        "awards": get_awards(row),
    }

def generate_cv_pdf(context, output_path):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_font("Pavonine", "", "pavonine.ttf", uni=True)
    # Header
    pdf.set_font("Pavonine", "", 22)
    pdf.set_text_color(40, 40, 80)
    pdf.cell(0, 14, f"{context['first_name']} {context['middle_name']} {context['last_name']}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font("Pavonine", "", 12)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(0, 8, f"Email: {context['email']} | Phone: {context['phone']}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 8, f"Address: {context['address']}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    if context['linkedin'] or context['website']:
        links = []
        if context['linkedin']:
            links.append(f"LinkedIn: {context['linkedin']}")
        if context['website']:
            links.append(f"Website: {context['website']}")
        pdf.cell(0, 8, " | ".join(links), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 8, f"Date of Birth: {context['dob']} | Gender: {context['gender']} | Nationality: {context['nationality']}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(4)
    # Profile Summary
    pdf.set_font("Pavonine", "", 14)
    pdf.set_text_color(40, 40, 80)
    pdf.cell(0, 10, "Profile Summary", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font("Pavonine", "", 12)
    pdf.set_text_color(0, 0, 0)
    pdf.multi_cell(0, 8, str(context['summary']))
    pdf.ln(2)
    # Work Experience
    pdf.set_font("Pavonine", "", 14)
    pdf.set_text_color(40, 40, 80)
    pdf.cell(0, 10, "Work Experience", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font("Pavonine", "", 12)
    pdf.set_text_color(0, 0, 0)
    if context['experiences']:
        for exp in context['experiences']:
            pdf.set_font("Pavonine", "", 12)
            pdf.cell(0, 8, f"{exp['job_title']} at {exp['company']} ({exp['start_date']} - {exp['end_date']})", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.cell(0, 8, f"Location: {exp['location']}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.multi_cell(0, 8, f"Responsibilities: {exp['responsibility']}")
            pdf.ln(1)
    else:
        pdf.cell(0, 8, "No work experience provided.", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(2)
    # Education
    pdf.set_font("Pavonine", "", 14)
    pdf.set_text_color(40, 40, 80)
    pdf.cell(0, 10, "Education", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font("Pavonine", "", 12)
    pdf.set_text_color(0, 0, 0)
    if context['education']:
        for edu in context['education']:
            pdf.set_font("Pavonine", "", 12)
            pdf.cell(0, 8, f"{edu['level']} in {edu['field']}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.cell(0, 8, f"Institution: {edu['institution']} | Dates: {edu['start_date']} - {edu['end_date']}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.cell(0, 8, f"Location: {edu['location']}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.ln(1)
    else:
        pdf.cell(0, 8, "No education information provided.", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(2)
    # Skills
    pdf.set_font("Pavonine", "", 14)
    pdf.set_text_color(40, 40, 80)
    pdf.cell(0, 10, "Skills & Tools", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font("Pavonine", "", 12)
    pdf.set_text_color(0, 0, 0)
    pdf.multi_cell(0, 8, str(context['skills']))
    pdf.ln(2)
    # Languages
    pdf.set_font("Pavonine", "", 14)
    pdf.set_text_color(40, 40, 80)
    pdf.cell(0, 10, "Languages", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font("Pavonine", "", 12)
    pdf.set_text_color(0, 0, 0)
    pdf.multi_cell(0, 8, str(context['languages']))
    pdf.ln(2)
    # Awards
    if context['awards']:
        pdf.set_font("Pavonine", "", 14)
        pdf.set_text_color(40, 40, 80)
        pdf.cell(0, 10, "Awards & Certificates", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_font("Pavonine", "", 12)
        pdf.set_text_color(0, 0, 0)
        for award in context['awards']:
            pdf.cell(0, 8, f"{award['name']} from {award['org']} ({award['date']})", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.multi_cell(0, 8, f"Description: {award['desc']}")
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