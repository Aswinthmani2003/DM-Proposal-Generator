import streamlit as st
from docx import Document
from datetime import datetime
import os
from docx.oxml.ns import qn
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
import uuid
import tempfile

# Proposal configurations
PROPOSAL_CONFIG = {
    "Shopify Website Development": {
        "template": "Shopify Website.docx",
        "pricing_fields": [
            ("Development", "Dev-Price"),
            ("Design", "Design-Price"),
            ("Testing and Live", "Testing-Price")
        ],
        "team_type": "general",
        "special_fields": [("VDate", "<<")],
        "proposal_type": "ecommerce"
    },
    "Single Vendor Ecommerce Website": {
        "template": "Single Vendor Ecommerce website.docx",
        "pricing_fields": [
            ("Development", "Dev-Price"),
            ("Design", "Design-Price"),
            ("Website Bot", "WB-Price"),
            ("Testing and Deployment", "TD-Price")
        ],
        "team_type": "general",
        "special_fields": [("VDate", "<<")],
        "proposal_type": "ecommerce"
    },
    "DM Proposal - All": {
        "template": "DM Proposal - All.docx",
        "pricing_fields": [
            ("Marketing Strategy", "MS"),
            ("Social Media Setup", "SM"),
            ("Meta & Google Ads Setup", "MG"),
            ("Creative Posts", "CP"),
            ("Meta Paid Ads", "MP"),
            ("Google Paid Ads", "GP"),
            ("SEO", "SEO"),
            ("Email Marketing", "EM"),
            ("Monthly Reporting", "MR")
        ],
        "team_type": "dm",
        "special_fields": [("VDate", "<<")],
        "proposal_type": "digital_marketing"
    },
    "Web Based AI Fintech": {
        "template": "Web based AI Fintech proposal.docx",
        "pricing_fields": [
            ("Development", "Dev-Price"),
            ("Design", "Design-Price"),
            ("AI/ML Models", "AIML-Price")
        ],
        "team_type": "general",
        "special_fields": [("VDate", "<<")],
        "proposal_type": "fintech"
    },
    "Community App Tech Proposal": {
        "template": "Community App Tech Proposal.docx",
        "pricing_fields": [
            ("Design", "Design-Price"),
            ("AI/ML & Development", "AIML-Price"),
            ("QA & Project Management", "QA-Price")
        ],
        "team_type": "general",
        "special_fields": [("VDate", "<<")],
        "proposal_type": "general"
    },
    "Job portal website Tech Proposal": {
        "template": "Job portal website Tech Proposal.docx",
        "pricing_fields": [
            ("Design", "Design-Price"),
            ("Development", "Dev-Price"),
            ("Automations", "Automation-Price"),
            ("Testing & Deployment", "TD-Price")
        ],
        "team_type": "general",
        "special_fields": [("VDate", "<<")],
        "proposal_type": "general"
    },
    "AI Based Search Engine Website Technical Consultation": {
        "template": "AI Based Search Engine Website Technical Consultation proposal.docx",
        "pricing_fields": [
            ("Designs", "Design-Price"),
            ("Development", "Dev-Price"),
            ("Testing & Deployment", "TD-Price")
        ],
        "team_type": "general",
        "special_fields": [("VDate", "<<")],
        "proposal_type": "ai_search"
    }
}

def apply_formatting(new_run, original_run):
    """Copy formatting from original run to new run"""
    if original_run.font.name:
        new_run.font.name = original_run.font.name
        new_run._element.rPr.rFonts.set(qn('w:eastAsia'), original_run.font.name)
    if original_run.font.size:
        new_run.font.size = original_run.font.size
    if original_run.font.color.rgb:
        new_run.font.color.rgb = original_run.font.color.rgb
    new_run.bold = original_run.bold
    new_run.italic = original_run.italic

def replace_in_paragraph(para, placeholders):
    """Handle paragraph replacements preserving formatting"""
    original_runs = para.runs.copy()
    full_text = para.text
    for ph, value in placeholders.items():
        full_text = full_text.replace(ph, str(value))

    if full_text != para.text:
        para.clear()
        new_run = para.add_run(full_text)
        if original_runs:
            original_run = next((r for r in original_runs if r.text), None)
            if original_run:
                apply_formatting(new_run, original_run)

def replace_and_format(doc, placeholders):
    """Enhanced replacement with table cell handling"""
    for para in doc.paragraphs:
        replace_in_paragraph(para, placeholders)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.tables:
                    for nested_table in cell.tables:
                        for nested_row in nested_table.rows:
                            for nested_cell in nested_row.cells:
                                for para in nested_cell.paragraphs:
                                    replace_in_paragraph(para, placeholders)
                else:
                    for para in cell.paragraphs:
                        replace_in_paragraph(para, placeholders)
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    return doc

def get_general_team_details():
    """Collect team composition for general projects"""
    st.subheader("Team Composition")
    team_roles = {
        "Project Manager": "P1",
        "Business Analyst": "B1",
        "UI/UX Members": "U1",
        "Backend Developers": "BD1",
        "Frontend Developers": "F1",
        "AI/ML Developers": "A1",
        "System Architect": "S1",
        "AWS Developer": "AD1"
    }
    team_details = {}
    cols = st.columns(2)

    for idx, (role, placeholder) in enumerate(team_roles.items()):
        with cols[idx % 2]:
            count = st.number_input(
                f"{role} Count:",
                min_value=0,
                step=1,
                key=f"team_{placeholder}"
            )
            team_details[f"<<{placeholder}>>"] = str(count)
    return team_details

def get_dm_team_details():
    """Collect team composition for DM projects"""
    st.subheader("Team Composition")
    team_roles = {
        "Digital Marketing Executive": "DME",
        "Digital Marketing Associate": "DMA",
        "Business Analyst": "BA",
        "Graphics Designer": "GD"
    }
    team_details = {}
    cols = st.columns(2)

    for idx, (role, placeholder) in enumerate(team_roles.items()):
        with cols[idx % 2]:
            count = st.number_input(
                f"{role} Count:",
                min_value=0,
                step=1,
                key=f"team_{placeholder}"
            )
            team_details[f"<<{placeholder}>>"] = str(count)
    return team_details

def remove_empty_rows(table):
    """Remove rows from the table where the pricing cell is empty or zero"""
    rows_to_remove = []
    for row in table.rows:
        if row.cells[0].text.strip().lower() == 'description':
            continue
            
        if len(row.cells) > 2:
            price_cell = row.cells[2].text.strip()
            if price_cell in {"", "$0", "₹0", "0", "<<>>"}:
                rows_to_remove.append(row)
    
    for row in reversed(rows_to_remove):
        table._tbl.remove(row._element)

def validate_phone_number(country, phone_number):
    """Validate phone number format"""
    if country.lower() == "india":
        return phone_number.startswith("+91")
    return phone_number.startswith("+1")

def format_number_with_commas(number):
    return f"{number:,}"

def generate_document():
    st.title("DM Proposal Generator")
    base_dir = os.getcwd()

    selected_proposal = st.selectbox("Select Proposal", list(PROPOSAL_CONFIG.keys()))
    config = PROPOSAL_CONFIG[selected_proposal]
    template_path = os.path.join(base_dir, config["template"])


    # Client Information
    col1, col2 = st.columns(2)
    with col1:
        client_name = st.text_input("Client Name:")
        client_email = st.text_input("Client Email:")
    with col2:
        if selected_proposal != "DM Proposal - All":
            country = st.text_input("Country:")
        else:
            country = ""
        client_number = st.text_input("Client Number:")
        if client_number and country and not validate_phone_number(country, client_number):
            st.error(f"Phone number should start with {'+91' if country.lower() == 'india' else '+1'}")

    date_field = st.date_input("Date:", datetime.today())

    # Currency Handling
    currency = st.selectbox("Select Currency", ["USD", "INR"])
    currency_symbol = "$" if currency == "USD" else "₹"

    # Special Fields
    special_data = {}
    st.subheader("Additional Details")
    vdate = st.date_input("Proposal Validity Until:")
    special_data["<<VDate>>"] = vdate.strftime("%d-%m-%Y")

    # Pricing Section
    st.subheader("Pricing Details")
    pricing_data = {}
    numerical_values = {}
    
    pricing_fields = config["pricing_fields"]
    num_fields = len(pricing_fields)
    num_rows = (num_fields + 1) // 2

    for row in range(num_rows):
        cols = st.columns(2)
        for col in range(2):
            idx = row * 2 + col
            if idx < num_fields:
                label, key = pricing_fields[idx]
                with cols[col]:
                    value = st.number_input(
                        f"{label} ({currency})",
                        min_value=0,
                        value=0,
                        step=100,
                        format="%d",
                        key=f"price_{key}"
                    )
                    numerical_values[key] = value
                    pricing_data[f"<<{key}>>"] = f"{currency_symbol}{format_number_with_commas(value)}" if value > 0 else ""

    # Payment Schedule Section
    instalment_data = {}
    if selected_proposal == "DM Proposal - All":
        st.subheader("Payment Schedule")
        cols = st.columns(2)
        with cols[0]:
            instalment1 = st.number_input(
                f"Instalment 1 ({currency})",
                min_value=0,
                value=0,
                step=100,
                format="%d",
                key="instalment_1"
            )
        with cols[1]:
            instalment2 = st.number_input(
                f"Instalment 2 ({currency})",
                min_value=0,
                value=0,
                step=100,
                format="%d",
                key="instalment_2"
            )
        instalment_data = {
            "<<Instalment 1>>": f"{currency_symbol}{format_number_with_commas(instalment1)}",
            "<<Instalment 2>>": f"{currency_symbol}{format_number_with_commas(instalment2)}"
        }

    # Total Calculation
    services_sum = sum(numerical_values.values())
    if config["proposal_type"] == "digital_marketing":
        total = services_sum
        pricing_data["<<Total>>"] = f"{currency_symbol}{format_number_with_commas(total)}"
    else:
        am_price = int(services_sum * 0.10)
        pricing_data["<<AM-Price>>"] = f"{currency_symbol}{format_number_with_commas(am_price)}"
        total = services_sum + am_price
        if currency == "INR":
            pricing_data["<<T-Price>>"] = f"{currency_symbol}{format_number_with_commas(total)} + 18% GST"
        else:
            pricing_data["<<T-Price>>"] = f"{currency_symbol}{format_number_with_commas(total)}"
        af_price = 250 if currency == "USD" else 25000
        pricing_data["<<AF-Price>>"] = f"{currency_symbol}{format_number_with_commas(af_price)}"

    # Team Composition
    if config["team_type"] == "dm":
        team_data = get_dm_team_details()
    elif config["team_type"] == "general":
        team_data = get_general_team_details()
    else:
        team_data = {}

    # Combine all placeholders
    placeholders = {
        "<<Client Name>>": client_name,
        "<<Client Email>>": client_email,
        "<<Client Number>>": client_number,
        "<<Date>>": date_field.strftime("%d-%m-%Y"),
        "<<Country>>": country
    }
    placeholders.update(pricing_data)
    placeholders.update(team_data)
    placeholders.update(special_data)
    placeholders.update(instalment_data)

    if st.button("Generate Proposal"):
        if client_number and country and not validate_phone_number(country, client_number):
            st.error("Invalid phone number format!")
        else:
            doc_filename = f"Technical Consultation Proposal - {client_name} {date_field.strftime('%d-%m-%Y')}.docx"

            with tempfile.TemporaryDirectory() as temp_dir:
                try:
                    doc = Document(template_path)
                except FileNotFoundError:
                    st.error(f"Template not found: {template_path}")
                    return

                doc = replace_and_format(doc, placeholders)

                if selected_proposal == "DM Proposal - All":
                    for table in doc.tables:
                        remove_empty_rows(table)

                doc_path = os.path.join(temp_dir, doc_filename)
                doc.save(doc_path)

                with open(doc_path, "rb") as f:
                    st.download_button(
                        label="Download Proposal",
                        data=f,
                        file_name=doc_filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )


if __name__ == "__main__":
    generate_document()