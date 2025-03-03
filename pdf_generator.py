import os
from pathlib import Path
from bs4 import BeautifulSoup
import re
from reportlab.lib.pagesizes import letter
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch

from num2words import num2words

def extract_data(html: str) -> dict:
    """Extract structured data from HTML content."""
    soup = BeautifulSoup(html, 'html.parser')
    data = {
        "object": "X",
        "company": "X",
        "acheteur": "X",
        "location": "X",
        "total_ht": "0",
        "total_tva": "0",
        "total_ttc": "0",
    }
    try:
        acheteur_tag = soup.find('span', string='Acheteur public')
        if acheteur_tag:
            data["acheteur"] = acheteur_tag.find_next('span').get_text(strip=True)
        cert_section = soup.find(string=re.compile(r'O\s*:'))
        if cert_section:
            data["company"] = re.search(r'O\s*:\s*(.*?)(<br>|$)', str(cert_section)).group(1).strip()
        obj_section = soup.find('span', class_='text-uppercase', string='Objet')
        if obj_section:
            data["object"] = obj_section.find_next('span', class_='text-black').get_text(strip=True)
        location_tag = soup.find('span', string=re.compile(r'Lieu d[\'’]exécution', re.IGNORECASE))
        if location_tag:
            data["location"] = location_tag.find_next('span').get_text(strip=True)
        totals = {
            "Total HT": "total_ht",
            "Total TVA": "total_tva",
            "Total TTC": "total_ttc"
        }
        for label, key in totals.items():
            elem = soup.find('p', string=label)
            if elem:
                value = elem.find_next('p').get_text(strip=True)
                data[key] = re.sub(r'[^\d]', '', value.split(',')[0])
        return data
    except Exception as e:
        print(f"Extraction error: {str(e)}")
        return data

def extract_table_data(html_content: str) -> list:
    """Extract table data from HTML content, skipping rows with total keywords."""
    soup = BeautifulSoup(html_content, 'html.parser')
    table = soup.find('table', id='article--table')
    if not table:
        return []
    headers = [th.get_text(strip=True) for th in table.find_all('th')]
    table_data = [headers]
    for row in table.find_all('tr'):
        cols = row.find_all('td')
        if cols:
            row_text = " ".join(col.get_text(strip=True) for col in cols)
            if any(total in row_text for total in ["Total HT", "Total TVA", "Total TTC"]):
                continue
            row_data = [col.get_text(strip=True) for col in cols]
            table_data.append(row_data)
    return table_data

def modify_pdf(html_content: str, output_pdf_path: str):
    """
    Modify PDF based on html_content.
    
    For F-files:
      - Draw header image F_kaddar.jpg on page 1, stretched to (8.47 + 2*0.82) inches wide × 2 inches high,
        with an x-offset of –0.82 inches so that it covers the full page width.
      - Add a subheader "FACTURE N°: <ref>/<year>" below the header.
      - Build the main table using the original eight columns.
      - Append a summary section (TOTAL HT/TVA/TTC) directly below the table.
    
    For BL-files:
      - Draw header image BL_kaddar.jpg on page 1 (same bleed: width = page width + 2×0.82 inches, height = 2 inches).
      - Add a subheader "BL N°: <ref>/<year>" below the header.
      - Build the main table using only three columns: "Désignation", "Unité de mesure", and "Quantité".
      - Do not append any summary section.
    
    In both cases, draw the footer image (footer_kaddar.jpg) on every page with bleed,
    stretched to (page width + 2×0.82 inches) wide × 1.5 inches, flush at the bottom.
    """
    # Determine if this is a BL file.
    is_bl = "bl_" in output_pdf_path.lower()
    
    # Parse reference and year from output filename.
    base = os.path.basename(output_pdf_path)
    name_part = os.path.splitext(base)[0]
    parts = name_part.split("_")
    try:
        mm_year = parts[0]
        year = mm_year.split("-")[1]
    except IndexError:
        year = "YYYY"
    ref_str = parts[2]
    ref_num = "".join(filter(str.isdigit, ref_str))
    
    # --------------------------
    # Determine Database Directory
    # --------------------------
    # Check for the OneDrive path first, then fall back to local.
    onedrive_db_dir = Path(os.path.expanduser(os.path.join("~", "OneDrive", "Desktop", "kaddar tarvaux beta version", "Database")))
    local_db_dir = Path(os.path.expanduser(os.path.join("~", "kaddar tarvaux beta version", "Database")))
    if onedrive_db_dir.exists():
        database_dir = onedrive_db_dir
    else:
        database_dir = local_db_dir

    pack_dir = database_dir / "pack_kaddar"
    
    # For full bleed, we set document margins to 0.
    doc = SimpleDocTemplate(
         output_pdf_path,
         pagesize=letter,
         leftMargin=0.0,
         rightMargin=0.0,
         topMargin=0.0,
         bottomMargin=0.0
    )
    
    elements = []
    styles = getSampleStyleSheet()
    
    # Define base cell style.
    cell_style = ParagraphStyle(
         'CellStyle',
         fontName='Helvetica',
         fontSize=9,
         leading=10,
         alignment=1,
         wordWrap=True
    )
    
    # Define header & footer bleed extension: extra 0.82 inches on each side.
    extra = 0.82 * inch
    
    # Header height = 2 inches, Footer height = 1.5 inches.
    header_height = 2 * inch
    footer_height = 1.5 * inch
    
    # Insert spacer at beginning so that main content starts below the header.
    elements.append(Spacer(1, header_height))
    
    # Add subheader with reference (appears on first page below header).
    if is_bl:
        subheader_text = f"<b>BL N°: {ref_num}/{year}</b>"
    else:
        subheader_text = f"<b>FACTURE N°: {ref_num}/{year}</b>"
    subheader_style = ParagraphStyle(
         'Subheader',
         parent=styles['Normal'],
         fontName='Helvetica-Bold',
         fontSize=12,
         alignment=1,
         spaceAfter=0.2 * inch
    )
    elements.append(Paragraph(subheader_text, subheader_style))
    
    # Main content paragraphs.
    extracted = extract_data(html_content)
    elements += [
         Paragraph(f"<b>ACHETEUR PUBLIC:</b> {extracted['acheteur']}", styles['Normal']),
         Paragraph(f"<b>OBJET:</b> {extracted['object']}", styles['Normal']),
         Paragraph(f"<b>LIEU D'EXÉCUTION:</b> {extracted['location']}", styles['Normal']),
         Spacer(1, 0.4 * inch)
    ]
    
    # Process main table data.
    table_data = extract_table_data(html_content)
    if table_data:
         if is_bl:
              # For BL files, keep only three columns: "Désignation", "Unité de mesure", "Quantité"
              allowed = {"désignation", "unité de mesure", "quantité"}
              header = table_data[0]
              indices = [i for i, h in enumerate(header) if h.strip().lower() in allowed]
              if not indices:
                  indices = list(range(3))
              new_table_data = []
              for row in table_data:
                  new_row = [row[i] for i in indices if i < len(row)]
                  new_table_data.append(new_row)
              table_data = new_table_data
              num_cols = len(table_data[0])
              col_widths = [doc.pagesize[0] / num_cols] * num_cols  # full page width
         else:
              # For F files, use original eight-column widths.
              col_widths = [0.5 * inch, 3.0 * inch, 0.5 * inch, 0.5 * inch,
                            1.5 * inch, 0.5 * inch, 1.2 * inch, 1.5 * inch]
         
         # Format table cells.
         formatted_data = []
         for row in table_data:
              formatted_row = [Paragraph(cell.replace('\n', '<br/>'), cell_style) for cell in row]
              formatted_data.append(formatted_row)
         
         main_table = Table(formatted_data, colWidths=col_widths, repeatRows=1)
         main_table.hAlign = 'CENTER'
         main_table.setStyle(TableStyle([
             ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#506F9B')),
             ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
             ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
             ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
             ('FONTSIZE', (0, 0), (-1, -1), 9),
             ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
             ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#F5F5F5')),
             ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
         ]))
         elements.append(main_table)
         
         # For F files, append the summary section (for BL files, no summary).
         if not is_bl:
              try:
                  total_ht_val = float(extracted.get("total_ht", "0"))
              except:
                  total_ht_val = 0.0
              try:
                  total_tva_val = float(extracted.get("total_tva", "0"))
              except:
                  total_tva_val = 0.0
              try:
                  total_ttc_val = float(extracted.get("total_ttc", "0"))
              except:
                  total_ttc_val = 0.0
              total_ht_str = f"{total_ht_val:.2f}"
              total_tva_str = f"{total_tva_val:.2f}"
              total_ttc_str = f"{total_ttc_val:.2f}"
              total_ttc_words = num2words(total_ttc_val, lang='fr').upper()
              left_text = f"<b>Arrêté la présente facture à la somme de {total_ttc_words} DIRHAMS</b>"
              summary_style = ParagraphStyle(
                  'SummaryStyle',
                  parent=cell_style,
                  fontName='Helvetica-Bold',
                  fontSize=12,
                  leading=14,
                  textColor=colors.whitesmoke,
                  alignment=1,
                  spaceAfter=4
              )
              summary_data = [
                  [Paragraph(left_text, summary_style), '', 
                   Paragraph("<b>TOTAL HT</b>", summary_style), '', 
                   Paragraph(f"<b>{total_ht_str}</b>", summary_style), ''],
                  ['', '', 
                   Paragraph("<b>TOTAL TVA</b>", summary_style), '', 
                   Paragraph(f"<b>{total_tva_str}</b>", summary_style), ''],
                  ['', '', 
                   Paragraph("<b>TOTAL TTC</b>", summary_style), '', 
                   Paragraph(f"<b>{total_ttc_str}</b>", summary_style), '']
              ]
              # Link summary with table (no spacer in between)
              total_summary_width = 9.2 * inch  # as originally designed
              left_width = 4.0 * inch
              right_width = total_summary_width - left_width
              summary_col_widths = [left_width/2, left_width/2] + [right_width/4]*4
              summary_table = Table(summary_data, colWidths=summary_col_widths)
              summary_table.hAlign = 'CENTER'
              summary_table.setStyle(TableStyle([
                  ('SPAN', (0,0), (1,2)),
                  ('SPAN', (2,0), (3,0)),
                  ('SPAN', (4,0), (5,0)),
                  ('SPAN', (2,1), (3,1)),
                  ('SPAN', (4,1), (5,1)),
                  ('SPAN', (2,2), (3,2)),
                  ('SPAN', (4,2), (5,2)),
                  ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#506F9B')),
                  ('TEXTCOLOR', (0,0), (-1,-1), colors.whitesmoke),
                  ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                  ('BOTTOMPADDING', (0,0), (-1,-1), 6),
                  ('TOPPADDING', (0,0), (-1,-1), 6),
                  ('BOX', (0,0), (-1,-1), 1, colors.black),
                  ('GRID', (0,0), (-1,-1), 0.5, colors.black),
              ]))
              elements.append(summary_table)
    
    # Header and Footer drawing callback.
    def draw_header_footer(canvas, doc):
         page_num = canvas.getPageNumber()
         # Get the full page width.
         page_width = letter[0]
         # Draw footer on every page with bleed.
         footer_img_path = str(pack_dir / "footer_kaddar.jpg")
         footer_width = page_width + 2 * extra
         x_footer = -extra
         if os.path.exists(footer_img_path):
              canvas.drawImage(footer_img_path, x_footer, 0, width=footer_width, height=footer_height, preserveAspectRatio=True)
         # Draw header only on the first page.
         if page_num == 1:
              if is_bl:
                   header_img = str(pack_dir / "BL_kaddar.jpg")
              else:
                   header_img = str(pack_dir / "F_kaddar.jpg")
              header_width = page_width + 2 * extra
              x_header = -extra
              if os.path.exists(header_img):
                   # Draw header flush at the top of the page.
                   canvas.drawImage(header_img, x_header, letter[1] - header_height, width=header_width, height=header_height, preserveAspectRatio=True)
    
    doc.build(elements, onFirstPage=draw_header_footer, onLaterPages=draw_header_footer)
