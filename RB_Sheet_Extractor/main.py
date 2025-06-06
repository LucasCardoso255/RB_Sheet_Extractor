import os
import docx
import re
import openpyxl

def extract_word_data(file_path):
    product_name = "N/A"
    client = "N/A"
    production_per_hour = "N/A"
    channel_break = "Não"
    deburring = "Não"
    sandblasting = "Não"
    sanding = "Não"
    packaging_inspection = "Não"

    try:
        doc = docx.Document(file_path)

        full_text = ""
        for paragraph in doc.paragraphs:
            full_text += paragraph.text + " "

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    full_text += cell.text + " "

        full_text = full_text.replace("\xa0", " ").replace("\t", " ").replace("\n", " ")

        name_match = re.search(r"Nome:\s*(.*?)\s*Código:", full_text, re.DOTALL)
        client_match = re.search(r"Cliente:\s*([^\n\r]+?)\s", full_text)
        production_match = re.search(r"(\d+)\s*Peças", full_text, re.IGNORECASE)

        if name_match:
            product_name = name_match.group(1).strip()
        if client_match:
            client = client_match.group(1).strip()
        if production_match:
            production_per_hour = production_match.group(1).strip()

        if "QUEBRA DO CANAL" in full_text.upper():
            channel_break = "Sim"
        if "REBARBAÇÃO" in full_text.upper():
            deburring = "Sim"
        if "JATO DE GRANALHA" in full_text.upper():
            sandblasting = "Sim"
        if "LIXA" in full_text.upper():
            sanding = "Sim"
        if "INSPEÇÃO FINAL" in full_text.upper():
            packaging_inspection = "Sim"

        return {
            "product_name": product_name,
            "client": client,
            "production_per_hour": production_per_hour,
            "channel_break": channel_break,
            "deburring": deburring,
            "sandblasting": sandblasting,
            "sanding": sanding,
            "packaging_inspection": packaging_inspection,
        }

    except Exception as e:
        print(f"Error processing {file_path}: {e}")
        return None

def process_folders_and_create_excel(root_folder, output_excel_file):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(
        [
            "Nome do Produto",
            "Cliente",
            "Producao/Hora",
            "Quebra de Canal",
            "Rebarbacao",
            "Jateamento",
            "Lixa",
            "Inspecao e Embalagem",
        ]
    )

    for current_folder, subfolders, files in os.walk(root_folder):
        for file_name in files:
            if file_name.endswith(".docx") and not file_name.startswith("~$"):
                full_path = os.path.join(current_folder, file_name)
                print(f"Processing file: {full_path}")
                data = extract_word_data(full_path)
                if data:
                    sheet.append(
                        [
                            data["product_name"],
                            data["client"],
                            data["production_per_hour"],
                            data["channel_break"],
                            data["deburring"],
                            data["sandblasting"],
                            data["sanding"],
                            data["packaging_inspection"],
                        ]
                    )
                else:
                    print(f"File {full_path} skipped due to extraction error.")

    workbook.save(output_excel_file)
    print(f"Excel file generated: {output_excel_file}")

if __name__ == "__main__":
    root_folder = os.getcwd()  
    output_excel_file = "produtos_simplificado.xlsx"
    process_folders_and_create_excel(root_folder, output_excel_file)