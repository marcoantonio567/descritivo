import openpyxl

def extract_text_from_excel(file_path, sheet_name, cell_range):

    try:
        # Carrega o arquivo Excel
        workbook = openpyxl.load_workbook(file_path,data_only=True)

        # Seleciona a aba especificada
        sheet = workbook[sheet_name]

        # Extrai o intervalo de células
        cells = sheet[cell_range]

        # Armazena os valores extraídos em uma lista única
        extracted_data = []
        for row in cells:
            extracted_data.extend([cell.value for cell in row])

        return extracted_data

    except Exception as e:
        print(f"Erro ao processar o arquivo: {e}")
        return []

# Exemplo de uso
if __name__ == "__main__":
    # Caminho para o arquivo Excel
    file_path = "estattis_automação.xlsx"

    # Nome da aba
    sheet_name = "QUADRO"  # Atualize para o nome da sua aba

    # Intervalo de células (exemplo: "A1:C10")
    cell_range = "C4:L9"

    # Extrai os dados
    extracted_data = extract_text_from_excel(file_path, sheet_name, cell_range)
    # Exibe os dados extraídos

    print("Dados extraídos:")
    print(extracted_data)