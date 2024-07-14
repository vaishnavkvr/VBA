import win32com.client

def extract_vba_from_excel(file_path):
    excel = win32com.client.Dispatch("Excel.Application")
    workbook = excel.Workbooks.Open(file_path)
    vba_project = workbook.VBProject
    vba_code = []

    for component in vba_project.VBComponents:
        if component.Type == 1:  # Standard module
            vba_code.append(component.CodeModule.Lines(1, component.CodeModule.CountOfLines))

    workbook.Close(False)
    excel.Quit()
    return "\n".join(vba_code)

if __name__ == "__main__":
    file_path = "path_to_your_excel_file.xlsm"
    vba_code = extract_vba_from_excel(file_path)
    with open("extracted_vba_code.txt", "w") as f:
        f.write(vba_code)
