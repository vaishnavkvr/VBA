def generate_documentation(vba_code):
    lines = vba_code.split("\n")
    doc_lines = ["# VBA Macro Documentation", ""]
    
    for line in lines:
        if "Sub " in line or "Function " in line:
            doc_lines.append(f"## {line.strip()}")
        elif line.strip().startswith("'"):
            doc_lines.append(line.strip().replace("'", ""))
    
    return "\n".join(doc_lines)

if __name__ == "__main__":
    with open("extracted_vba_code.txt", "r") as f:
        vba_code = f.read()

    documentation = generate_documentation(vba_code)
    with open("documentation.md", "w") as f:
        f.write(documentation)
