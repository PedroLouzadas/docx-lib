import sys
from PyPDF2 import PdfMerger
import os

def merge_pdfs(input_paths, output_path):
    merger = PdfMerger()
    
    for pdf in input_paths:
        if not os.path.exists(pdf):
            print(f"Arquivo não encontrado: {pdf}")
            sys.exit(1)
        print(f"Adicionando: {pdf}")
        merger.append(pdf)
    
    print(f"Salvando PDF mesclado em: {output_path}")
    merger.write(output_path)
    merger.close()
    print("Mesclagem concluída com sucesso!")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Uso: python merge_pdfs.py <output_path> <input1.pdf> <input2.pdf> ...")
        sys.exit(1)

    output_file = sys.argv[1]
    input_files = sys.argv[2:]
    merge_pdfs(input_files, output_file)