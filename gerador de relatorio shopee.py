import tkinter as tk
from tkinter import filedialog, messagebox
import re
import pandas as pd
import os

def extract_sku(text):
    """Extrai o SKU principal do texto e retorna apenas a numeração."""
    pattern = r"SKU principal:\s*(\d+)"
    matches = re.findall(pattern, text)
    return matches

def process_text():
    """Processa o texto inserido na interface e gera a planilha."""
    text = text_input.get("1.0", tk.END)
    skus = extract_sku(text)
    
    if not skus:
        messagebox.showinfo("Aviso", "Nenhum SKU principal encontrado no texto.")
        return
    
    output_dir = r"C:\\Users\\USER\\OneDrive\\Desktop\\Platinhas Retorno BOT\\planilhas"
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, "RELATORIO SHOPEE.xlsx")
    
    df = pd.DataFrame(skus, columns=["SKU Principal"])
    df.to_excel(output_path, index=False)
    
    messagebox.showinfo("Sucesso", f"Planilha gerada com sucesso em:\n{output_path}")

root = tk.Tk()
root.title("Extrator de SKU")
root.geometry("500x300")

label = tk.Label(root, text="Cole o texto abaixo:")
label.pack(pady=5)

text_input = tk.Text(root, height=10, width=50)
text_input.pack(pady=5)

process_button = tk.Button(root, text="Gerar Planilha", command=process_text)
process_button.pack(pady=10)

tk.mainloop()