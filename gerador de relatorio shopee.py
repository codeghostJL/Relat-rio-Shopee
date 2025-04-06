import tkinter as tk
from tkinter import messagebox, filedialog
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

    output_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Planilha Excel", "*.xlsx")],
        title="Salvar planilha",
        initialfile="RELATORIO SHOPEE.xlsx"
    )

    if not output_path:
        return 

    df = pd.DataFrame(skus, columns=["SKU Principal"])
    df.to_excel(output_path, index=False)

    messagebox.showinfo("Sucesso", f"Planilha gerada com sucesso em:\n{output_path}")

def clear_text():
    """Limpa o conteúdo do campo de texto."""
    text_input.delete("1.0", tk.END)

root = tk.Tk()
root.title("Extrator de SKU")
root.geometry("520x370")

label = tk.Label(root, text="Cole o texto abaixo:")
label.pack(pady=5)

frame = tk.Frame(root)
frame.pack()

text_input = tk.Text(frame, height=10, width=60, wrap="word")
scrollbar = tk.Scrollbar(frame, command=text_input.yview)
text_input.configure(yscrollcommand=scrollbar.set)

text_input.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

button_frame = tk.Frame(root)
button_frame.pack(pady=10)

process_button = tk.Button(button_frame, text="Gerar Planilha", command=process_text)
process_button.pack(side=tk.LEFT, padx=10)

clear_button = tk.Button(button_frame, text="Excluir", command=clear_text)
clear_button.pack(side=tk.LEFT, padx=10)

footer = tk.Label(root, text="Desenvolvido por João Lucas", font=("Arial", 10), fg="gray")
footer.pack(pady=10)

root.mainloop()
