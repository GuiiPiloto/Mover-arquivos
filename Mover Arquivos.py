import os
import shutil
import tkinter as tk
from tkinter import scrolledtext, messagebox
import unicodedata
import threading

def normalize_text(text):
    return ''.join(c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn').upper()

def find_prestados_in_company(company_path):
    for root, dirs, files in os.walk(company_path):
        if "PRESTADOS" in dirs and "2025" in root and "03-2025" in root:
            return os.path.join(root, "PRESTADOS")
    return None

def find_company_in_z(company_name, dest_root):
    company_norm = normalize_text(company_name)
    company_words = set(company_norm.split())
    
    best_match = None
    best_score = 0
    
    for folder in os.listdir(dest_root):
        folder_norm = normalize_text(folder)
        folder_words = set(folder_norm.split())
        common_words = company_words.intersection(folder_words)
        score = len(common_words)
        
        if score > 0 and (len(folder_words) <= len(company_words) or "LTDA" in company_words):
            if score > best_score:
                best_score = score
                best_match = folder
    
    return os.path.join(dest_root, best_match) if best_match else None

def scan_paths():
    source_root = r"C:\Users\FISCAL01\Downloads"
    dest_root = r"Z:\\"
    results = []
    
    for company in os.listdir(source_root):
        source_path = os.path.join(source_root, company)
        if not os.path.isdir(source_path):
            continue
        
        company_dest_path = find_company_in_z(company, dest_root)
        if not company_dest_path:
            results.append({"company": company, "path": f"Erro: Empresa não encontrada em Z:\\", "approved": False})
            continue
        
        prestados_path = find_prestados_in_company(company_dest_path)
        if not prestados_path:
            results.append({"company": company, "path": f"Erro: PRESTADOS não encontrado em {company_dest_path}", "approved": False})
        else:
            results.append({"company": company, "path": prestados_path, "approved": True})
    
    return results

def move_files(results, output_text):
    allowed_extensions = (".pdf", ".xml", ".txt", ".xlsx", ".xls")
    
    for item in results:
        if not item["approved"]:
            output_text.insert(tk.END, f"Pulando {item['company']} - não aprovado ou erro.\n")
            continue
        
        source_path = os.path.join(r"C:\Users\FISCAL01\Downloads", item["company"])
        dest_path = item["path"]
        
        for filename in os.listdir(source_path):
            if filename.lower().endswith(allowed_extensions):
                source_file = os.path.join(source_path, filename)
                dest_file = os.path.join(dest_path, filename)
                
                if os.path.exists(dest_file):
                    output_text.insert(tk.END, f"Pulando {filename} em '{item['company']}' - já existe.\n")
                    continue
                
                try:
                    shutil.move(source_file, dest_file)
                    output_text.insert(tk.END, f"Movido: {filename} de '{item['company']}' para {dest_path}\n")
                except Exception as e:
                    output_text.insert(tk.END, f"Erro ao mover {filename} de '{item['company']}': {e}\n")
    output_text.insert(tk.END, "Concluído!\n")

def populate_gui(window, frame, output_text, check_vars):
    results = scan_paths()
    if not results:
        messagebox.showerror("Erro", "Nenhuma pasta de empresa encontrada em Downloads!", parent=window)
        window.destroy()
        return
    
    for widget in frame.winfo_children():
        widget.destroy()
    
    tk.Label(frame, text="Empresa", width=30, anchor="w", bg="#1C2526", fg="#00FFFF", font=("Helvetica", 10, "bold")).grid(row=0, column=0, padx=2, pady=2)
    tk.Label(frame, text="Caminho PRESTADOS", width=100, anchor="w", bg="#1C2526", fg="#00FFFF", font=("Helvetica", 10, "bold")).grid(row=0, column=1, padx=2, pady=2)
    tk.Label(frame, text="Aprovar", width=10, bg="#1C2526", fg="#00FFFF", font=("Helvetica", 10, "bold")).grid(row=0, column=2, padx=2, pady=2)
    
    for i, item in enumerate(results):
        tk.Label(frame, text=item["company"][:29], width=30, anchor="w", bg="#2E3839", fg="white", font=("Helvetica", 9)).grid(row=i+1, column=0, padx=2, pady=2)
        path_text = tk.Text(frame, height=2, width=100, wrap="word", bg="#2E3839", fg="white", font=("Helvetica", 9), borderwidth=0)
        path_text.insert("1.0", item["path"])
        path_text.config(state="disabled")
        path_text.grid(row=i+1, column=1, padx=2, pady=2)
        var = tk.BooleanVar(value=item["approved"])
        tk.Checkbutton(frame, variable=var, bg="#2E3839", fg="#00FFFF", selectcolor="#1C2526").grid(row=i+1, column=2, padx=2, pady=2)
        check_vars.append(var)
    
    def approve_all():
        for var in check_vars:
            var.set(True)
    
    tk.Button(window, text="Sim para Todos", command=approve_all, bg="#00FFFF", fg="black", font=("Helvetica", 10, "bold"), relief="flat", bd=0, padx=10, pady=5).pack(pady=5)
    
    def start_move():
        for i, var in enumerate(check_vars):
            results[i]["approved"] = var.get()
        move_files(results, output_text)
    
    tk.Button(window, text="Mover Arquivos", command=start_move, bg="#00FFFF", fg="black", font=("Helvetica", 10, "bold"), relief="flat", bd=0, padx=10, pady=5).pack(pady=5)

def create_gui():
    window = tk.Tk()
    window.title("Movimentação de Arquivos")
    window.geometry("1205x600")
    window.configure(bg="#1C2526")
    
    frame = tk.Frame(window, bg="#1C2526")
    frame.pack(fill="both", expand=True, padx=10, pady=10)
    
    tk.Label(frame, text="CAÇANDO SUAS PASTAS...", font=("Helvetica", 16, "bold"), bg="#1C2526", fg="#00FFFF").pack(pady=20)
    
    output_text = scrolledtext.ScrolledText(window, height=10, bg="#2E3839", fg="white", font=("Helvetica", 9), insertbackground="white")
    output_text.pack(fill="both", expand=True, padx=10, pady=5)
    
    check_vars = []
    
    threading.Thread(target=populate_gui, args=(window, frame, output_text, check_vars), daemon=True).start()
    
    window.mainloop()

def main():
    create_gui()

if __name__ == "__main__":
    main()