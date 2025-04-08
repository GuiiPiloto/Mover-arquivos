import os
import shutil
import tkinter as tk
from tkinter import scrolledtext, messagebox, ttk
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
            results.append({"company": company, "path": f"Erro: Empresa não encontrada em Z:\\", "approved": False, "target": "PRESTADOS"})
            continue
        
        prestados_path = find_prestados_in_company(company_dest_path)
        if not prestados_path:
            results.append({"company": company, "path": f"Erro: PRESTADOS não encontrado em {company_dest_path}", "approved": False, "target": "PRESTADOS"})
            continue
        
        # Deriva o caminho do AR a partir do caminho do PRESTADOS
        ar_path = prestados_path.replace("PRESTADOS", "AR")
        
        results.append({
            "company": company,
            "prestados_path": prestados_path,
            "ar_path": ar_path,
            "path": prestados_path,  # Caminho padrão é PRESTADOS
            "approved": True,
            "target": "PRESTADOS"
        })
    
    return results

def move_files(results, output_text):
    allowed_extensions_prestados = (".pdf", ".xml", ".txt", ".xlsx", ".xls")
    
    for item in results:
        if not item["approved"]:
            output_text.insert(tk.END, f"Pulando {item['company']} - não aprovado ou erro.\n")
            continue
        
        source_path = os.path.join(r"C:\Users\FISCAL01\Downloads", item["company"])
        dest_path = item["path"]
        
        if "Erro" in dest_path:
            output_text.insert(tk.END, f"Pulando {item['company']} - caminho inválido: {dest_path}\n")
            continue
        
        if item["target"] == "PRESTADOS":
            # Lógica de PRESTADOS (sem alterações)
            for filename in os.listdir(source_path):
                if filename.lower().endswith(allowed_extensions_prestados):
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
        else:  # item["target"] == "AR"
            # Lógica de AR: move as subpastas Entrada e Saída
            entry_folders = []
            for folder in os.listdir(source_path):
                folder_path = os.path.join(source_path, folder)
                if os.path.isdir(folder_path):
                    normalized_folder = normalize_text(folder)
                    if normalized_folder in ("ENTRADA", "SAIDA"):
                        entry_folders.append(folder)
            
            if not entry_folders:
                output_text.insert(tk.END, f"Nenhuma subpasta Entrada ou Saída encontrada em '{item['company']}'.\n")
                continue
            
            output_text.insert(tk.END, f"Subpastas encontradas em '{item['company']}': {', '.join(entry_folders)}\n")
            
            # Remove subpastas Entrada e Saída existentes no destino
            for folder in entry_folders:
                dest_folder_path = os.path.join(dest_path, folder)
                if os.path.exists(dest_folder_path):
                    try:
                        shutil.rmtree(dest_folder_path)
                        output_text.insert(tk.END, f"Subpasta '{folder}' já existia em {dest_path} e foi substituída.\n")
                    except Exception as e:
                        output_text.insert(tk.END, f"Erro ao remover subpasta '{folder}' em {dest_path}: {e}\n")
                        continue
            
            # Move as subpastas Entrada e Saída
            for folder in entry_folders:
                source_folder_path = os.path.join(source_path, folder)
                dest_folder_path = os.path.join(dest_path, folder)
                try:
                    shutil.move(source_folder_path, dest_folder_path)
                    output_text.insert(tk.END, f"Movida subpasta '{folder}' de '{item['company']}' para {dest_folder_path}\n")
                except Exception as e:
                    output_text.insert(tk.END, f"Erro ao mover subpasta '{folder}' de '{item['company']}': {e}\n")
    
    output_text.insert(tk.END, "Concluído!\n")

def populate_gui(window, frame, output_text, check_vars, target_vars):
    results = scan_paths()
    if not results:
        messagebox.showerror("❌ Erro", "Nenhuma pasta de empresa encontrada em Downloads!")
        window.destroy()
        return

    for widget in frame.winfo_children():
        widget.destroy()

    header_style = {"font": ("Segoe UI", 10, "bold")}
    tk.Label(frame, text="🏢 Empresa", width=30, anchor="w", **header_style).grid(row=0, column=0)
    tk.Label(frame, text="🎯 Destino", width=10, **header_style).grid(row=0, column=1)
    tk.Label(frame, text="📁 Caminho", width=90, anchor="w", **header_style).grid(row=0, column=2)
    tk.Label(frame, text="✅ Aprovar", width=10, **header_style).grid(row=0, column=3)

    def update_path(row, target_var, path_label, result):
        target = target_var.get()
        result["target"] = target
        result["path"] = result["prestados_path"] if target == "PRESTADOS" else result["ar_path"]
        path_label.config(text=result["path"][:89] if result["path"] else "❌ Caminho não encontrado")

    for i, item in enumerate(results):
        tk.Label(frame, text=item["company"][:29], width=30, anchor="w").grid(row=i+1, column=0)
        target_var = tk.StringVar(value=item["target"])
        combo = ttk.Combobox(frame, textvariable=target_var, values=["PRESTADOS", "AR"], width=10, state="readonly")
        combo.grid(row=i+1, column=1)
        target_vars.append(target_var)
        path_label = tk.Label(frame, text=item["path"][:89] if item["path"] else "❌ Caminho não encontrado", width=90, anchor="w")
        path_label.grid(row=i+1, column=2)
        target_var.trace("w", lambda *args, r=i, tv=target_var, pt=path_label, res=item: update_path(r, tv, pt, res))
        var = tk.BooleanVar(value=item["approved"])
        tk.Checkbutton(frame, variable=var).grid(row=i+1, column=3)
        check_vars.append(var)

    def approve_all():
        for var in check_vars:
            var.set(True)

    tk.Button(window, text="✅ Aprovar Todos", command=approve_all).pack(pady=5)

    def start_move():
        for i, var in enumerate(check_vars):
            results[i]["approved"] = var.get()
        move_files(results, output_text)

    tk.Button(window, text="📦 Mover Arquivos", command=start_move).pack(pady=5)

def create_gui():
    window = tk.Tk()
    window.title("📁 Movimentação de Arquivos")
    window.geometry("1200x650")

    frame = tk.Frame(window)
    frame.pack(fill="both", expand=True, padx=10, pady=10)

    tk.Label(frame, text="🔍 CAÇANDO SUAS PASTAS...", font=("Segoe UI", 14)).pack(pady=20)

    output_text = scrolledtext.ScrolledText(window, height=10, font=("Consolas", 10))
    output_text.pack(fill="both", expand=True, padx=10, pady=5)

    check_vars = []
    target_vars = []

    threading.Thread(target=populate_gui, args=(window, frame, output_text, check_vars, target_vars), daemon=True).start()

    window.mainloop()

def main():
    create_gui()

if __name__ == "__main__":
    main()