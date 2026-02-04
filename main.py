import pandas as pd
from datetime import datetime

class ArchiveSystem:
    def __init__(self):
        self.database = []

    def add_box(self, box_id, content, department):
        # Aqui pegamos a data de um jeito que não dá erro
        hoje = datetime.now().strftime("%d/%m/%Y")
        
        entry = {
            "ID_Caixa": box_id,
            "Conteudo": content,
            "Departamento": department,
            "Data_Entrada": hoje
        }
        self.database.append(entry)
        print(f"📦 Caixa {box_id} cadastrada com sucesso!")

    def export_to_excel(self):
        if not self.database:
            print("❌ Erro: Nenhuma caixa para exportar.")
            return
            
        df = pd.DataFrame(self.database)
        df.to_excel("relatorio_estoque.xlsx", index=False)
        print("📊 Relatório Excel gerado com sucesso!")

# --- EXECUÇÃO DO PROGRAMA ---
try:
    system = ArchiveSystem()
    
    # Adicionando os dados (Exemplos da sua rotina na Seedoc/Cargill)
    system.add_box("CX-2026-001", "Notas Fiscais Cargill", "Financeiro")
    system.add_box("CX-2026-002", "Contratos Seedoc", "Administrativo")
    system.add_box("CX-2026-003", "Documentos RH", "Recursos Humanos")
    
    system.export_to_excel()

except Exception as e:
    print(f"⚠️ Opa, deu um erro aqui: {e}")