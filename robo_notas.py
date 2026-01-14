# ---------------------------------------------------------
# Projeto: Automação de Entrada de Matéria Prima - VERSÃO TURBO
# Autor: Gabriel Sonvezzo
# Descrição: Processamento em lote com salvamento único e busca em dicionário
# ---------------------------------------------------------

import os
import xmltodict
import re
from datetime import datetime, timedelta
import openpyxl
from copy import copy
from openpyxl.styles import Alignment
import shutil

# --- CONFIGURAÇÕES ---
pasta_xml = Notas_Pendentes
pasta_processadas = Notas_Processadas
arquivo_excel = modelo_analise.xlsx
COLUNA_LOTE = "W" 

MAPA = {
    "NF": "C", "ORDEM": "D", "FORNEC": "E", "DATA_NF": "F",
    "PESO": "G", "PROD": "H", "QUALIDADE": "I", "ESP": "J", "LARG": "K",
    "VALOR_TOTAL": "L", "ICMS": "M", "IPI": "N",
    "VENC": "T", "FATURA": "U", "VALOR_NF": "V"
}

def calcular_dia_util(data_emissao):
    vencimento = data_emissao + timedelta(days=7)
    if vencimento.weekday() == 5: return vencimento + timedelta(days=2)
    elif vencimento.weekday() == 6: return vencimento + timedelta(days=1)
    return vencimento

def aplicar_estilo_personalizado(celula_origem, celula_destino, col):
    if celula_origem.has_style:
        celula_destino.font = copy(celula_origem.font)
        celula_destino.border = copy(celula_origem.border)
        celula_destino.fill = copy(celula_origem.fill)
        
        if col == "J":
            celula_destino.alignment = Alignment(horizontal='right', vertical='bottom')
            celula_destino.number_format = '_-* #,##0.00_-;_-* #,##0.00_-;_-* "-"??_-;_-@_-'
        elif col == "T":
            celula_destino.alignment = Alignment(vertical='bottom')
            celula_destino.number_format = 'dd/mm/yyyy'
        elif col == "F":
            celula_destino.number_format = 'dd/mm/yy'
            celula_destino.alignment = Alignment(horizontal='right', vertical='bottom')
        else:
            celula_destino.alignment = copy(celula_origem.alignment)
            celula_destino.number_format = copy(celula_origem.number_format)

def processar_arquivos(callback=None):
    if not os.path.exists(arquivo_excel):
        if callback: callback("⚠️ Planilha não encontrada!"); return

    try:
        # Carregamento inicial
        if callback: callback("⏳ Abrindo planilha... aguarde.")
        wb = openpyxl.load_workbook(arquivo_excel, data_only=False)
        sheet = next((wb[n] for n in wb.sheetnames if "Arcelor Usina" in n), None)
    except Exception as e:
        if callback: callback(f"⚠️ Erro ao abrir Excel: {e}"); return

    arquivos_xml = [f for f in os.listdir(pasta_xml) if f.lower().endswith(".xml")]
    if not arquivos_xml:
        if callback: callback("ℹ️ Nenhum XML na pasta pendente."); return

    # --- TURBO 1: MAPA DE LOTES (Busca instantânea) ---
    mapa_lotes_excel = {}
    for cell in sheet[COLUNA_LOTE]:
        if cell.value:
            val = str(cell.value).strip().upper()
            mapa_lotes_excel[val] = cell.row

    arquivos_para_mover = []
    houve_atualizacao_geral = False

    for nome_arquivo in arquivos_xml:
        caminho_xml = os.path.join(pasta_xml, nome_arquivo)
        houve_gravacao_nesta_nf = False 
        
        try:
            with open(caminho_xml, "rb") as f:
                xml_data = xmltodict.parse(f.read())
            
            inf = xml_data['nfeProc']['NFe']['infNFe']
            num_nf = inf['ide']['nNF']
            detalhes = inf['det']
            if not isinstance(detalhes, list): detalhes = [detalhes]

            # --- BUSCA O LOTE NA TAG NUMERAÇÃO (nVol) ---
            lote_xml = ""
            try:
                transp = inf.get('transp', {})
                vol = transp.get('vol', {})
                lote_xml = str(vol[0].get('nVol', '') if isinstance(vol, list) else vol.get('nVol', '')).upper().strip()
            except:
                lote_xml = ""

            # --- TURBO 2: BUSCA NO DICIONÁRIO ---
            linha_encontrada = mapa_lotes_excel.get(lote_xml)

            if linha_encontrada:
                for item in detalhes:
                    xProd = str(item['prod']['xProd']).upper()
                    dt_emi = datetime.strptime(inf['ide']['dhEmi'][:10], "%Y-%m-%d")
                    totais = inf['total']['ICMSTot']
                    
                    # Lógica de Sigla e Qualidade
                    if any(x in xProd for x in ["GALV", "ZC", "NBR", "7008"]):
                        sigla, qualidade = "BZ", (re.search(r'(NBR\s?\d+[^,]*|ZC\s?\d+)', xProd).group(1) if re.search(r'(NBR\s?\d+[^,]*|ZC\s?\d+)', xProd) else "NBR 7008 ZC")
                    elif "FRIO" in xProd or "BF" in xProd:
                        sigla, qualidade = "BF", (re.search(r'(SAE\s?J?\d+\s?\d*)', xProd).group(1) if re.search(r'(SAE\s?J?\d+\s?\d*)', xProd) else "SAE 1008")
                    else:
                        sigla, qualidade = "BQ", (re.search(r'(SAE\s?J?\d+\s?\d*|A36)', xProd).group(1) if re.search(r'(SAE\s?J?\d+\s?\d*|A36)', xProd) else "SAE")
                    
                    esp_valor = float(re.search(r'(\d+\.\d{3})', xProd).group(1)) if re.search(r'(\d+\.\d{3})', xProd) else 0.0
                    larg = int(re.search(r'X\s?(\d+)\.', xProd).group(1)) if re.search(r'X\s?(\d+)\.', xProd) else ""

                    dados = {
                        "NF": int(num_nf),
                        "ORDEM": re.search(r'(\d{10})', xProd).group(1) if re.search(r'(\d{10})', xProd) else "",
                        "FORNEC": 36003, "DATA_NF": dt_emi, "PESO": int(float(item['prod']['qCom']) * 1000),
                        "PROD": sigla, "QUALIDADE": qualidade, "ESP": esp_valor, "LARG": larg,
                        "VALOR_TOTAL": float(totais['vNF']), "ICMS": float(totais['vICMS']),
                        "IPI": float(totais['vIPI']), "VENC": calcular_dia_util(dt_emi),
                        "FATURA": f"00{num_nf}01", "VALOR_NF": float(totais['vNF'])
                    }

                    for ref, col in MAPA.items():
                        celula_dest = sheet[f"{col}{linha_encontrada}"]
                        celula_dest.value = dados.get(ref)
                        if linha_encontrada > 1:
                            aplicar_estilo_personalizado(sheet[f"{col}{linha_encontrada-1}"], celula_dest, col)

                if callback: callback(f"🚀 NF {num_nf} processada (Linha {linha_encontrada})")
                houve_gravacao_nesta_nf = True
                houve_atualizacao_geral = True
                arquivos_para_mover.append(nome_arquivo)
            else:
                if callback: callback(f"❌ NF {num_nf}: Lote {lote_xml} não localizado.")

        except Exception as e:
            if callback: callback(f"⚠️ Erro no XML {nome_arquivo}: {e}")

    # --- TURBO 3: ÚNICO SALVAMENTO FINAL ---
    if houve_atualizacao_geral:
        if callback: callback("💾 Gravando todas as notas no Excel... aguarde.")
        wb.save(arquivo_excel)
        
        # Move os arquivos apenas após o sucesso do salvamento
        for nome in arquivos_para_mover:
            shutil.move(os.path.join(pasta_xml, nome), os.path.join(pasta_processadas, nome))
            
        if callback: callback("🏁 PROCESSO CONCLUÍDO!")
    else:
        if callback: callback("ℹ️ Fim. Nenhuma alteração feita.")

if __name__ == "__main__": processar_arquivos()