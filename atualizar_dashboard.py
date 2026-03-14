#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os
import pandas as pd
from datetime import datetime
from pathlib import Path
import subprocess
import sys

# --- CONFIGURAÇÕES VISUAIS ---
AZUL_MC = "#2d3175"
VERDE = "#22c55e"
VERMELHO = "#ef4444"

def criar_html_individual(m, periodo):
    """Gera o layout do relatório individual com o botão de impressão"""
    eventos = "".join([
        f"<tr><td>{e['data']}</td><td>{e['desc']}</td><td>{e['pecas']}</td><td>{e['tempo']} min</td></tr>"
        for e in m['events'] if str(e['desc']).upper() != 'OK'
    ])
    
    return f"""
    <!DOCTYPE html>
    <html lang="pt-BR">
    <head>
        <meta charset="UTF-8">
        <title>Relatório - {m['name']}</title>
        <style>
            body {{ font-family: sans-serif; margin: 40px; color: #333; background: #f4f7f6; }}
            .no-print {{ text-align: right; margin-bottom: 20px; }}
            .btn-print {{ background: {AZUL_MC}; color: white; border: none; padding: 12px 20px; border-radius: 5px; cursor: pointer; font-weight: bold; }}
            .header {{ background: white; padding: 20px; border-bottom: 4px solid {AZUL_MC}; margin-bottom: 20px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }}
            .kpis {{ display: flex; gap: 20px; margin: 20px 0; }}
            .card {{ background: white; padding: 15px; border-radius: 8px; flex: 1; text-align: center; border: 1px solid #ddd; }}
            table {{ width: 100%; border-collapse: collapse; background: white; margin-top: 20px; }}
            th {{ background: {AZUL_MC}; color: white; padding: 12px; text-align: left; }}
            td {{ padding: 10px; border-bottom: 1px solid #eee; font-size: 13px; }}
            @media print {{ .no-print {{ display: none; }} body {{ margin: 0; background: white; }} }}
        </style>
    </head>
    <body>
        <div class="no-print"><button class="btn-print" onclick="window.print()">🖨️ IMPRIMIR / SALVAR PDF</button></div>
        <div class="header">
            <h2 style="margin:0; color:{AZUL_MC}">{m['name']}</h2>
            <p style="margin:5px 0 0 0; color:#666;">Competência: {periodo} | MultCaixa</p>
        </div>
        <div class="kpis">
            <div class="card"><b>Disponibilidade</b><br><span style="color:{VERDE if m['av']>90 else VERMELHO}; font-size:20px;">{m['av']:.1f}%</span></div>
            <div class="card"><b>Ocorrências</b><br><span style="font-size:20px;">{m['count']}</span></div>
            <div class="card"><b>Tempo Parado</b><br><span style="font-size:20px;">{m['stop']} min</span></div>
        </div>
        <table>
            <thead><tr><th>Data</th><th>Serviço/Problema</th><th>Peças Trocadas</th><th>Tempo</th></tr></thead>
            <tbody>{eventos if eventos else '<tr><td colspan="4" style="text-align:center;">Sem ocorrências.</td></tr>'}</tbody>
        </table>
    </body>
    </html>"""

def run():
    # Caminho base do script
    base_path = Path(os.path.abspath(os.path.dirname(sys.argv[0])))
    os.chdir(base_path)
    
    # 1. Localizar Planilha
    planilhas = list(base_path.glob('*.xlsx'))
    if not planilhas:
        print("❌ Nenhuma planilha .xlsx encontrada!")
        return
    
    xlsx_path = planilhas[0]
    xls = pd.ExcelFile(xlsx_path)
    
    # Criar pasta para relatórios se não existir
    pasta_relatorios = base_path / "relatorios"
    pasta_relatorios.mkdir(exist_ok=True)
    
    all_machines = []

    for sheet in xls.sheet_names:
        if "Resina" not in sheet and "Gel Coat" not in sheet: continue
        
        df = pd.read_excel(xls, sheet_name=sheet, skiprows=6)
        df.columns = [str(c).strip() for c in df.columns]
        if 'DATA' not in df.columns: continue
        df = df.dropna(subset=['DATA'])

        m_stop, m_count, m_events = 0, 0, []
        for _, row in df.iterrows():
            desc = str(row.get('Problemas e Serviços', 'OK')).strip()
            tempo = row.get('TEMPO DE PARADA (min)', 0)
            try: t_val = int(tempo) if str(tempo).isdigit() else 0
            except: t_val = 0
            
            m_events.append({'data': str(row['DATA'])[:10], 'desc': desc, 'pecas': str(row.get('PEÇAS TROCADAS', '-')), 'tempo': t_val})
            if desc.upper() != 'OK' and desc != '' and desc.upper() != 'NAN':
                m_stop += t_val
                m_count += 1

        dispo = ((13200 - m_stop) / 13200) * 100
        m_dados = {'name': sheet, 'av': dispo, 'stop': m_stop, 'count': m_count, 'events': m_events}
        all_machines.append(m_dados)

        # Salvar o HTML Individual
        nome_html = f"relatorio_{sheet.replace(' ','_')}.html"
        with open(pasta_relatorios / nome_html, "w", encoding="utf-8") as f:
            f.write(criar_html_individual(m_dados, xlsx_path.stem))

    print(f"✅ {len(all_machines)} relatórios gerados em /relatorios")

    # 2. Atualizar o index.html com links de impressão
    # Aqui o script avisa o Git para subir tudo
    try:
        subprocess.run(['git', 'add', '.'], shell=True)
        subprocess.run(['git', 'commit', '-m', 'Adicionado botões de impressão'], shell=True)
        subprocess.run(['git', 'push'], shell=True)
        print("🚀 GitHub atualizado com os novos relatórios!")
    except Exception as e:
        print(f"⚠️ Erro no Git: {e}")

if __name__ == "__main__":
    run()
    input("\nFim do processo. Pressione Enter para fechar...")