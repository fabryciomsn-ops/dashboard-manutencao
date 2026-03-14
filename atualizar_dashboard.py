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
    """Gera o layout do relatório individual para cada máquina"""
    eventos = "".join([
        f"<tr><td>{e['data']}</td><td>{e['desc']}</td><td>{e['pecas']}</td><td>{e['tempo']}</td></tr>"
        for e in m['events'] if str(e['desc']).upper() != 'OK'
    ])
    
    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <style>
            body {{ font-family: sans-serif; margin: 40px; color: #333; }}
            .header {{ border-bottom: 3px solid {AZUL_MC}; padding-bottom: 10px; margin-bottom: 20px; }}
            .title {{ color: {AZUL_MC}; font-size: 24px; font-weight: bold; }}
            .kpis {{ display: flex; gap: 20px; margin: 20px 0; }}
            .card {{ background: #f1f5f9; padding: 15px; border-radius: 8px; flex: 1; text-align: center; border: 1px solid #ddd; }}
            table {{ width: 100%; border-collapse: collapse; }}
            th {{ background: {AZUL_MC}; color: white; padding: 10px; text-align: left; }}
            td {{ padding: 10px; border-bottom: 1px solid #eee; font-size: 13px; }}
            .no-print {{ display: block; margin-bottom: 20px; text-align: right; }}
            @media print {{ .no-print {{ display: none; }} }}
        </style>
    </head>
    <body>
        <div class="no-print"><button onclick="window.print()" style="padding:10px; cursor:pointer;">🖨️ Imprimir PDF</button></div>
        <div class="header">
            <div class="title">Relatório Técnico: {m['name']}</div>
            <div>Competência: {periodo}</div>
        </div>
        <div class="kpis">
            <div class="card"><b>Disponibilidade</b><br><span style="color:{VERDE if m['av']>90 else VERMELHO}">{m['av']:.1f}%</span></div>
            <div class="card"><b>Ocorrências</b><br>{m['count']}</div>
            <div class="card"><b>Tempo Parado</b><br>{m['stop']} min</div>
        </div>
        <table>
            <thead><tr><th>Data</th><th>Descrição do Serviço</th><th>Peças</th><th>Tempo</th></tr></thead>
            <tbody>{eventos if eventos else '<tr><td colspan="4" style="text-align:center;">Nenhuma ocorrência registrada.</td></tr>'}</tbody>
        </table>
        <div style="margin-top:30px; font-size:10px; color:#999; text-align:center;">Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}</div>
    </body>
    </html>"""

def run():
    # Define a pasta atual de forma segura
    base_path = Path(__file__).parent.absolute()
    os.chdir(base_path)
    
    print(f"📂 Pasta de trabalho: {base_path}")
    
    # 1. Busca a planilha
    arquivos = list(base_path.glob('*.xlsx'))
    if not arquivos:
        print("❌ ERRO: Nenhuma planilha .xlsx encontrada na pasta!")
        return
    
    planilha_path = arquivos[0]
    print(f"📊 Lendo dados de: {planilha_path.name}")

    try:
        xls = pd.ExcelFile(planilha_path)
        machines = []
        pasta_out = base_path / "relatorios_individuais"
        pasta_out.mkdir(exist_ok=True)

        for sheet in xls.sheet_names:
            if "Resina" not in sheet and "Gel Coat" not in sheet: continue
            
            df = pd.read_excel(xls, sheet_name=sheet, skiprows=6)
            df.columns = [str(c).strip().replace('\n', ' ') for c in df.columns]
            
            if 'DATA' not in df.columns: continue
            
            df = df.dropna(subset=['DATA'])
            m_stop = 0
            m_count = 0
            m_events = []

            for _, row in df.iterrows():
                desc = str(row.get('Problemas e Serviços', 'OK')).strip()
                tempo = row.get('TEMPO DE PARADA (min)', 0)
                try: 
                    tempo_val = int(tempo) if str(tempo).isdigit() else 0
                except: 
                    tempo_val = 0
                
                m_events.append({
                    'data': str(row['DATA'])[:10],
                    'desc': desc,
                    'pecas': str(row.get('PEÇAS TROCADAS', '-')),
                    'tempo': tempo_val
                })
                
                if desc.upper() != 'OK' and desc != '':
                    m_stop += tempo_val
                    m_count += 1

            dispo = ((13200 - m_stop) / 13200) * 100
            
            m_dados = {
                'name': sheet,
                'av': dispo, 'stop': m_stop, 'count': m_count, 'events': m_events
            }
            machines.append(m_dados)

            # Salva Relatório Individual
            filename = f"Relatorio_{sheet.replace(' ','_')}.html"
            with open(pasta_out / filename, "w", encoding="utf-8") as f:
                f.write(criar_html_individual(m_dados, planilha_path.stem))

        print(f"✅ {len(machines)} relatórios gerados em /relatorios_individuais")

        # 2. Sincronização com GitHub
        print("📤 Sincronizando com GitHub...")
        # Força o Git a usar o shell do sistema para evitar erros de caminho
        commands = [
            'git add .',
            f'git commit -m "Atualizacao automatica {datetime.now().strftime("%d/%m %H:%M")}"',
            'git push'
        ]
        
        for cmd in commands:
            subprocess.run(cmd, shell=True, cwd=base_path)
            
        print("🚀 Processo concluído com sucesso!")

    except Exception as e:
        print(f"❌ ERRO CRÍTICO: {e}")

if __name__ == "__main__":
    run()