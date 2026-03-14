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
    """Gera o relatório técnico para impressão"""
    eventos = "".join([
        f"<tr><td>{e['data']}</td><td>{e['desc']}</td><td>{e['pecas']}</td><td>{e['tempo']} min</td></tr>"
        for e in m['events'] if str(e['desc']).upper() != 'OK'
    ])
    
    return f"""
    <!DOCTYPE html>
    <html lang="pt-BR">
    <head>
        <meta charset="UTF-8">
        <style>
            body {{ font-family: sans-serif; margin: 40px; color: #333; }}
            .no-print {{ text-align: right; margin-bottom: 20px; }}
            .btn-print {{ background: {AZUL_MC}; color: white; border: none; padding: 12px 20px; border-radius: 5px; cursor: pointer; font-weight: bold; font-size: 14px; }}
            .header {{ border-bottom: 4px solid {AZUL_MC}; padding-bottom: 10px; margin-bottom: 20px; }}
            .kpis {{ display: flex; gap: 15px; margin-bottom: 20px; }}
            .card-kpi {{ background: #f8fafc; padding: 15px; border-radius: 8px; border: 1px solid #ddd; flex: 1; text-align: center; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
            th {{ background: {AZUL_MC}; color: white; padding: 10px; text-align: left; }}
            td {{ padding: 8px; border-bottom: 1px solid #eee; font-size: 13px; }}
            @media print {{ .no-print {{ display: none; }} body {{ margin: 0; }} }}
        </style>
    </head>
    <body>
        <div class="no-print"><button class="btn-print" onclick="window.print()">🖨️ IMPRIMIR / GERAR PDF</button></div>
        <div class="header">
            <h2 style="margin:0; color:{AZUL_MC}">{m['name']}</h2>
            <p style="margin:5px 0; color:#666;">Relatório de Manutenção - Período: {periodo}</p>
        </div>
        <div class="kpis">
            <div class="card-kpi"><b>Disponibilidade</b><br><span style="font-size:18px; color:{VERDE if m['av']>90 else VERMELHO}">{m['av']:.1f}%</span></div>
            <div class="card-kpi"><b>Total de Paradas</b><br><span style="font-size:18px;">{m['count']}</span></div>
            <div class="card-kpi"><b>Tempo Parado</b><br><span style="font-size:18px;">{m['stop']} min</span></div>
        </div>
        <table>
            <thead><tr><th>Data</th><th>Serviço/Problema</th><th>Peças</th><th>Tempo</th></tr></thead>
            <tbody>{eventos if eventos else '<tr><td colspan="4" style="text-align:center;">Nenhuma intercorrência no período.</td></tr>'}</tbody>
        </table>
    </body>
    </html>"""

def run():
    base_path = Path(os.path.abspath(os.path.dirname(sys.argv[0])))
    os.chdir(base_path)
    
    planilhas = list(base_path.glob('*.xlsx'))
    if not planilhas:
        print("❌ Nenhuma planilha encontrada.")
        return
    
    xlsx = planilhas[0]
    xls = pd.ExcelFile(xlsx)
    
    pasta_relatorios = base_path / "relatorios"
    pasta_relatorios.mkdir(exist_ok=True)
    
    machines_data = []

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
            if desc.upper() != 'OK' and desc != '':
                m_stop += t_val
                m_count += 1

        dispo = ((13200 - m_stop) / 13200) * 100
        relat_file = f"relatorio_{sheet.replace(' ','_')}.html"
        
        m_info = {'name': sheet, 'av': dispo, 'stop': m_stop, 'count': m_count, 'events': m_events, 'link': f"relatorios/{relat_file}"}
        machines_data.append(m_info)

        with open(pasta_relatorios / relat_file, "w", encoding="utf-8") as f:
            f.write(criar_html_individual(m_info, xlsx.stem))

    # --- GERAÇÃO DO NOVO INDEX.HTML ---
    cards_html = ""
    for m in machines_data:
        cor_av = VERDE if m['av'] > 90 else VERMELHO
        cards_html += f"""
        <div style="background:white; border-radius:10px; padding:20px; box-shadow:0 2px 5px rgba(0,0,0,0.1); border-top: 5px solid {AZUL_MC};">
            <div style="display:flex; justify-content:space-between; align-items:flex-start;">
                <h3 style="margin:0; color:{AZUL_MC}">{m['name']}</h3>
                <a href="{m['link']}" target="_blank" style="background:{AZUL_MC}; color:white; padding:6px 12px; text-decoration:none; border-radius:5px; font-size:12px; font-weight:bold;">📄 ABRIR / IMPRIMIR</a>
            </div>
            <p style="margin:15px 0 5px 0; font-size:14px;">Disponibilidade: <b style="color:{cor_av}; font-size:18px;">{m['av']:.1f}%</b></p>
            <div style="background:#eee; height:8px; border-radius:4px; margin-bottom:15px;"><div style="background:{cor_av}; width:{m['av']}%; height:100%; border-radius:4px;"></div></div>
            <p style="margin:0; font-size:12px; color:#666;">Paradas: {m['count']} | Tempo: {m['stop']} min</p>
        </div>"""

    index_content = f"""
    <!DOCTYPE html>
    <html lang="pt-BR">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Dashboard Manutenção - MultCaixa</title>
        <style>
            body {{ font-family: sans-serif; background:#f4f7f6; margin:0; padding:20px; }}
            .container {{ max-width: 1200px; margin: auto; }}
            .grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 20px; }}
            header {{ background:{AZUL_MC}; color:white; padding:20px; border-radius:10px; margin-bottom:20px; text-align:center; }}
        </style>
    </head>
    <body>
        <div class="container">
            <header>
                <h1 style="margin:0;">Manutenção Industrial</h1>
                <p style="margin:5px 0 0 0; opacity:0.8;">Dashboard de Disponibilidade de Máquinas</p>
            </header>
            <div class="grid">{cards_html}</div>
            <footer style="margin-top:30px; text-align:center; font-size:12px; color:#999;">
                Atualizado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}
            </footer>
        </div>
    </body>
    </html>"""

    with open(base_path / "index.html", "w", encoding="utf-8") as f:
        f.write(index_content)

    # SUBIR PARA GITHUB
    try:
        subprocess.run(['git', 'add', '.'], shell=True)
        subprocess.run(['git', 'commit', '-m', 'Atualização: Dashboard com botões de impressão'], shell=True)
        subprocess.run(['git', 'push'], shell=True)
        print("🚀 Dashboard atualizado com sucesso no GitHub!")
    except Exception as e:
        print(f"⚠️ Erro ao subir para o GitHub: {e}")

if __name__ == "__main__":
    run()