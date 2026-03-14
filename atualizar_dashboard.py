#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os
import pandas as pd
from datetime import datetime
from pathlib import Path
import subprocess
import sys

# --- CONFIGURAÇÕES DE CORES ---
AZUL_MC = "#2d3175"
VERDE = "#22c55e"
VERMELHO = "#ef4444"

def criar_html_individual(m, periodo):
    """Gera o relatório que será impresso"""
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
            .btn-print {{ background: {AZUL_MC}; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; font-weight: bold; }}
            .header {{ border-bottom: 3px solid {AZUL_MC}; padding-bottom: 10px; margin-bottom: 20px; }}
            .card {{ background: #f8fafc; padding: 15px; border-radius: 8px; border: 1px solid #ddd; text-align: center; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
            th {{ background: {AZUL_MC}; color: white; padding: 10px; text-align: left; }}
            td {{ padding: 8px; border-bottom: 1px solid #eee; }}
            @media print {{ .no-print {{ display: none; }} }}
        </style>
    </head>
    <body>
        <div class="no-print"><button class="btn-print" onclick="window.print()">🖨️ IMPRIMIR / SALVAR PDF</button></div>
        <div class="header"><h2>Relatório: {m['name']}</h2><p>Período: {periodo}</p></div>
        <div style="display:flex; gap:10px;">
            <div class="card" style="flex:1;"><b>Disponibilidade</b><br>{m['av']:.1f}%</div>
            <div class="card" style="flex:1;"><b>Paradas</b><br>{m['count']}</div>
        </div>
        <table>
            <thead><tr><th>Data</th><th>Serviço</th><th>Peças</th><th>Tempo</th></tr></thead>
            <tbody>{eventos if eventos else '<tr><td colspan="4">Sem paradas registradas.</td></tr>'}</tbody>
        </table>
    </body>
    </html>"""

def run():
    base_path = Path(os.path.abspath(os.path.dirname(sys.argv[0])))
    os.chdir(base_path)
    
    planilhas = list(base_path.glob('*.xlsx'))
    if not planilhas:
        print("❌ Erro: Planilha não encontrada na pasta.")
        return
    
    xlsx = planilhas[0]
    xls = pd.ExcelFile(xlsx)
    
    # Criar pasta de relatórios
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
        relat_name = f"relatorio_{sheet.replace(' ','_')}.html"
        
        m_info = {
            'name': sheet, 'av': dispo, 'stop': m_stop, 'count': m_count, 
            'events': m_events, 'link': f"relatorios/{relat_name}"
        }
        machines_data.append(m_info)

        # Salvar Relatório Individual
        with open(pasta_relatorios / relat_name, "w", encoding="utf-8") as f:
            f.write(criar_html_individual(m_info, xlsx.stem))

    # --- ATUALIZAR O DASHBOARD (index.html) ---
    print("📝 Atualizando o Dashboard Principal...")
    cards_html = ""
    for m in machines_data:
        cor = VERDE if m['av'] > 90 else VERMELHO
        cards_html += f"""
        <div style="border:1px solid #ddd; padding:15px; border-radius:8px; background:white; position:relative;">
            <a href="{m['link']}" target="_blank" style="float:right; background:{AZUL_MC}; color:white; padding:5px 10px; text-decoration:none; border-radius:4px; font-size:12px; font-weight:bold;">📄 GERAR PDF</a>
            <h3 style="margin:0; color:{AZUL_MC};">{m['name']}</h3>
            <p style="margin:5px 0;">Disponibilidade: <b style="color:{cor};">{m['av']:.1f}%</b></p>
            <p style="margin:5px 0; font-size:13px; color:#666;">Tempo Parado: {m['stop']} min</p>
        </div>"""

    index_content = f"""
    <!DOCTYPE html>
    <html lang="pt-BR">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Dashboard MultCaixa</title>
        <style>
            body {{ font-family: sans-serif; background:#f4f7f6; margin:0; padding:20px; }}
            .grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 20px; }}
            header {{ background:{AZUL_MC}; color:white; padding:20px; border-radius:8px; margin-bottom:20px; text-align:center; }}
        </style>
    </head>
    <body>
        <header>
            <h1>Manutenção Industrial - MultCaixa</h1>
            <p>Atualizado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
        </header>
        <div class="grid">{cards_html}</div>
    </body>
    </html>"""

    with open(base_path / "index.html", "w", encoding="utf-8") as f:
        f.write(index_content)

    # Subir para o GitHub
    try:
        subprocess.run(['git', 'add', '.'], shell=True)
        subprocess.run(['git', 'commit', '-m', 'Fix: Botão PDF restaurado'], shell=True)
        subprocess.run(['git', 'push'], shell=True)
        print("🚀 SUCESSO! Site atualizado no GitHub.")
    except:
        print("⚠️ Erro ao enviar para o GitHub, mas os ficheiros locais estão prontos.")

if __name__ == "__main__":
    run()