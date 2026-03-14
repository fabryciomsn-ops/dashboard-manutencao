#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os
import pandas as pd
from datetime import datetime
from pathlib import Path
import subprocess
import sys

# --- CONFIGURAÇÕES DE ESTILO ---
AZUL_MC = "#2d3175"
VERDE = "#22c55e"
VERMELHO = "#ef4444"

def criar_html_individual(m, periodo):
    """Gera o layout com o botão de impressão fixo no topo"""
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
            body {{ font-family: 'Segoe UI', Arial, sans-serif; margin: 40px; background-color: #f4f7f6; }}
            .no-print {{ position: sticky; top: 0; background: white; padding: 10px; text-align: right; border-bottom: 1px solid #ddd; margin-bottom: 20px; z-index: 1000; }}
            .btn-print {{ background: {AZUL_MC}; color: white; border: none; padding: 12px 24px; border-radius: 6px; cursor: pointer; font-weight: bold; font-size: 14px; box-shadow: 0 2px 4px rgba(0,0,0,0.2); }}
            .header {{ background: white; padding: 20px; border-left: 8px solid {AZUL_MC}; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); margin-bottom: 20px; }}
            .title {{ color: {AZUL_MC}; font-size: 26px; font-weight: bold; margin: 0; }}
            .kpis {{ display: flex; gap: 15px; margin: 20px 0; }}
            .card {{ background: white; padding: 15px; border-radius: 8px; flex: 1; text-align: center; border: 1px solid #e2e8f0; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }}
            .card b {{ display: block; font-size: 12px; color: #64748b; text-transform: uppercase; margin-bottom: 5px; }}
            .card span {{ font-size: 22px; font-weight: bold; }}
            table {{ width: 100%; border-collapse: collapse; background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }}
            th {{ background: {AZUL_MC}; color: white; padding: 15px; text-align: left; }}
            td {{ padding: 12px; border-bottom: 1px solid #eee; font-size: 14px; color: #444; }}
            tr:nth-child(even) {{ background: #f9fafb; }}
            .footer {{ margin-top: 40px; font-size: 11px; color: #94a3b8; text-align: center; }}
            @media print {{ .no-print {{ display: none !important; }} body {{ margin: 0; background: white; }} .header {{ box-shadow: none; border: 1px solid #ddd; }} }}
        </style>
    </head>
    <body>
        <div class="no-print">
            <button class="btn-print" onclick="window.print()">🖨️ IMPRIMIR PARA PDF / PAPEL</button>
        </div>
        <div class="header">
            <h1 class="title">{m['name']}</h1>
            <div style="margin-top:5px; color:#666;">Relatório de Manutenção Industrial | Competência: {periodo}</div>
        </div>
        <div class="kpis">
            <div class="card"><b>Disponibilidade</b><span style="color:{VERDE if m['av']>90 else VERMELHO}">{m['av']:.1f}%</span></div>
            <div class="card"><b>Ocorrências</b><span>{m['count']}</span></div>
            <div class="card"><b>Tempo Parado</b><span>{m['stop']} min</span></div>
        </div>
        <table>
            <thead><tr><th>Data</th><th>Problema / Serviço</th><th>Peças Trocadas</th><th>Tempo</th></tr></thead>
            <tbody>{eventos if eventos else '<tr><td colspan="4" style="text-align:center; padding:30px;">Nenhum problema registrado no período.</td></tr>'}</tbody>
        </table>
        <div class="footer">MultCaixa Indústria - Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}</div>
    </body>
    </html>"""

def run():
    # DETERMINA A PASTA ONDE O SCRIPT ESTÁ REALMENTE SALVO
    pasta_script = Path(os.path.abspath(os.path.dirname(sys.argv[0])))
    os.chdir(pasta_script)
    
    print(f"🚀 Iniciando em: {pasta_script}")

    # 1. Encontrar a planilha
    planilhas = list(pasta_script.glob('*.xlsx'))
    if not planilhas:
        print("❌ ERRO: Nenhuma planilha .xlsx encontrada nesta pasta!")
        input("\nPressione Enter para sair...")
        return
    
    caminho_xlsx = planilhas[0]
    print(f"📊 Processando planilha: {caminho_xlsx.name}")

    try:
        # 2. Criar a pasta de relatórios (forçado)
        pasta_relatorios = pasta_script / "relatorios_individuais"
        if not pasta_relatorios.exists():
            os.makedirs(pasta_relatorios, exist_ok=True)
            print("📁 Pasta de relatórios criada!")

        xls = pd.ExcelFile(caminho_xlsx)
        maquinas_processadas = 0

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
                try: tempo_val = int(tempo) if str(tempo).isdigit() else 0
                except: tempo_val = 0
                
                m_events.append({
                    'data': str(row['DATA'])[:10],
                    'desc': desc,
                    'pecas': str(row.get('PEÇAS TROCADAS', '-')),
                    'tempo': tempo_val
                })
                
                if desc.upper() != 'OK' and desc != '' and desc.upper() != 'NAN':
                    m_stop += tempo_val
                    m_count += 1

            dispo = ((13200 - m_stop) / 13200) * 100
            m_dados = {'name': sheet, 'av': dispo, 'stop': m_stop, 'count': m_count, 'events': m_events}

            # SALVAR HTML
            nome_arquivo = f"Relatorio_{sheet.replace(' ','_')}.html"
            caminho_final = pasta_relatorios / nome_arquivo
            
            with open(caminho_final, "w", encoding="utf-8") as f:
                f.write(criar_html_individual(m_dados, caminho_xlsx.stem))
            
            maquinas_processadas += 1

        print(f"✅ Sucesso! {maquinas_processadas} relatórios gerados em 'relatorios_individuais'.")

        # 3. GitHub (Silencioso se falhar)
        print("📤 Sincronizando GitHub...")
        subprocess.run('git add .', shell=True)
        subprocess.run(f'git commit -m "Auto-update {datetime.now().strftime("%H:%M")}"', shell=True)
        subprocess.run('git push', shell=True)
        
        print("\n✨ TUDO PRONTO!")
        
    except Exception as e:
        print(f"❌ OCORREU UM ERRO: {e}")
    
    input("\nPressione Enter para fechar...")

if __name__ == "__main__":
    run()