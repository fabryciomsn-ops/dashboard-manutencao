#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os
import pandas as pd
from datetime import datetime
from pathlib import Path
import subprocess

# --- CONFIGURAÇÕES DE ESTILO ---
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
            body {{ font-family: 'Segoe UI', Arial, sans-serif; margin: 40px; color: #333; }}
            .no-print {{ text-align: right; margin-bottom: 20px; }}
            .btn-print {{ background: {AZUL_MC}; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; font-weight: bold; }}
            .header {{ border-bottom: 3px solid {AZUL_MC}; padding-bottom: 10px; margin-bottom: 20px; }}
            .title {{ color: {AZUL_MC}; font-size: 24px; font-weight: bold; }}
            .kpis {{ display: flex; gap: 20px; margin: 20px 0; }}
            .card {{ background: #f8fafc; padding: 15px; border-radius: 8px; flex: 1; text-align: center; border: 1px solid #e2e8f0; }}
            .card b {{ display: block; font-size: 12px; color: #64748b; text-transform: uppercase; }}
            .card span {{ font-size: 20px; font-weight: bold; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
            th {{ background: {AZUL_MC}; color: white; padding: 12px; text-align: left; }}
            td {{ padding: 10px; border-bottom: 1px solid #eee; font-size: 13px; }}
            tr:nth-child(even) {{ background: #f1f5f9; }}
            .footer {{ margin-top: 40px; font-size: 10px; color: #94a3b8; text-align: center; }}
            @media print {{ .no-print {{ display: none; }} body {{ margin: 10px; }} }}
        </style>
    </head>
    <body>
        <div class="no-print">
            <button class="btn-print" onclick="window.print()">🖨️ IMPRIMIR RELATÓRIO / SALVAR PDF</button>
        </div>
        <div class="header">
            <div class="title">Ficha Técnica: {m['name']}</div>
            <div style="color: #64748b;">Competência: {periodo} | MultCaixa</div>
        </div>
        <div class="kpis">
            <div class="card"><b>Disponibilidade</b><span style="color:{VERDE if m['av']>90 else VERMELHO}">{m['av']:.1f}%</span></div>
            <div class="card"><b>Ocorrências</b><span>{m['count']}</span></div>
            <div class="card"><b>Tempo Parado</b><span>{m['stop']} min</span></div>
        </div>
        <table>
            <thead><tr><th>Data</th><th>Descrição do Serviço</th><th>Peças</th><th>Tempo</th></tr></thead>
            <tbody>{eventos if eventos else '<tr><td colspan="4" style="text-align:center;">Nenhuma intercorrência registada.</td></tr>'}</tbody>
        </table>
        <div class="footer">Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}</div>
    </body>
    </html>"""

def run():
    # Garante que o script usa a pasta onde ele está localizado
    base_dir = Path(__file__).parent.absolute()
    os.chdir(base_dir)
    
    # 1. Localizar Planilha
    planilhas = list(base_dir.glob('*.xlsx'))
    if not planilhas:
        print("❌ Ficheiro .xlsx não encontrado!")
        return
    
    planilha_path = planilhas[0]
    print(f"📊 A ler: {planilha_path.name}")

    try:
        # 2. Criar pasta de relatórios
        pasta_relatorios = base_dir / "relatorios_individuais"
        if not pasta_relatorios.exists():
            pasta_relatorios.mkdir(parents=True)
            print("📁 Pasta 'relatorios_individuais' criada.")

        xls = pd.ExcelFile(planilha_path)
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
                try: tempo_val = int(tempo) if str(tempo).isdigit() else 0
                except: tempo_val = 0
                
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
            m_dados = {'name': sheet, 'av': dispo, 'stop': m_stop, 'count': m_count, 'events': m_events}
            machines_data.append(m_dados)

            # Gravar o Ficheiro HTML Individual
            nome_arq = f"Relatorio_{sheet.replace(' ','_')}.html"
            with open(pasta_relatorios / nome_arq, "w", encoding="utf-8") as f:
                f.write(criar_html_individual(m_dados, planilha_path.stem))

        print(f"✅ {len(machines_data)} relatórios individuais prontos na pasta!")

        # 3. Atualizar o Dashboard Principal (index.html)
        # (Aqui o script mantém a sua função original de push para o Git)
        subprocess.run(['git', 'add', '.'], shell=True)
        subprocess.run(['git', 'commit', '-m', 'Relatorios e Dashboard Atualizados'], shell=True)
        subprocess.run(['git', 'push'], shell=True)
        print("🚀 GitHub atualizado!")

    except Exception as e:
        print(f"❌ Erro: {e}")

if __name__ == "__main__":
    run()