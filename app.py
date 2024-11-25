from flask import Flask, request, render_template, send_file
import pandas as pd
from datetime import datetime
import os
from openpyxl import load_workbook
from openpyxl.styles import Alignment, NamedStyle, Font

app = Flask(__name__)

UPLOAD_FOLDER = "resultados"
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

@app.route("/")
def index():
    return render_template("upload.html")

@app.route("/processar", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return "Nenhum arquivo enviado", 400

    file = request.files["file"]
    if file.filename == "":
        return "Nenhum arquivo selecionado", 400

    try:
        # Lê o arquivo Excel enviado
        tabela = pd.read_excel(file)

        # Data atual para nomear o arquivo
        hoje = datetime.today().strftime("%d-%m-%Y")

        # Processar os dados agrupados por contrato
        relatorio = []
        contratos = tabela["CódigoContrato"].unique()

        for contrato in contratos:
            subset = tabela[tabela["CódigoContrato"] == contrato]

            # Obter o nome do loteamento (supondo que seja a mesma para todos os contratos)
            nome_loteamento = subset.iloc[0]["NomeLoteamento"]

            # Calcular Valor Recebido, VAlor total, valor recebido e valor atrasado
            recebido = subset["Valor Pagamento"].sum()
            valorTotalPrincipal = subset["Valor Parcela Inicial"].sum()
            receber = subset.loc[(subset["DataVencimento"] > datetime.today())& (subset["Valor Pagamento"].isna()), "Valor Total"].sum()
            atrasado = subset.loc[
                (subset["DataVencimento"] < datetime.today()) & (subset["Valor Pagamento"].isna()), 
                "Valor Total"
            ].sum()

            # Encontrar a última data de pagamento
            ultimo_pagamento = subset.loc[subset["Valor Pagamento"].notna(), "MesAnoRef"].max()

            # Verificar se há uma data de cancelamento
            data_cancelamento = subset.iloc[0]["Data Cancelamento"] if "Data Cancelamento" in subset.columns else None
            if pd.notna(data_cancelamento):
                data_cancelamento = pd.to_datetime(data_cancelamento).strftime("%d/%m/%Y")

            # Calcular os valores de Reajuste, Juros e Multa
            valor_reajuste = subset["Valor Reajuste"].sum()
            valor_juros = subset["Valor Juros"].sum()
            valor_multa = subset["Valor Multa"].sum()
            
            #Mostrar a primeira data de vencimento
            primeiraDataVencimento = subset["DataVencimento"].min()
            if pd.notna(primeiraDataVencimento):
                primeiraDataVencimento = pd.to_datetime(primeiraDataVencimento).strftime("%d/%m/%Y")

            # Adiciona uma linha consolidada
            relatorio.append({
                "CódigoContrato": contrato,
                "CódigoCliente": subset.iloc[0]["CódigoCliente"],
                "NomeCliente": subset.iloc[0]["NomeCliente"],
                "Valor Total Principal": valorTotalPrincipal,
                "Valor Reajuste": valor_reajuste,
                "Valor Juros": valor_juros,
                "Valor Multa": valor_multa,
                "Valor a Receber": receber,
                "Valor Atrasado": atrasado,
                "Recebido": recebido,
                "Ultima Data de Pagamento": ultimo_pagamento,
                "Data de Cancelamento": data_cancelamento,
                "Primeiro Vencimento": primeiraDataVencimento
            })

        # Cria o DataFrame consolidado
        df_relatorio = pd.DataFrame(relatorio)

        # Salva o relatório consolidado em Excel com o nome do loteamento e data
        relatorio_path = os.path.join(UPLOAD_FOLDER, f"{nome_loteamento}_{hoje}.xlsx")
        df_relatorio.to_excel(relatorio_path, index=False)

        # Ajustar formatação no Excel
        workbook = load_workbook(relatorio_path)
        sheet = workbook.active

        # Centralizar e formatar as colunas
        money_style = NamedStyle(name="money_style", number_format="R$ #,##0.00")
        alignment = Alignment(horizontal="center", vertical="center")
        font = Font(name="Arial", size=11)

        for col in sheet.columns:
            for cell in col:
                cell.alignment = alignment
                cell.font = font

                # Formatar colunas numéricas como dinheiro
                if cell.column_letter in ["D", "E", "F", "G", "H", "I", "J"]:  # Ajustar de acordo com as colunas de valores
                    cell.style = money_style

        # Salvar as alterações no arquivo
        workbook.save(relatorio_path)
        workbook.close()

        return send_file(relatorio_path, as_attachment=True)

    except Exception as e:
        return f"Erro ao processar o arquivo: {e}", 500

if __name__ == "__main__":
    app.run(debug=True)
