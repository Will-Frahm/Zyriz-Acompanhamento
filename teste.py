import re
import pandas as pd

def parse_line_regex(line):
    # Regex para data no início da linha (DD/MM/AAAA)
    data_match = re.match(r'(\d{2}/\d{2}/\d{4})\s+', line)
    if not data_match:
        return None
    data = data_match.group(1)
    rest = line[data_match.end():]

    # Capturar os valores monetários e contra partida (no final da linha)
    valores_pattern = re.compile(r'(\d{1,3}(?:\.\d{3})*,\d{2}|\d+,?\d*)\s+(\d{1,3}(?:\.\d{3})*,\d{2}|\d+,?\d*)\s+(.*)$')
    valores_match = valores_pattern.search(rest)
    if not valores_match:
        return None
    movto_debito_raw = valores_match.group(1)
    movto_credito_raw = valores_match.group(2)
    contra_partida = valores_match.group(3).strip()

    before_vals = rest[:valores_match.start()].rstrip()

    # Capturar Lote/Lancto/Seq
    lote_pattern = re.compile(r'(\d+/\d+/\d+)')
    lote_match = lote_pattern.search(before_vals)
    if not lote_match:
        return None
    lote = lote_match.group(1)

    after_lote = before_vals[lote_match.end():].strip()

    # Capturar Ori, Estab, UN (primeiros 3 tokens depois do lote)
    tokens = after_lote.split()
    if len(tokens) < 3:
        return None
    ori = tokens[0]
    estab = tokens[1]
    un = tokens[2]

    historico = before_vals[:lote_match.start()].strip()

    def parse_float(val):
        val = val.replace('.', '').replace(',', '.')
        try:
            return float(val)
        except:
            return 0.0

    movto_debito = parse_float(movto_debito_raw)
    movto_credito = parse_float(movto_credito_raw)

    return {
        'Data': data,
        'Histórico': historico,
        'Lote/Lancto/Seq': lote,
        'Ori': ori,
        'Estab': estab,
        'UN': un,
        'Movto Débito': movto_debito,
        'Movto Crédito': movto_credito,
        'Contra Partida': contra_partida,
    }

def processar_arquivo_txt(caminho_arquivo):
    registros = []
    historico_extra = ""
    linha_atual = ""

    with open(caminho_arquivo, 'r', encoding='latin-1') as f:
        linhas = f.readlines()

    for i, linha in enumerate(linhas):
        linha = linha.rstrip('\n\r')

        # Ignorar linhas de cabeçalho, rodapé e separadores
        if (linha.startswith('----------------------------------------------------------------') or
            linha.startswith('Página:') or
            linha.strip() == '' or
            linha.startswith('AUDIOFRAHM') or
            linha.startswith('Período:') or
            linha.startswith('Cenário Contábil:') or
            linha.startswith('Data       Histórico')):
            continue

        # Se a linha começa com data, é nova entrada
        if re.match(r'\d{2}/\d{2}/\d{4}', linha):
            # Se tem linha_atual com dado, parsear ela
            if linha_atual:
                # Adiciona historico extra antes do parse
                linha_completa = linha_atual + " " + historico_extra.strip()
                registro = parse_line_regex(linha_completa)
                if registro:
                    registros.append(registro)

            linha_atual = linha
            historico_extra = ""
        else:
            # Linha de histórico extra (continuação), junta ao histórico_extra
            historico_extra += " " + linha.strip()

    # Processa o último registro
    if linha_atual:
        linha_completa = linha_atual + " " + historico_extra.strip()
        registro = parse_line_regex(linha_completa)
        if registro:
            registros.append(registro)

    df = pd.DataFrame(registros)
    return df

if __name__ == "__main__":
    arquivo = "razao_limpo.txt"  # ajuste para o nome do seu arquivo txt
    df = processar_arquivo_txt(arquivo)
    df.to_excel("saida_razao.xlsx", index=False)
    print("Arquivo Excel gerado: saida_razao.xlsx")
