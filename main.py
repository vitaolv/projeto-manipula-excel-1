import datetime

import openpyxl

import classes

try:
    # acao = input("Qual o código da ação que você quer processar?").upper()

    acao = 'BIDI4'

    indice = 2

    acess = classes.ReadAcess(caminho_arq='./dados/')
    acess.process_arq(acao)

    create = classes.manager()

    new_date = create.add(title='Dados')
    create.add_header(["DATA", "COTAÇÃO", "BANDA INFERIOR", "BANDA SUPERIOR"])
    create.cells_size('A', size=15)
    create.cells_size('B', size=15)
    create.cells_size('C', size=15)
    create.cells_size('D', size=15)

    for linha in acess.get_dates():
        # DATA: 2022-01-06 21:00:00, separamos entre data e horas.
        # acessamos apenas aa-mm-dd (eliminando hora)
        a_m_d = linha[0].split(' ')[0]
        # hours = linha[0].split(' ')[1] # acessa apenas hora.

        # separamos entre ano, mes e dia
        data = datetime.date(
            year=int(a_m_d.split('-')[0]),
            month=int(a_m_d.split('-')[1]),
            day=int(a_m_d.split('-')[2])
        )

        cotation = float(linha[1])

        # Prever como um ativo vai se comportar no movimento de preço.
        # Para isso são determinados:
        # o preço médio, mínimo e máxima de um período.
        # BANDA INFERIOR
        # Formula: média móvel (20p) - 2x desvio_padrao(20p)
        # BANDA SUPERIOR
        # Formula: média móvel (20p) + 2x desvio_padrao(20p)

        create.update_date(celula=f'A{indice}', dado=data)
        create.update_date(celula=f'B{indice}', dado=cotation)
        create.update_date(celula=f'C{indice}',
                           dado=f'=AVERAGE(B{indice}:B{indice + 19}) - 2*STDEV(B{indice}:B{indice + 19})')
        create.update_date(celula=f'D{indice}',
                           dado=f'=AVERAGE(B{indice}:B{indice + 19}) + 2*STDEV(B{indice}:B{indice + 19})')

        indice = indice + 1

    create.column_alignment()

    ########################################
    # A partir daqui criamos nova planinha #
    create.add(title='Gráfico')
    # Mesclagem de células para criação do cabeçalho do Gráfico.
    create.merge_date(cel_start='A1', cel_end='R2')

    # Aplica estilo para a planilha "Gráfico"
    create.apply_styles(
        cel='A1',
        styles=[
            ('font', openpyxl.styles.Font(b=True, sz=18, color='000000')),
            ('fill', openpyxl.styles.PatternFill('solid', fgColor='3fa2d4')),
            ('alignment', openpyxl.styles.Alignment(
                vertical="center", horizontal='center'))
        ]
    )

    ref_cotacoes = openpyxl.chart.Reference(
        new_date, min_col=2, min_row=2, max_col=4, max_row=indice)
    ref_datas = openpyxl.chart.Reference(
        new_date, min_col=1, min_row=2, max_col=1, max_row=indice)

    create.update_date('A1', 'Histórico de cotações')

    create.add_graph(
        celula='A3',
        comprimento=33.62,
        height=14.82,
        title=f'Cotações - {acao}',
        title_x='Data da cotação',
        title_y='Valor da cotação',
        ref_x=ref_cotacoes,
        ref_y=ref_datas,
        GraphPriority=[classes.SeriesGraphPriority(thick=0, color='0a55ab'),
                       classes.SeriesGraphPriority(thick=0, color='a61508'),
                       classes.SeriesGraphPriority(thick=0, color='12a154')
                       ])

    create.add_img(celula='E32', path='./recursos/cot.jpeg')

    create.save(path='./saida/Planilha.xlsx')
    print('Sucesso! A planilha foi criada e está em arquivo "saida".')

# Erro por não encontrar o atributo.
except AttributeError:
    print('O atributo não existe!')

# erro devido ao formato de dados.
except ValueError:
    print('Formato de dados incorreto! Verifique, por favor!')

# erro por não encontrar arquivo
except FileNotFoundError as er:
    print(f'O arquivo não foi encontrado! {er}.')

# outros erros
except Exception as err:
    print(f'Ops! Ocorreu algo erro durante a execução do programa. {err}')

if __name__ == '__main__':
    pass
