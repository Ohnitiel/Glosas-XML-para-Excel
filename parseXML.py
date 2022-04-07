import os
from lxml.etree import parse, fromstring
import pandas as pd

codigosGlosa = pd.read_csv("codigosGlosa.csv", header=None, index_col=0)
codigosGlosa = codigosGlosa.to_dict()[1]


def parseXML(path):
    if len(path) < 257 and os.path.isfile(path):
            doc = parse(path)
    else:
        doc = fromstring(path).getroottree()
    dfProcedimentos = glosaProcedimento(doc)
    dfGuias = glosaGuia(doc)
    df1 = pd.DataFrame()
    df2 = pd.DataFrame()
    dfrel1 = pd.DataFrame()
    dfrel2 = pd.DataFrame()
    if not dfProcedimentos.empty:
        df1 = dfProcedimentos.groupby([
            'Protocolo', 'Código Glosa', 'Descrição'])['Valor Glosa'] \
            .sum().round(2).reset_index()
        dfrel1 = dfProcedimentos.groupby([
            'Código Glosa', 'Descrição', 'Nome', 'Número Guia', 'Protocolo']) \
            .agg({'Data': 'min', 'Valor Glosa': 'sum'}).reset_index()
    if not dfGuias.empty:
        df2 = dfGuias.groupby([
            'Protocolo', 'Código Glosa', 'Descrição'])['Valor Glosa'] \
            .sum().round(2).reset_index()
        dfrel2 = dfProcedimentos.groupby([
            'Código Glosa', 'Descrição', 'Nome', 'Número Guia', 'Protocolo']) \
            .agg({'Data': 'min', 'Valor Glosa': 'sum'}).reset_index()
    dfCodigos = pd.concat([df1, df2])
    dfRelatorio = pd.concat([dfrel1, dfrel2])
    if not dfRelatorio.empty:
        dfRelatorio = dfRelatorio[['Código Glosa', 'Descrição', 'Nome',
                                'Número Guia', 'Data', 'Valor Glosa',
                                'Protocolo']]
    return dfCodigos, dfGuias, dfProcedimentos, dfRelatorio


def glosaGuia(doc):
    motivoGlosa = doc.findall("//{*}motivoGlosaGuia")
    relacaoGuia = [x.getparent() for x in motivoGlosa]
    tipoGlosa = [x.find("{*}codigoGlosa").text for x in motivoGlosa]
    try:
        valorGlosa = [
            x.find("{*}valorGlosaGuia").text for x in relacaoGuia
        ]
    except AttributeError:
        return pd.DataFrame()
    datas_inicio = [
        x.find("{*}dataInicioFat").text for x in relacaoGuia
    ]
    try:
        datas_fim = [
            x.find("{*}dataFimFat").text for x in relacaoGuia
        ]
    except AttributeError:
        datas_fim = [
            x.find("{*}dataInicioFat").text for x in relacaoGuia
        ]
    nomes = [
        x.find("{*}nomeBeneficiario").text for x in relacaoGuia
    ]
    matriculas = [
        x.find("{*}numeroCarteira").text for x in relacaoGuia
    ]
    try:
        senhas = [
            x.find("{*}senha").text for x in relacaoGuia
        ]
    except AttributeError:
        senhas = ["-"] * len(relacaoGuia)
    numerosGuia = [
        x.find("{*}numeroGuiaPrestador").text for x in relacaoGuia
    ]
    protocolo = [
        doc.find('//{*}dadosProtocolo').find("{*}numeroProtocolo").text
    ] * len(relacaoGuia)

    valorGlosa = pd.Series(valorGlosa)
    tipoGlosa = pd.Series(tipoGlosa)
    datas_inicio = pd.Series(datas_inicio)
    datas_fim = pd.Series(datas_fim)
    nomes = pd.Series(nomes)
    matriculas = pd.Series(matriculas)
    senhas = pd.Series(senhas)
    numerosGuia = pd.Series(numerosGuia)
    protocolo = pd.Series(protocolo)

    series = [
        protocolo,
        numerosGuia,
        datas_inicio,
        datas_fim,
        matriculas,
        senhas,
        nomes,
        valorGlosa,
        tipoGlosa,
    ]
    columns = [
        "Protocolo",
        "Número Guia",
        "Data Inicial",
        "Data Final",
        "Matrícula",
        "Senha",
        "Nome",
        "Valor Glosa",
        "Código Glosa",
    ]

    df = pd.concat(series, axis=1, keys=columns)
    df = df.applymap(lambda x: pd.to_numeric(x, errors='ignore'))
    df['Descrição'] = df['Código Glosa'].map(codigosGlosa)

    return df


def glosaProcedimento(doc):
    valorGlosa = doc.findall("//{*}valorGlosa")
    tipoGlosa = doc.findall("//{*}tipoGlosa")
    detalhesGuia = [x.getparent().getparent() for x in valorGlosa]
    valorGlosa = [x.text for x in valorGlosa]
    tipoGlosa = [x.text for x in tipoGlosa]
    procedimentos = [
        x.find("{*}procedimento/{*}codigoProcedimento").text for x in detalhesGuia
    ]
    datas = [
        x.find("{*}dataRealizacao").text for x in detalhesGuia
    ]
    nomes = [
        x.getparent().find("{*}nomeBeneficiario").text for x in detalhesGuia
    ]
    matriculas = [
        x.getparent().find("{*}numeroCarteira").text for x in detalhesGuia
    ]
    try:
        senhas = [
            x.getparent().find("{*}senha").text for x in detalhesGuia
        ]
    except AttributeError:
        senhas = ["-"] * len(detalhesGuia)
    numerosGuia = [
        x.getparent().find("{*}numeroGuiaPrestador").text for x in detalhesGuia
    ]
    protocolo = [
        doc.find('//{*}dadosProtocolo').find("{*}numeroProtocolo").text
    ] * len(detalhesGuia)

    valorGlosa = pd.Series(valorGlosa)
    tipoGlosa = pd.Series(tipoGlosa)
    procedimentos = pd.Series(procedimentos)
    datas = pd.Series(datas)
    nomes = pd.Series(nomes)
    matriculas = pd.Series(matriculas)
    senhas = pd.Series(senhas)
    numerosGuia = pd.Series(numerosGuia)
    protocolo = pd.Series(protocolo)

    series = [
        protocolo,
        numerosGuia,
        datas,
        matriculas,
        senhas,
        nomes,
        procedimentos,
        valorGlosa,
        tipoGlosa,
    ]
    columns = [
        "Protocolo",
        "Número Guia",
        "Data",
        "Matrícula",
        "Senha",
        "Nome",
        "Procedimento",
        "Valor Glosa",
        "Código Glosa",
    ]

    df = pd.concat(series, axis=1, keys=columns)
    df = df.applymap(lambda x: pd.to_numeric(x, errors='ignore'))
    df['Descrição'] = df['Código Glosa'].map(codigosGlosa)

    return df
