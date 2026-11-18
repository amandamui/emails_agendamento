"""Regras de negócio da cobrança de notas atrasadas (fonte única).

Atraso = dias ÚTEIS decorridos desde a data da Agenda (sem contar o próprio dia
da agenda). A nota só é cobrada a partir do 4º dia útil de atraso.

Pensado para a rotina semanal (toda terça-feira): todas as funções aceitam uma
`data_referencia`, para que o cálculo seja reproduzível (ex.: reprocessar "a
terça passada") e testável. Sem ela, usa a data de hoje.
"""
import datetime as dt

import numpy as np
import pandas as pd

DIAS_UTEIS_MINIMO = 4   
DIAS_CORRIDOS_CRITICO = 60
DIA_ENVIO_SEMANAL = 1
FERIADOS_EXTRAS: set[dt.date] = set()

def _pascoa(ano: int) -> dt.date:
    """Domingo de Páscoa (algoritmo de Meeus/Jones/Butcher)."""
    a, b, c = ano % 19, ano // 100, ano % 100
    d, e = b // 4, b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = c // 4, c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    mes = (h + l - 7 * m + 114) // 31
    dia = (h + l - 7 * m + 114) % 31 + 1
    return dt.date(ano, mes, dia)


def feriados_nacionais(ano: int) -> set[dt.date]:
    """Feriados nacionais + Carnaval e Corpus Christi (pontos facultativos usuais)."""
    pascoa = _pascoa(ano)
    fixos = [(1, 1), (4, 21), (5, 1), (9, 7), (10, 12), (11, 2), (11, 15), (11, 20), (12, 25)]
    moveis = [
        pascoa - dt.timedelta(days=48),  # Carnaval (segunda)
        pascoa - dt.timedelta(days=47),  # Carnaval (terça)
        pascoa - dt.timedelta(days=2),   # Sexta-feira Santa
        pascoa + dt.timedelta(days=60),  # Corpus Christi
    ]
    return {dt.date(ano, m, d) for m, d in fixos} | set(moveis)


def _feriados(anos) -> list[np.datetime64]:
    datas = set(FERIADOS_EXTRAS)
    for ano in anos:
        datas |= feriados_nacionais(ano)
    return sorted(np.datetime64(d, "D") for d in datas)


def hoje() -> pd.Timestamp:
    return pd.Timestamp.today().normalize()


def dias_uteis_atraso(agenda: pd.Series, data_referencia: pd.Timestamp | None = None) -> pd.Series:
    """Dias úteis entre a Agenda (exclusive) e a data de referência (inclusive).

    Ex.: agenda numa segunda → terça=1, quarta=2, quinta=3, sexta=4 (cobra).
    Agenda vazia/inválida → <NA>. Agenda no futuro → 0.
    """
    ref = (data_referencia or hoje()).normalize()
    datas = pd.to_datetime(agenda, dayfirst=True, errors="coerce").dt.normalize()
    validas = datas.notna()
    resultado = pd.Series(pd.NA, index=agenda.index, dtype="Int64")
    if not validas.any():
        return resultado

    anos = range(min(datas[validas].min().year, ref.year), ref.year + 1)
    inicio = (datas[validas] + pd.Timedelta(days=1)).values.astype("datetime64[D]")
    fim = np.datetime64(ref + pd.Timedelta(days=1), "D")
    contagem = np.busday_count(inicio, fim, holidays=_feriados(anos))
    resultado[validas] = np.maximum(contagem, 0)
    return resultado


def dias_corridos_desde(datas: pd.Series, data_referencia: pd.Timestamp | None = None) -> pd.Series:
    """Dias corridos entre cada data (ex.: Emissão) e a data de referência, como HOJE()-data no Excel."""
    ref = (data_referencia or hoje()).normalize()
    datas = pd.to_datetime(datas, dayfirst=True, errors="coerce").dt.normalize()
    return (ref - datas).dt.days.clip(lower=0).astype("Int64")


def deve_cobrar(dias_uteis: pd.Series) -> pd.Series:
    """True para notas que já atingiram o 4º dia útil de atraso."""
    return dias_uteis.fillna(0) >= DIAS_UTEIS_MINIMO


def e_critico(dias_corridos: pd.Series) -> pd.Series:
    """True para notas com mais de DIAS_CORRIDOS_CRITICO dias corridos de atraso."""
    return dias_corridos.fillna(0) > DIAS_CORRIDOS_CRITICO
