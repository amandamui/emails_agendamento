"""Devolutiva semanal: compara a foto da semana passada com a de hoje (atrasadas/5_historico) e manda o
balanço para a gestão. As fotos são gravadas pelo disparador.py (rodar o disparador ANTES, na terça).

Uso:
    python devolutiva.py                         # SIMULAÇÃO: mostra o balanço e salva a prévia HTML em 3_envios
    python devolutiva.py --teste                 # envia o balanço para o EMAIL_TESTE do .env
    python devolutiva.py --enviar                # envia para EMAIL_DEVOLUTIVA do .env (separados por ';')
    python devolutiva.py 2026-11-30 [--teste | --enviar]   # usa a foto dessa data como "hoje"

Base da comparação = notas que o robô COBRARIA em cada foto (notas_a_cobrar na data da foto), pela chave
CNPJ do emitente + NF (o número da NF sozinho se repete entre fornecedores):
    Resolvidas = estavam na semana passada e não estão hoje
    Crônicas   = estavam na semana passada e continuam hoje
    Novas      = não estavam na semana passada e apareceram hoje
"""
import datetime as dt
import pathlib
import re
import sys
from datetime import datetime

import pandas as pd

from config import PASTA_ENVIOS, PASTA_HISTORICO, obrigatorio
from disparador import COLUNA_CNPJ_NOTA, cnpj_digitos, registrar_log
from regras import DIAS_CORRIDOS_CRITICO
from teste_email import (AZUL, CINZA_BORDA, CINZA_CABECALHO, EMAIL_DESTINATARIO, TEXTO, TEXTO_SUAVE, VERMELHO,
                         VERMELHO_CLARO, _bloco, _espaco, _fmt_data, _fmt_int, carregar_assinatura,
                         enviar_outlook, ler_excel, notas_a_cobrar)

INTERVALO_DIAS = 7          # "semana passada" = foto mais recente com pelo menos 7 dias a menos
VERDE = "#2e7d32"
VERDE_CLARO = "#e8f5e9"

# =========================================================
# 🗂️ FOTOS DO HISTÓRICO
# =========================================================


def fotos_historico() -> dict[dt.date, pathlib.Path]:
    """{data: arquivo} das fotos Atrasadas_AAAA-MM-DD.xlsx salvas pelo disparador."""
    fotos = {}
    for arquivo in PASTA_HISTORICO.glob("Atrasadas_*.xlsx"):
        achou = re.fullmatch(r"Atrasadas_(\d{4}-\d{2}-\d{2})\.xlsx", arquivo.name)
        if achou:
            fotos[dt.date.fromisoformat(achou.group(1))] = arquivo
    return fotos


def escolher_fotos(data_atual: dt.date | None = None) -> tuple[dt.date, dt.date] | None:
    """(semana passada, hoje). Hoje = foto da data pedida ou a mais recente; semana passada = a mais
    recente com 7+ dias a menos (se não houver, a imediatamente anterior, com aviso)."""
    fotos = fotos_historico()
    if data_atual is None:
        if not fotos:
            print(f"❌ Nenhuma foto em {PASTA_HISTORICO}.")
            print("⚠️  A devolutiva depende do disparador.py.")
            print("👉 Execute primeiro: python disparador.py")
            print("   Depois execute novamente: python devolutiva.py")
            return None
        data_atual = max(fotos)
        if data_atual < dt.date.today():
            print(f"⚠️  A foto mais recente é de {data_atual:%d/%m/%Y}; para comparar com hoje, "
                  "rode antes o disparador.py.")
            print("👉 Atualize os dados: python disparador.py --atualizar")
            print("   Depois execute novamente: python devolutiva.py")
    elif data_atual not in fotos:
        print(f"❌ Não há foto de {data_atual:%d/%m/%Y} em {PASTA_HISTORICO}.")
        datas_disponiveis = ", ".join(d.strftime("%d/%m/%Y") for d in sorted(fotos))
        if datas_disponiveis:
            print(f"📅 Fotos disponíveis: {datas_disponiveis}")
        print("👉 Gere as fotos na sequência correta:")
        print("   1. python disparador.py --atualizar")
        print("   2. python devolutiva.py")
        return None

    anteriores = [d for d in fotos if d < data_atual]
    if not anteriores:
        print(f"❌ Só existe a foto de {data_atual:%d/%m/%Y}; a devolutiva precisa da foto da semana anterior.")
        print("👉 Gere uma primeira foto em uma semana e outra na semana seguinte:")
        print("   1. Execute python disparador.py na primeira semana.")
        print("   2. Execute python disparador.py --atualizar na semana seguinte.")
        print("   3. Depois execute python devolutiva.py")
        return None
    semana_passada = [d for d in anteriores if (data_atual - d).days >= INTERVALO_DIAS]
    anterior = max(semana_passada or anteriores)
    if not semana_passada:
        print(f"⚠️  Ainda não há foto de {INTERVALO_DIAS}+ dias atrás: o balanço será PARCIAL "
              f"(compara com a foto mais próxima, de {anterior:%d/%m/%Y}).")
    return anterior, data_atual


def descrever_periodo(anterior: dt.date, atual: dt.date) -> dict[str, str]:
    """Textos do console/e-mail. Só diz "semana passada" se as fotos tiverem 7+ dias de distância;
    com menos, cita a data/quantidade de dias, para o leitor não achar que é o balanço da semana."""
    dias = (atual - anterior).days
    if dias >= INTERVALO_DIAS:
        return {"titulo": "Balanço Semanal", "foto_anterior": "semana passada",
                "kpi_cobradas": "Cobradas na semana passada", "kpi_novas": "Novas nesta semana",
                "cabecalho_resumo": "Resumo da semana", "periodo_anterior": f"Na última semana ({anterior:%d/%m/%Y})",
                "periodo_atual": f"Nesta semana ({atual:%d/%m/%Y})", "resolvidas": "Resolvidas na semana",
                "introducao": "na última semana"}
    intervalo = f"{dias} dia" + ("s" if dias > 1 else "")
    return {"titulo": "Balanço Parcial", "foto_anterior": f"{intervalo} antes",
            "kpi_cobradas": f"Cobradas em {anterior:%d/%m}", "kpi_novas": f"Novas desde {anterior:%d/%m}",
            "cabecalho_resumo": f"Resumo dos últimos {intervalo}" if dias > 1 else "Resumo do último dia",
            "periodo_anterior": f"Em {anterior:%d/%m/%Y}", "periodo_atual": f"Em {atual:%d/%m/%Y}",
            "resolvidas": "Resolvidas no período", "introducao": f"desde {anterior:%d/%m/%Y}"}


def _nf(valor) -> str | None:
    if pd.isna(valor):
        return None
    if isinstance(valor, float):  # Excel pode ler a NF como número
        valor = int(valor)
    return re.sub(r"\D", "", str(valor)).lstrip("0") or None


def notas_da_foto(data: dt.date) -> pd.DataFrame:
    """Notas que o robô cobraria no dia da foto, com 'Chave' (CNPJ|NF) e 'CNPJ'."""
    df = ler_excel(fotos_historico()[data])
    df = notas_a_cobrar(df, pd.Timestamp(data))  # Dias/Crítico recalculados na data da foto
    df["CNPJ"] = df[COLUNA_CNPJ_NOTA].map(cnpj_digitos)
    nf = df["NF"].map(_nf)
    sem_chave = df["CNPJ"].isna() | nf.isna()
    if sem_chave.any():
        print(f"⚠️  {sem_chave.sum()} nota(s) da foto de {data:%d/%m/%Y} sem CNPJ ou NF (fora da comparação).")
    df["Chave"] = df["CNPJ"] + "|" + nf
    return df[~sem_chave].drop_duplicates("Chave")


def comparar(anterior: pd.DataFrame, atual: pd.DataFrame) -> dict[str, pd.DataFrame]:
    continua = anterior["Chave"].isin(atual["Chave"])
    return {
        "anterior": anterior,
        "resolvidas": anterior[~continua],                           # dados da semana passada
        "cronicas": atual[atual["Chave"].isin(anterior["Chave"])],   # dados de hoje (Dias atualizado)
        "novas": atual[~atual["Chave"].isin(anterior["Chave"])],
    }


def resumo_por_fornecedor(r: dict[str, pd.DataFrame]) -> pd.DataFrame:
    """Uma linha por CNPJ: Cobradas (semana passada), Resolvidas, Continuam, Novas."""
    contagens = {nome: r[chave].groupby("CNPJ").size()
                 for nome, chave in (("Cobradas", "anterior"), ("Resolvidas", "resolvidas"),
                                     ("Continuam", "cronicas"), ("Novas", "novas"))}
    tabela = pd.DataFrame(contagens).fillna(0).astype(int)
    nomes = pd.concat([r["anterior"], r["novas"]]).groupby("CNPJ")["Razão Social"].first()
    tabela.insert(0, "Fornecedor", nomes.reindex(tabela.index).fillna("").astype(str).str.strip())
    return tabela.sort_values(["Continuam", "Cobradas", "Fornecedor"], ascending=[False, False, True])


# =========================================================
# 🎨 LAYOUT DO E-MAIL (mesmo padrão visual do teste_email.py)
# =========================================================

BORDA = f"border:1px solid {CINZA_BORDA};"
TH = (f"{BORDA} background-color:{CINZA_CABECALHO}; color:{TEXTO}; padding:8px 6px; font-size:11px; "
      "font-weight:bold; text-transform:uppercase; white-space:nowrap;")
TD = f"{BORDA} padding:7px 6px; font-size:12px; color:{TEXTO};"


def _tabela(cabecalho: list[tuple[str, str]], linhas: list[str]) -> str:
    ths = "".join(f'<th style="{TH} text-align:{al};">{t}</th>' for t, al in cabecalho)
    return (f'<table cellpadding="0" cellspacing="0" width="100%" style="border-collapse:collapse; '
            f'font-family:Segoe UI, Arial, sans-serif; margin:10px 0 20px 0;">'
            f"<thead><tr>{ths}</tr></thead><tbody>{''.join(linhas)}</tbody></table>")


def _kpi(valor: int, rotulo: str, cor: str, fundo: str) -> str:
    return (f'<td width="25%" style="background-color:{fundo}; border:1px solid {CINZA_BORDA}; '
            f'padding:12px 6px; text-align:center;">'
            f'<div style="font-size:24px; font-weight:bold; color:{cor};">{valor}</div>'
            f'<div style="font-size:11px; color:{TEXTO_SUAVE}; text-transform:uppercase;">{rotulo}</div></td>')


def tabela_resumo(resumo: pd.DataFrame) -> str:
    cab = [("Fornecedor", "left"), ("Cobradas", "center"), ("Resolvidas", "center"),
           ("Continuam", "center"), ("Novas", "center"), ("% Resolvido", "center")]
    linhas = []
    for _, r in resumo.iterrows():
        perc = f"{r['Resolvidas'] / r['Cobradas']:.0%}" if r["Cobradas"] else "-"
        continuam = f'<b style="color:{VERMELHO};">{r["Continuam"]}</b>' if r["Continuam"] else "0"
        valores = [r["Fornecedor"], r["Cobradas"], f'<span style="color:{VERDE};">{r["Resolvidas"]}</span>',
                   continuam, r["Novas"], perc]
        linhas.append("<tr>" + "".join(f'<td style="{TD} text-align:{al};">{v}</td>'
                                       for v, (_, al) in zip(valores, cab)) + "</tr>")
    return _tabela(cab, linhas)


def tabela_notas(df: pd.DataFrame, destacar_criticas: bool) -> str:
    """Notas agrupadas por fornecedor (linha de grupo com nome e quantidade), mais atrasadas primeiro."""
    cab = [("Filial", "left"), ("NF", "center"), ("Emissão", "center"), ("Agenda", "center"), ("Atraso", "center")]
    linhas = []
    grupos = sorted(df.groupby("CNPJ"), key=lambda g: (-len(g[1]), str(g[1]["Razão Social"].iloc[0])))
    for _, grupo in grupos:
        nome = str(grupo["Razão Social"].iloc[0]).strip().upper()
        linhas.append(f'<tr><td colspan="{len(cab)}" style="{BORDA} background-color:#e9eef4; padding:7px 8px; '
                      f'font-size:12px; font-weight:bold; color:{TEXTO};">{nome} '
                      f'<span style="font-weight:normal; color:{TEXTO_SUAVE};">({len(grupo)})</span></td></tr>')
        for _, r in grupo.sort_values("Dias", ascending=False).iterrows():
            critico = destacar_criticas and bool(r.get("Crítico", False))
            fundo = f"background-color:{VERMELHO_CLARO};" if critico else ""
            dias = "" if pd.isna(r.get("Dias")) else f"{int(r['Dias'])} dias"
            valores = [r.get("Filial", ""), _fmt_int(r.get("NF")), _fmt_data(r.get("Emissão")),
                       _fmt_data(r.get("Agenda")),
                       f'<b style="color:{VERMELHO};">{dias}</b>' if critico else dias]
            linhas.append("<tr>" + "".join(f'<td style="{TD} {fundo} text-align:{al}; white-space:nowrap;">{v}</td>'
                                           for v, (_, al) in zip(valores, cab)) + "</tr>")
    return _tabela(cab, linhas)


def _secao(titulo: str, cor: str, conteudo: str) -> str:
    return (_espaco(20) + f'<div style="font-size:15px; font-weight:bold; color:{cor}; '
            f'border-bottom:2px solid {cor}; padding-bottom:4px;">{titulo}</div>' + conteudo)


def montar_corpo_devolutiva(r: dict[str, pd.DataFrame], anterior: dt.date, atual: dt.date) -> str:
    t = descrever_periodo(anterior, atual)
    x, y, z, w = (len(r[k]) for k in ("anterior", "resolvidas", "cronicas", "novas"))
    s = lambda n, um, varios: um if n == 1 else varios  # noqa: E731
    perc = f" ({y / x:.0%})" if x else ""
    criticas = int(r["cronicas"]["Crítico"].sum())
    aviso_criticas = (f'<br><span style="color:{VERMELHO};">&#9888; <b>{criticas}</b> das pendentes já passam de '
                      f"{DIAS_CORRIDOS_CRITICO} dias corridos desde a emissão.</span>" if criticas else "")

    kpis = (f'<table cellpadding="0" cellspacing="6" width="100%"><tr>'
            + _kpi(x, t["kpi_cobradas"], TEXTO, "#ffffff")
            + _kpi(y, "Resolvidas", VERDE, VERDE_CLARO)
            + _kpi(z, "Continuam pendentes", VERMELHO, VERMELHO_CLARO)
            + _kpi(w, t["kpi_novas"], AZUL, "#eef5fb")
            + "</tr></table>")
    resumo = _bloco(
        f'<b style="font-size:14px;">{t["cabecalho_resumo"]}</b><br>'
        f"{t['periodo_anterior']}, o robô identificou <b>{x} {s(x, 'nota', 'notas')}</b> a cobrar. "
        f"Deste total, <b style=\"color:{VERDE};\">{y} {s(y, 'foi resolvida', 'foram resolvidas')}</b>{perc} e "
        f"<b style=\"color:{VERMELHO};\">{z} {s(z, 'continua pendente', 'continuam pendentes')}</b>. "
        f"{t['periodo_atual']} {s(w, 'surgiu', 'surgiram')} <b>{w} {s(w, 'nota nova', 'notas novas')}</b>."
        f"{aviso_criticas}",
        fundo="#eef5fb", borda_esquerda=AZUL,
    )
    secoes = _secao("Balanço por fornecedor", AZUL, tabela_resumo(resumo_por_fornecedor(r)))
    if z:
        secoes += _secao(f"Continuam pendentes ({z})", VERMELHO, tabela_notas(r["cronicas"], destacar_criticas=True))
    if y:
        secoes += _secao(f"{t['resolvidas']} ({y})", VERDE, tabela_notas(r["resolvidas"], destacar_criticas=False))

    return f"""
<html>
<body style="margin:0; padding:20px; background-color:#f4f5f7; font-family:Segoe UI, Arial, sans-serif; font-size:14px; color:{TEXTO}; line-height:1.5;">
<table cellpadding="0" cellspacing="0" width="100%" style="max-width:900px; margin:0 auto; background-color:#ffffff; border:1px solid #e3e5e8;">
  <tr><td style="background-color:{AZUL}; padding:18px 24px; text-align:center; color:#ffffff;">
    <div style="font-size:20px; font-weight:bold;">{t["titulo"]}: Agendamentos Atrasados</div>
    <div style="font-size:14px;">{anterior:%d/%m/%Y} &rarr; {atual:%d/%m/%Y}</div>
  </td></tr>
  <tr><td style="padding:24px;">
    <p style="margin:0 0 12px 0;">Olá,</p>
    <p style="margin:0;">Segue o acompanhamento das notas em No-Show cobradas dos fornecedores {t["introducao"]}.</p>
    {_espaco(12)}
    {kpis}
    {resumo}
    {secoes}
    {_espaco(24)}
    <p style="margin:0;">Atenciosamente,</p>
    {carregar_assinatura()}
  </td></tr>
  <tr><td style="background-color:#f8f9fa; border-top:1px solid #e3e5e8; padding:12px; text-align:center; font-size:11px; color:{TEXTO_SUAVE};">
    Este é um e-mail automático gerado pelo sistema de Agendamentos da Brautoparts.
  </td></tr>
</table>
</body>
</html>"""


# =========================================================
# 🏁 EXECUÇÃO
# =========================================================


def run(data_atual: dt.date | None = None, modo: str = "simulacao"):
    escolha = escolher_fotos(data_atual)
    if escolha is None:
        return
    anterior, atual = escolha
    t = descrever_periodo(anterior, atual)
    rotulo_atual = "hoje" if atual == dt.date.today() else "mais recente"
    print(f"📂 Comparando as fotos de {anterior:%d/%m/%Y} ({t['foto_anterior']}) e "
          f"{atual:%d/%m/%Y} ({rotulo_atual})…")
    r = comparar(notas_da_foto(anterior), notas_da_foto(atual))
    print(f"   📊 Cobradas: {len(r['anterior'])} | ✅ Resolvidas: {len(r['resolvidas'])} | "
          f"⚠️  Continuam: {len(r['cronicas'])} | 🆕 Novas: {len(r['novas'])}")

    corpo = montar_corpo_devolutiva(r, anterior, atual)
    assunto = f"{t['titulo']}: Agendamentos Atrasados – {atual:%d/%m/%Y}"
    if modo == "simulacao":
        previa = PASTA_ENVIOS / f"Devolutiva_{atual:%Y-%m-%d}.html"
        previa.write_text(corpo, encoding="utf-8")
        print("\n🧪 SIMULAÇÃO concluída (nada foi enviado).")
        print(f"📬  Destinatário do teste: {EMAIL_DESTINATARIO}")
        print("ℹ️  Enviar teste: python devolutiva.py --teste")
        print("📧 Envio real: python devolutiva.py --enviar")
        return

    teste = modo == "teste"
    destino = EMAIL_DESTINATARIO if teste else obrigatorio("EMAIL_DEVOLUTIVA")
    print(f"📧 Enviando o balanço para {destino}…")
    ok = enviar_outlook(destino, ("[TESTE] " if teste else "") + assunto, corpo)
    print("✅ Balanço enviado!" if ok else "❌ Falha ao enviar o balanço (Outlook aberto?).")
    registrar_log([{
        "Data/Hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Fornecedor": "DEVOLUTIVA SEMANAL",
        "E-mail": destino,
        "Qtd Notas": len(r["anterior"]),
        "Status": ("Sucesso" if ok else "Erro") + (" (teste)" if teste else ""),
    }])


def main():
    flags = {a for a in sys.argv[1:] if a.startswith("--")}
    if flags - {"--teste", "--enviar"} or {"--teste", "--enviar"} <= flags:
        raise SystemExit("Uso: python devolutiva.py [AAAA-MM-DD] [--teste | --enviar]")
    datas = [a for a in sys.argv[1:] if not a.startswith("--")]
    if len(datas) > 1:
        raise SystemExit("❌ Passe apenas uma data opcional.\n"
                         "Uso: python devolutiva.py [AAAA-MM-DD] [--teste | --enviar]")
    try:
        data_atual = dt.date.fromisoformat(datas[0]) if datas else None
    except ValueError:
        raise SystemExit("❌ Data inválida. Use o formato AAAA-MM-DD, por exemplo: 2026-11-30.\n"
                         "Uso: python devolutiva.py [AAAA-MM-DD] [--teste | --enviar]")
    run(data_atual,
        modo="teste" if "--teste" in flags else "enviar" if "--enviar" in flags else "simulacao")


if __name__ == "__main__":
    main()
