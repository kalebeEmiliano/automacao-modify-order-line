"""
Automação: Modify Order Line (Outbound Planner - WMS)
------------------------------------------------------
Para cada linha da planilha (PEDIDO, ITEM):
  pesquisa pedido -> abre pedido -> filtra item -> seleciona item -> Action
  -> Modify Order Line -> Processing -> flag Parcial -> Save
  -> Reason Code 999 + Comentário RP -> OK -> Voltar

COMO USAR
1) Instale (uma vez):
       pip install playwright pandas openpyxl
2) Feche o Chrome e abra-o com a porta de depuração (cmd / atalho):
       "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\\chrome-automacao"
   (no Edge: msedge.exe com os mesmos parâmetros)
   Nesse Chrome, faça login no portal e deixe aberta a tela do Outbound Planner.
3) Rode:
       python modify_order_line.py              -> processa a planilha toda
       python modify_order_line.py --teste      -> só a 1ª linha e pede confirmação antes do OK
       python modify_order_line.py --botoes     -> lista os botões visíveis (para ajustar textos)
       python modify_order_line.py --parcial    -> só liga o switch do Parcial na tela aberta (teste)

Resultado de cada linha vai para log_AAAAMMDD_HHMMSS.csv; erros geram print em ./erros/.
"""
import csv
import re
import sys
import time
from datetime import datetime
from pathlib import Path

import os

import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

VERSAO = "12.1 (endereço configurável)"

# ===================== CONFIGURAÇÃO =====================
PLANILHA = "pedidos.xlsx"      # .xlsx ou .csv
COL_PEDIDO = "PEDIDO"
COL_ITEM = "ITEM"
ITEM_DIGITOS = 0               # 0 = usa o item exatamente como está na planilha
PREFIXO_ITEM = "item:"         # texto digitado antes do item no campo de filtro

CDP_URL = "http://localhost:9222"
# Trecho do endereço da tela do Outbound Planner (ex.: "servidor-wms.suaempresa.com").
# Pode ser definido aqui ou na variável de ambiente WMS_HOST.
TRECHO_URL_TELA = os.environ.get("WMS_HOST", "COLOQUE_AQUI_O_ENDERECO_DO_WMS")

REASON_CODE = "999"
COMENTARIO = "RP"
LABEL_PARCIAL = ["Partial", "Parcial"]   # texto do rótulo ao lado do switch do parcial

TIMEOUT_MS = 15000
PAUSA = 0.8                    # respiro entre ações (o ExtJS precisa)

# Textos possíveis dos botões que só têm ícone (texto, tooltip ou aria-label).
# Se algum não for encontrado, rode com --botoes e ajuste aqui.
BTN_ACTION = ["Actions", "Action", "Ações"]
BTN_SAVE = ["Save", "Salvar"]
BTN_OK = ["OK", "Ok"]
BTN_VOLTAR = ["Back", "Voltar", "Return", "Previous"]
VOLTAR_IDS = ["button-1601"]   # id do botão Voltar (só ícone). Se mudar, atualize aqui
# ========================================================


def pausa(mult=1.0):
    time.sleep(PAUSA * mult)


def achar(page, seletor, texto=None, exato=True, timeout=TIMEOUT_MS):
    """Procura o elemento visível em todos os frames da página.
    Os IDs do ExtJS (ext-gen..., button-4197...) mudam a cada sessão,
    por isso a busca é por name, classe e texto."""
    fim = time.time() + timeout / 1000
    while time.time() < fim:
        for fr in page.frames:
            try:
                loc = fr.locator(seletor)
                if texto is not None:
                    padrao = (re.compile(rf"^\s*{re.escape(texto)}\s*$") if exato
                              else re.compile(re.escape(texto), re.I))
                    loc = loc.filter(has_text=padrao)
                for i in range(loc.count() - 1, -1, -1):   # o mais recente (janela aberta por último)
                    el = loc.nth(i)
                    if el.is_visible():
                        return el
            except Exception:
                pass
        time.sleep(0.25)
    raise PWTimeout(f"Não encontrei: {seletor} {texto or ''}".strip())


def clicar_botao(page, textos, timeout=TIMEOUT_MS):
    sels = []
    for t in textos:
        sels += [
            f'.x-btn-inner:text-is("{t}")',
            f'.x-btn:has(.x-btn-inner:text-is("{t}"))',
            f'.x-btn[data-qtip="{t}" i]',
            f'.x-btn[aria-label="{t}" i]',
            f'.x-btn[title="{t}" i]',
        ]
    achar(page, ", ".join(sels), timeout=timeout).click()
    pausa()


def digitar(campo, valor, enter=True):
    campo.click()
    campo.fill("")
    campo.fill(str(valor))
    if enter:
        campo.press("Enter")
    pausa()


JS_TOGGLES = """() => [...document.querySelectorAll('div.toggle')]
    .filter(el => el.offsetParent && el.querySelector('input.toggle_switch'))
    .map(el => {
        let lbl = '', p = el;
        for (let i = 0; i < 6 && p && !lbl; i++) {          // sobe até achar um rótulo
            p = p.parentElement;
            const l = p && p.querySelector('label, .x-form-item-label');
            if (l && l.innerText.trim()) lbl = l.innerText.trim();
        }
        const v = el.querySelector('.toggle_value');
        return {id: el.id, label: lbl, valor: v ? v.innerText.trim() : ''};
    })"""


def listar_toggles(page):
    """Lista os switches visíveis: id do div.toggle, rótulo e valor (Yes/No)."""
    achados = []
    for fr in page.frames:
        try:
            for t in fr.evaluate(JS_TOGGLES):
                t["frame"] = fr
                achados.append(t)
        except Exception:
            pass
    return achados


def _valor(fr, toggle_id):
    try:
        return fr.locator(f'[id="{toggle_id}"] .toggle_value').inner_text().strip()
    except Exception:
        return ""


def ajustar_parcial(page):
    """Liga o switch do Parcial (div.toggle > input.toggle_switch + .toggle_value Yes/No)."""
    toggles = listar_toggles(page)
    if not toggles:
        raise PWTimeout("Nenhum switch (div.toggle) visível na tela")

    alvo = None
    for t in toggles:                                    # 1) pelo rótulo
        if any(x.lower() in t["label"].lower() for x in LABEL_PARCIAL):
            alvo = t
            break
    if alvo is None and len(toggles) == 1:               # 2) só existe um
        alvo = toggles[0]
    if alvo is None:
        resumo = [(t["id"], t["label"], t["valor"]) for t in toggles]
        raise PWTimeout(f"Vários switches e nenhum com rótulo {LABEL_PARCIAL}: {resumo}")

    fr, tid = alvo["frame"], alvo["id"]
    cmp_id = tid.replace("-toggleEl", "")               # rpToggle-3008
    if _valor(fr, tid).lower() == "yes":
        print("(Parcial já estava Yes) ", end="", flush=True)
        return "ja_estava"

    caixa = fr.locator(f'[id="{tid}"]')
    botao = caixa.locator("input.toggle_switch")

    def arrastar():
        bb = botao.bounding_box()
        x, y = bb["x"] + bb["width"] / 2, bb["y"] + bb["height"] / 2
        page.mouse.move(x, y)
        page.mouse.down()
        page.mouse.move(x + 45, y, steps=8)
        page.mouse.up()

    tentativas = [
        ("clique no switch", lambda: caixa.click()),
        ("clique no botão", lambda: botao.click(force=True)),
        ("clique no texto", lambda: caixa.locator(".toggle_value").click(force=True)),
        ("arrastar", arrastar),
        ("eventos JS", lambda: botao.evaluate("""el => { for (const t of ['mousedown','mouseup','click'])
                el.dispatchEvent(new MouseEvent(t, {bubbles:true, cancelable:true, view:window})); }""")),
        ("ExtJS setValue", lambda: fr.evaluate("""id => { const c = window.Ext && Ext.getCmp(id);
                if (!c) return;
                if (c.setValue) c.setValue(true);
                else if (c.toggle) c.toggle(); }""", cmp_id)),
    ]
    feitas = []
    for nome, tentar in tentativas:
        try:
            tentar()
        except Exception as e:
            feitas.append(f"{nome}: erro")
            continue
        pausa()
        if _valor(fr, tid).lower() == "yes":
            print(f"(Parcial ligado via {nome}) ", end="", flush=True)
            return "ligado"
        feitas.append(f"{nome}: continuou {_valor(fr, tid)}")
    raise PWTimeout(f"Switch {cmp_id} não ligou ({'; '.join(feitas)})")


def preencher_reason(page):
    rc = achar(page, 'input[name="reasonCode"]')
    rc.click()
    rc.fill("")
    rc.press_sequentially(REASON_CODE, delay=80)
    pausa()
    try:   # escolhe "999 - Sem Interface" na lista
        achar(page, ".x-boundlist-item", texto=REASON_CODE, exato=False, timeout=4000).click()
    except PWTimeout:
        # lista não abriu: abre pela setinha do combobox e tenta de novo
        try:
            rc.locator("xpath=ancestor::table[1]//div[contains(@class,'x-form-trigger')]").first.click()
            pausa()
            achar(page, ".x-boundlist-item", texto=REASON_CODE, exato=False, timeout=4000).click()
        except Exception:
            rc.press("Enter")
    pausa()
    com = achar(page, 'textarea[name="reasonComment"]')
    com.click()
    com.fill("")
    com.press_sequentially(COMENTARIO, delay=80)   # digitação real: o ExtJS só valida com tecla
    com.press("Tab")                                # sai do campo para liberar o OK
    pausa()


JS_ACHAR_OK = r"""() => {
    const ta = [...document.querySelectorAll('textarea[name="reasonComment"]')].filter(e => e.offsetParent).pop();
    if (!ta) return null;
    let p = ta.parentElement;
    while (p) {                                     // sobe até o bloco que tem o botão OK
        const bts = [...p.querySelectorAll('a.x-btn')].filter(b => b.offsetParent);
        const ok = bts.find(b => /^\s*OK\s*$/i.test(b.innerText || '')) ||
                   bts.find(b => /ok/i.test((b.getAttribute('data-qtip') || '') + (b.getAttribute('aria-label') || '')));
        if (ok) return {id: ok.id, desabilitado: ok.classList.contains('x-disabled') || ok.classList.contains('x-btn-disabled')};
        p = p.parentElement;
    }
    return null;
}"""


JS_ACIONAR_BOTAO = r"""id => {
    const c = window.Ext && Ext.getCmp(id);
    if (!c) return 'sem componente ExtJS';
    if (c.disabled) return 'desabilitado';
    const ev = Ext.EventObject;
    if (c.fireHandler) { c.fireHandler(ev); return 'fireHandler'; }
    if (c.handler) { c.handler.call(c.scope || c, c, ev); return 'handler'; }
    c.fireEvent('click', c, ev);
    return 'fireEvent';
}"""

JS_EVENTOS_MOUSE = r"""el => {
    const alvo = el.querySelector('.x-btn-button, .x-btn-inner') || el;
    for (const t of ['mouseover', 'mousedown', 'mouseup', 'click'])
        alvo.dispatchEvent(new MouseEvent(t, {bubbles: true, cancelable: true, view: window, button: 0}));
}"""

JS_MENSAGENS = r"""() => {
    // janelas/mensagens visíveis que NÃO são a do motivo, com botão OK/Yes/Sim
    const r = [];
    for (const w of document.querySelectorAll('.x-window, .x-message-box')) {
        if (!w.offsetParent || w.querySelector('textarea[name="reasonComment"]')) continue;
        const b = [...w.querySelectorAll('a.x-btn')].filter(x => x.offsetParent)
            .find(x => /^\s*(OK|Yes|Sim)\s*$/i.test(x.innerText || ''));
        if (b) r.push({janela: w.id, botao: b.id, texto: (w.innerText || '').trim().slice(0, 120)});
    }
    return r;
}"""


def fechar_mensagens(page, max_vezes=3):
    """Fecha avisos de confirmação que aparecem depois do OK (ex.: 'linha modificada')."""
    for _ in range(max_vezes):
        achou = False
        for fr in page.frames:
            try:
                msgs = fr.evaluate(JS_MENSAGENS)
            except Exception:
                continue
            for m in msgs:
                achou = True
                print(f"(aviso: {m['texto'][:60]!r}) ", end="", flush=True)
                btn = fr.locator(f'[id="{m["botao"]}"]')
                try:
                    fr.evaluate(JS_ACIONAR_BOTAO, m["botao"])
                    pausa(0.5)
                    if btn.is_visible():
                        btn.focus()
                        page.keyboard.press("Enter")
                        pausa(0.5)
                    if btn.is_visible():
                        btn.click(force=True)
                except Exception:
                    pass
                pausa()
        if not achou:
            return


def clicar_ok(page):
    """Acha o OK subindo a partir do campo de comentário, espera ficar habilitado,
    clica e confere se a janela do motivo fechou."""
    fim = time.time() + TIMEOUT_MS / 1000
    info = fr_ok = None
    while time.time() < fim:
        for fr in page.frames:
            try:
                r = fr.evaluate(JS_ACHAR_OK)
            except Exception:
                r = None
            if r:
                info, fr_ok = r, fr
                if not r["desabilitado"]:
                    break
        if info and not info["desabilitado"]:
            break
        time.sleep(0.3)
    if not info:
        raise PWTimeout("Não achei o botão OK perto do campo de comentário")
    if info["desabilitado"]:
        raise PWTimeout(f"OK ({info['id']}) está desabilitado: Reason Code ou comentário não foram aceitos")

    btn = fr_ok.locator(f'[id="{info["id"]}"]')
    comentario = fr_ok.locator('textarea[name="reasonComment"]:visible').first

    def fechou(timeout):
        try:
            comentario.wait_for(state="hidden", timeout=timeout)
            return True
        except PWTimeout:
            return False

    bid = info["id"]
    tentativas = [
        ("ExtJS handler", lambda: fr_ok.evaluate(JS_ACIONAR_BOTAO, bid)),
        ("foco + ENTER", lambda: (btn.focus(), page.keyboard.press("Enter"))),
        ("clique", lambda: btn.click(force=True)),
        ("eventos de mouse", lambda: btn.evaluate(JS_EVENTOS_MOUSE)),
    ]
    feitas = []
    for nome, tentar in tentativas:
        if not comentario.is_visible():
            break
        try:
            r = tentar()
        except Exception as e:
            feitas.append(f"{nome}: erro")
            continue
        if fechou(4000):
            print(f"(OK via {nome}) ", end="", flush=True)
            break
        feitas.append(f"{nome}: não fechou{' ' + str(r) if isinstance(r, str) else ''}")
    if comentario.is_visible():
        raise PWTimeout(f"OK ({bid}) não fechou a janela: {'; '.join(feitas)}")
    pausa(1.5)
    fechar_mensagens(page)


def clicar_voltar(page):
    """Voltar é um botão só com ícone (button-1601). Tenta pelo id e depois por tooltip."""
    for bid in VOLTAR_IDS:
        for fr in page.frames:
            loc = fr.locator(f'a.x-btn[id="{bid}"]')
            if loc.count() and loc.first.is_visible():
                loc.first.click()
                pausa(1.5)
                return
    clicar_botao(page, BTN_VOLTAR)
    pausa(1.5)


def processar(page, pedido, item, teste=False):
    etapa = ""
    try:
        etapa = "Pesquisar pedido"
        digitar(achar(page, 'input[name="editField"]'), pedido)
        etapa = "Abrir pedido"
        achar(page, "span.rpux-link-grid-column-link", texto=pedido).click()
        pausa(1.5)

        etapa = "Filtrar item"
        filtro = achar(page, 'input[id^="rpFilterComboBox-"][id$="-inputEl"]')
        filtro.click()
        filtro.fill("")
        filtro.press_sequentially(f"{PREFIXO_ITEM}{item}", delay=60)
        pausa(0.5)
        filtro.press("Enter")
        pausa(1.5)
        etapa = "Selecionar item"
        selecionar_linha_item(page, item)

        etapa = "Actions"
        clicar_botao(page, BTN_ACTION)
        etapa = "Modify Order Line"
        achar(page, ".x-menu-item-text", texto="Modify Order Line").click()
        pausa(1.5)

        etapa = "Processing"
        achar(page, ".wm-storecount-text", texto="Processing").click()
        pausa()
        etapa = "Ligar Parcial"
        estado_parcial = ajustar_parcial(page)

        etapa = "Save"
        clicar_botao(page, BTN_SAVE)
        pausa()
        etapa = ("Reason Code (Parcial já estava Yes, talvez nada mudou e a janela não abriu)"
                 if estado_parcial == "ja_estava" else "Reason Code / Comentário")
        preencher_reason(page)
        if teste:
            input("\n>> MODO TESTE PAUSADO. Confira a tela e aperte ENTER AQUI NO POWERSHELL "
              "(não no navegador) para o script clicar em OK... ")
        etapa = "OK"
        clicar_ok(page)

        etapa = "Voltar"
        fechar_mensagens(page)
        clicar_voltar(page)
    except Exception as e:
        raise RuntimeError(f"[{etapa}] {str(e).splitlines()[0]}") from e


def selecionar_linha_item(page, item):
    """Marca o checkbox (x-grid-row-checker) da linha filtrada.
    Prefere a linha que contém o código do item; senão, a primeira linha visível."""
    fim = time.time() + TIMEOUT_MS / 1000
    while time.time() < fim:
        for fr in page.frames:
            try:
                linhas = fr.locator(".x-grid-row:visible")
                n = linhas.count()
                if not n:
                    continue
                alvo = None
                for i in range(n):
                    if item.lower() in (linhas.nth(i).inner_text() or "").lower():
                        alvo = linhas.nth(i)
                        break
                if alvo is None:
                    alvo = linhas.first
                checker = alvo.locator(".x-grid-row-checker")
                if checker.count():
                    checker.first.click()
                else:
                    alvo.locator(".x-grid-cell-inner").first.click()
                pausa()
                return
            except Exception:
                pass
        time.sleep(0.3)
    raise PWTimeout(f"Nenhuma linha apareceu depois de filtrar {PREFIXO_ITEM}{item}")


def recuperar(page):
    """Tenta fechar janelas abertas e voltar à pesquisa depois de um erro."""
    for _ in range(3):
        page.keyboard.press("Escape")
        time.sleep(0.3)
    try:
        clicar_voltar(page)
    except Exception:
        pass


def ler_planilha():
    caminho = Path(PLANILHA)
    if caminho.suffix.lower() == ".csv":
        df = pd.read_csv(caminho, dtype=str, sep=None, engine="python")
    else:
        df = pd.read_excel(caminho, dtype=str)
    df.columns = [c.strip().upper() for c in df.columns]
    linhas = []
    for _, r in df.iterrows():
        ped = str(r.get(COL_PEDIDO, "")).strip()
        it = str(r.get(COL_ITEM, "")).strip()
        if not ped or ped.lower() == "nan" or not it or it.lower() == "nan":
            continue
        ped = re.sub(r"\.0$", "", ped)
        it = re.sub(r"\.0$", "", it)
        if it.lower().startswith(PREFIXO_ITEM.lower()):   # aceita "item:PR..." ou só "PR..."
            it = it[len(PREFIXO_ITEM):].strip()
        if ITEM_DIGITOS and it.isdigit():
            it = it.zfill(ITEM_DIGITOS)
        linhas.append((ped, it))
    return linhas


def pegar_aba(browser):
    if TRECHO_URL_TELA.startswith("COLOQUE_AQUI"):
        sys.exit("Configure o endereço do WMS: edite TRECHO_URL_TELA no topo do script "
                 "ou defina a variável de ambiente WMS_HOST.")
    for ctx in browser.contexts:
        for pg in ctx.pages:
            if TRECHO_URL_TELA in pg.url:
                return pg
    sys.exit("Não achei a aba do portal. Abra a tela do Outbound Planner no Chrome com a porta 9222.")


def listar_botoes(page):
    js = """() => [...document.querySelectorAll('.x-btn')].filter(b => b.offsetParent)
        .map(b => ({id: b.id, texto: (b.innerText || '').trim(),
                    qtip: b.getAttribute('data-qtip') || '', aria: b.getAttribute('aria-label') || '',
                    title: b.getAttribute('title') || ''}))"""
    for fr in page.frames:
        try:
            for b in fr.evaluate(js):
                print(b)
        except Exception:
            pass


def main():
    # trabalha sempre na pasta do script (planilha, log e erros ficam lá),
    # não importa de onde foi executado (VS Code, PowerShell, atalho...)
    os.chdir(Path(__file__).resolve().parent)
    print(f"modify_order_line.py - VERSAO {VERSAO}")
    print(f"Pasta: {Path.cwd()}")
    if not Path(PLANILHA).exists():
        sys.exit(f"Não achei '{PLANILHA}' nesta pasta. Coloque a planilha junto do script "
                 f"(ou mude PLANILHA no topo do arquivo).")
    teste = "--teste" in sys.argv
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(CDP_URL)
        page = pegar_aba(browser)
        page.bring_to_front()

        if "--botoes" in sys.argv:
            listar_botoes(page)
            return
        if "--parcial" in sys.argv:
            for t in listar_toggles(page):
                print({k: v for k, v in t.items() if k != "frame"})
            ajustar_parcial(page)
            print("Parcial ligado.")
            return

        linhas = ler_planilha()
        if teste:
            linhas = linhas[:1]
        print(f"{len(linhas)} linha(s) para processar.")

        Path("erros").mkdir(exist_ok=True)
        log_nome = f"log_{datetime.now():%Y%m%d_%H%M%S}.csv"
        with open(log_nome, "w", newline="", encoding="utf-8-sig") as f:
            log = csv.writer(f, delimiter=";")
            log.writerow(["DATA_HORA", "PEDIDO", "ITEM", "STATUS", "DETALHE"])
            ok = erro = 0
            for n, (pedido, item) in enumerate(linhas, 1):
                print(f"[{n}/{len(linhas)}] Pedido {pedido} | Item {item} ... ", end="", flush=True)
                try:
                    processar(page, pedido, item, teste)
                    log.writerow([f"{datetime.now():%d/%m/%Y %H:%M:%S}", pedido, item, "OK", ""])
                    ok += 1
                    print("OK")
                except Exception as e:
                    msg = str(e).splitlines()[0][:200]
                    try:
                        page.screenshot(path=f"erros/{pedido}_{item}.png")
                    except Exception:
                        pass
                    log.writerow([f"{datetime.now():%d/%m/%Y %H:%M:%S}", pedido, item, "ERRO", msg])
                    erro += 1
                    print(f"ERRO: {msg}")
                    recuperar(page)
                f.flush()
        print(f"\nFim. OK: {ok} | Erros: {erro} | Log: {log_nome}")


if __name__ == "__main__":
    main()
