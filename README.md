# Automação Modify Order Line (Outbound Planner)

Script em Python (Playwright) que lê uma planilha de pedidos e itens e repete no Outbound Planner do WMS o fluxo:

pesquisar pedido → abrir pedido → filtrar `item:<código>` → selecionar item → **Actions → Modify Order Line** → Processing → ligar **Parcial** → Save → Reason Code **999 - Sem Interface** + comentário **RP** → OK → Voltar.

## Instalação (uma vez)

```powershell
py -m pip install -r requirements.txt
```

O Playwright usa o seu próprio Chrome, então não precisa rodar `playwright install`.

## Uso

1. Feche o Chrome e abra com a porta de depuração:
   ```powershell
   taskkill /F /IM chrome.exe
   & "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\chrome-automacao"
   ```
2. Faça login no portal e deixe aberta a tela do Outbound Planner.
3. Configure o endereço do WMS: edite `TRECHO_URL_TELA` no topo do script ou defina a variável de ambiente antes de rodar:
   ```powershell
   $env:WMS_HOST = "servidor-wms.suaempresa.com"
   ```
4. Copie `pedidos_modelo.xlsx` para `pedidos.xlsx` e preencha as colunas **PEDIDO** e **ITEM**. O item pode ser escrito com ou sem o prefixo `item:`.
5. Rode:
   ```powershell
   py modify_order_line.py            # processa a planilha inteira
   py modify_order_line.py --teste    # só a 1ª linha, pausa antes do OK (ENTER no terminal)
   py modify_order_line.py --botoes   # lista os botões visíveis (diagnóstico)
   py modify_order_line.py --parcial  # testa só o switch do Parcial na tela aberta
   ```

## Saídas

- `log_AAAAMMDD_HHMMSS.csv`: OK ou ERRO por linha, com a etapa em que parou.
- `erros/<pedido>_<item>.png`: print da tela quando uma linha falha.

## Configuração

As constantes ficam no topo de `modify_order_line.py`: nome da planilha, Reason Code, comentário, rótulo do switch do Parcial, textos e id do botão Voltar, tempos de espera.
