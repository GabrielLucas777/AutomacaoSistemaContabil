import json
import random

from playwright.sync_api import sync_playwright  

# --- PARTE 1: CONFIGURAÇÃO DE CAMINHO ---
# Caminho do arquivo JSON com os produtos e preços de referência
ARQUIVO_JSON = r"D:\Desenvolvimento\automacao_ecommerce\precos.json"

# --- PARTE 2: FUNÇÕES DE APOIO ---

def tratar_preco(texto_bruto):
    
    preco_limpo = texto_bruto.replace('R$', '').replace('.', '').replace(',', '.').strip()
    return float(preco_limpo)

def carregar_dados_json():
    """ 
    LER O DOCUMENTO (JSON)
        Carrega o arquivo JSON e retorna os dados como um dicionário.
    """
    try:
        with open(ARQUIVO_JSON, 'r', encoding='utf-8-sig') as f:
            return json.load(f)
    except Exception as e:
        print(f" Erro ao ler o JSON: {e}")
        return {}

# --- PARTE 3: AÇÃO WEB ---

def verificar_produto_no_site(page, produto):
    """ 
    VERIFICA O PRODUTO NO SITE
        Acessa a página do produto, verifica disponibilidade e retorna o preço atual.
    """
    print(f"\n--- Verificando: {produto['nome']} ---")
    print(f"Alvo atual: R$ {produto['preco_referencia']}")
    
    # Navegação
    page.goto(produto['url'], wait_until='domcontentloaded')
    
    # Espera aleatória para simular comportamento humano
    page.wait_for_timeout(random.randint(3000, 5000)) 

    # Lógica do seletor de disponibilidade
    aviso_indisponivel = page.get_by_role("heading", name=" Produto Indisponível")

    if not aviso_indisponivel.is_visible():
        try:
            # Localizando preço 
            elemento_preco = page.locator('#valVista').first
            elemento_preco.wait_for(state="visible", timeout=10000)
            
            # Tratando o preço para formato float
            preco_atual = tratar_preco(elemento_preco.inner_text())
            return preco_atual

        except Exception:
            print(f" Preço não localizado (pode ser Captcha).")
            return None
    else:
        print(f" STATUS: Produto Indisponível no momento.")
        return None

# --- PARTE 4: FLUXO PRINCIPAL ---

def executar_bot():
    with sync_playwright() as p: 
        # Lançamento com camuflagem anti-bot 
        browser = p.chromium.launch(headless=False, args=["--disable-blink-features=AutomationControlled"]) 
        context = browser.new_context(
            viewport={"width": 1366, "height": 768},
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        # Carrega os dados uma vez para começar a rodar
        dados_json = carregar_dados_json()

        # 3. ITERAR PELOS PRODUTOS 
        for item in dados_json:
            produto = dados_json[item]
            
            # Chama a função que vai no site e traz o preço real
            preco_site = verificar_produto_no_site(page, produto)

            # Se conseguimos pegar o preço, fazemos a comparação
            if preco_site is not None:
                print(f" STATUS: Disponível | Preço Site: R$ {preco_site:.2f}")

                if preco_site <= produto['preco_referencia']:
                    print(" ALERTA: O PREÇO ATINGIU O SEU ALVO! RECOMENDO COMPRAR AGORA!")
                else:
                    diferenca = preco_site - produto['preco_referencia']
                    print(f" Ainda caro. Diferença: R$ {diferenca:.2f}. Não recomendo comprar.")

            # Pausa entre produtos para evitar bloqueios
            page.wait_for_timeout(2000)

        print("\n" + "="*40)
        print("Automação concluída.")
        input("Pressione Enter para sair...")
        browser.close()

# Inicia o programa
if __name__ == "__main__":
    executar_bot()