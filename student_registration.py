import os  # manipular caminhos de arquivos
import sys  # ler argumentos passados na linha de comando
import threading  # criar o lock para proteger o acesso ao Excel entre navegadores
from concurrent.futures import ThreadPoolExecutor  # rodar os navegadores em paralelo
import pandas as pd  # ler e salvar o arquivo Excel
from dotenv import load_dotenv  # carregar links e senha do arquivo .env
from selenium import webdriver  # controlar o navegador Chrome
from selenium.webdriver.common.by import By  # localizar elementos por XPATH
from selenium.webdriver.support.ui import WebDriverWait  # esperar um elemento aparecer
from selenium.webdriver.support import expected_conditions as EC  # condições de espera
from selenium.webdriver.common.action_chains import ActionChains  # ações do mouse
import time  # pausas entre passos

# Diretório onde o script está, para montar o caminho do Excel sempre do jeito certo
current_dir = os.path.dirname(os.path.abspath(__file__))

# Carrega as variáveis do arquivo .env (que fica nesta mesma pasta)
load_dotenv(os.path.join(current_dir, ".env"))

# Diretório onde ficam o Excel e os arquivos gerados (configurável no .env)
DIRETORIO_DADOS = os.getenv("DIR_DATA", current_dir)
file_path = os.path.join(DIRETORIO_DADOS, "dados.xlsx")

# ADICIONE OS LINKS E A SENHA NO ARQUIVO .env (nunca suba o .env para o Git!)
URL_CADASTRAR = os.getenv("URL_CADASTRAR")
URL_LOGIN = os.getenv("URL_LOGIN")
PASS_FIXED = os.getenv("PASS_FIXED")

# Se o usuário rodou "python register_completo.py 4", usa 4 navegadores.
# Se não passou nada, pergunta e o padrão é 1 (um de cada vez).
if len(sys.argv) > 1:
    NUM_NAVEGADORES = int(sys.argv[1])
else:
    NUM_NAVEGADORES = int(input("Quantos navegadores em paralelo? (padrão 1): ") or 1)

# Carrega o Excel em um DataFrame.
# O RA vira texto sem ".0" (ex.: 9024611.0 vira "9024611") para digitar corretamente.
try:
    df = pd.read_excel(file_path)
    df["RA"] = df["RA"].astype(str).str.replace(r"\.0$", "", regex=True)
except FileNotFoundError:
    print("Arquivo não encontrado!")
    df = pd.DataFrame()

# Lock: garante que só um navegador escreve no Excel por vez (evita conflito)
lock = threading.Lock()


# Função executada por cada navegador, cada um recebe uma fatia de índices
def processar_fatia(fatia):
    driver = webdriver.Chrome()  # abre um navegador Chrome para esta fatia
    for index in fatia:  # percorre os alunos desta fatia
        # Bloco protegido pelo lock: lê o estado atual do aluno e copia os dados
        with lock:
            if str(df.at[index, "INSERIDO"]).strip().upper() == "OK":
                continue  # aluno já inserido na turma, pula
            registrado_ok = str(df.at[index, "REGISTRADO"]).strip().upper() == "OK"
            nome = df.at[index, "NOME"]
            ra = df.at[index, "RA"]
            email = df.at[index, "EMAIL"]
            cod_turma = df.at[index, "COD. TURMA"]

        # Flag: False = precisamos navegar até a tela de login;
        # True  = acabamos de cadastrar e a página já voltou para o login
        ja_no_login = False

        # ETAPA 1: CADASTRO NA PLATAFORMA (só se ainda não está registrado)
        if not registrado_ok:
            try:
                driver.get(URL_CADASTRAR)  # abre a página de cadastro
                time.sleep(7)  # espera a página carregar

                # Preenche o formulário de cadastro
                driver.find_element(By.XPATH, '//input[@placeholder="Nome"]').send_keys(nome)
                driver.find_element(By.XPATH, '//input[@placeholder="Nickname"]').send_keys(ra)
                driver.find_element(By.XPATH, '//input[@placeholder="Email de contato"]').send_keys(email)
                driver.find_element(By.XPATH, '//input[@placeholder="Senha"]').send_keys(PASS_FIXED)
                time.sleep(2)

                # Clica fora (no corpo da página) para tirar o foco dos campos
                elemento_body = driver.find_element(By.TAG_NAME, "body")
                ActionChains(driver).move_to_element(elemento_body).click().perform()
                time.sleep(2)

                # Marca o checkbox de aceite
                driver.find_element(By.XPATH, '//input[@type="checkbox"]').click()

                # Clica no botão "Cadastrar-se"
                botao_cadastrar = driver.find_element(By.XPATH, '//div[contains(@class, "flex-row-reverse")]//p[text()="Cadastrar-se"]/ancestor::button')
                botao_cadastrar.click()
                time.sleep(5)  # espera o cadastro concluir e voltar para o login

                # Cadastro deu certo: grava "OK" em REGISTRADO e salva o Excel
                with lock:
                    df.at[index, "REGISTRADO"] = "OK"
                    df.to_excel(file_path, index=False)
                print(f"CADASTRO OK - {nome} (RA {ra})")
                registrado_ok = True  # libera a etapa de inserção
                ja_no_login = True  # já estamos na tela de login (redirect)
            except Exception as e:
                # Cadastro falhou: grava "FALHA" e salva, depois segue para o próximo
                with lock:
                    df.at[index, "REGISTRADO"] = "FALHA"
                    df.to_excel(file_path, index=False)
                print(f"CADASTRO FALHA - {nome} (RA {ra}): {e}")

        # ETAPA 2: LOGIN + INSERIR CÓDIGO DA TURMA (só se o cadastro deu certo)
        if registrado_ok:
            try:
                # Se não veio direto do cadastro, navega até a tela de login
                if not ja_no_login:
                    driver.get(URL_LOGIN)

                # Espera o campo "Usuário" aparecer (máximo 30s)
                WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, '//input[@placeholder="Usuário"]'))
                )

                # Faz login com RA + senha padrão
                driver.find_element(By.XPATH, '//input[@placeholder="Usuário"]').send_keys(ra)
                driver.find_element(By.XPATH, '//input[@placeholder="Senha"]').send_keys(PASS_FIXED)

                # Clica no botão "Entrar"
                botao_entrar = driver.find_element(By.XPATH, '//div[contains(@class, "flex-row-reverse")]//p[text()="Entrar"]/ancestor::button')
                botao_entrar.click()
                time.sleep(20)  # espera o painel carregar

                # Se aparecer o tutorial, clica em "Dispensar Tutorial"
                dispensar_tutorial_elements = driver.find_elements(By.XPATH, '//button[p[text()="Dispensar Tutorial"]]')
                if dispensar_tutorial_elements:
                    dispensar_tutorial_elements[0].click()
                    time.sleep(5)

                # Clica em "Inserir Código"
                insert_code = driver.find_element(By.XPATH, '//button[p[text()="Inserir Código"]]')
                insert_code.click()
                time.sleep(5)

                # Digita o código da turma e clica em "Continuar"
                driver.find_element(By.XPATH, '//input[@placeholder="Digite o código"]').send_keys(cod_turma)
                codigo_turma = driver.find_element(By.XPATH, '//button[p[text()="Continuar"]]')
                codigo_turma.click()
                time.sleep(5)

                # Confirma o código da turma
                confirma_codigo_turma = driver.find_element(By.XPATH, '//button[p[text()="Continuar"]]')
                confirma_codigo_turma.click()
                time.sleep(15)  # espera a turma ser adicionada

                # Clica no avatar (menu do usuário)
                svg_element = driver.find_element(By.XPATH, '//img[@src="/assets/images/avatar_1.png"]')
                svg_element.click()
                time.sleep(5)

                # Clica em "Sair" para encerrar a sessão
                sair = driver.find_element(By.XPATH, '//button[p[text()="Sair"]]')
                sair.click()
                time.sleep(5)

                # Inserção deu certo: grava "OK" em INSERIDO e salva o Excel
                with lock:
                    df.at[index, "INSERIDO"] = "OK"
                    df.to_excel(file_path, index=False)
                print(f"INSERIDO OK - {nome} (RA {ra})")
            except Exception as e:
                # Inserção falhou: grava "FALHA" e salva, depois segue para o próximo
                with lock:
                    df.at[index, "INSERIDO"] = "FALHA"
                    df.to_excel(file_path, index=False)
                print(f"INSERIDO FALHA - {nome} (RA {ra}): {e}")

    driver.quit()  # fecha o navegador quando a fatia termina


# Lista dos alunos que ainda não foram inseridos (INSERIDO diferente de "OK")
pendencias = [i for i in df.index if str(df.at[i, "INSERIDO"]).strip().upper() != "OK"]

# Divide os pendentes entre os navegadores (intercalando: 1º, 2º, 3º, 4º, 1º, 2º...)
fatias = [[] for _ in range(max(1, NUM_NAVEGADORES))]
for pos, i in enumerate(pendencias):
    fatias[pos % len(fatias)].append(i)
fatias = [f for f in fatias if f]  # remove fatias vazias

print(f"Total pendentes: {len(pendencias)}")
print(f"Navegadores: {len(fatias)}")

# Executa cada fatia em uma thread (um navegador por fatia, de verdade em paralelo)
with ThreadPoolExecutor(max_workers=len(fatias)) as executor:
    executor.map(processar_fatia, fatias)
