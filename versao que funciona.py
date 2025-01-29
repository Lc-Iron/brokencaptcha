import os  
import tempfile  
from selenium import webdriver  
from selenium.webdriver.edge.service import Service  
from selenium.webdriver.edge.webdriver import WebDriver  
from selenium.webdriver.common.by import By  
import requests  
import whisper  
import subprocess  
import re 
 
# Adiciona o caminho do FFmpeg ao PATH  
ffmpeg_path = r'C:\ffmpeg\bin'  
os.environ["PATH"] += os.pathsep + ffmpeg_path  
 
# Verifica se o FFmpeg está acessível  
try:  
    result = subprocess.run(['ffmpeg', '-version'], capture_output=True, text=True)  
    print("FFmpeg version:", result.stdout.split('\n')[0])  
except FileNotFoundError:  
    print("FFmpeg não encontrado no PATH. Verifique o caminho e a instalação.")  
    exit(1)  
 
# Configuração do EdgeDriver  
service = Service(r"c:\Users\FX4S\Downloads\Python\edgedriver_win64\msedgedriver.exe")  
driver = WebDriver(service=service)  
 
try:  
    # Acesse a página  
    driver.get("https://sei.anp.gov.br/sei/modulos/pesquisa/md_pesq_processo_pesquisar.php?acao_externa=protocolo_pesquisar&acao_origem_externa=protocolo_pesquisar&id_orgao_acesso_externo=0")  
    driver.implicitly_wait(2)  
 
    # Captura o link de áudio  
    src = driver.find_element(By.XPATH, "/html/body/div[1]/div/div/form/div/div[2]/div[1]/div[2]/audio/source")  
    link_audio = src.get_attribute("src")  
    print("Link do áudio:", link_audio)  
 
    # Baixa o áudio usando cookies do Selenium  
    session = requests.Session()  
    for cookie in driver.get_cookies():  
        session.cookies.set(cookie['name'], cookie['value'])  
 
    # Baixa e salva o áudio em um arquivo temporário  
    temp_dir = tempfile.gettempdir()  
    mp3_file_path = os.path.join(temp_dir, "temp_audio.mp3")  
    wav_file_path = os.path.join(temp_dir, "temp_audio.wav")  
 
    response = session.get(link_audio, stream=True)  
    if response.status_code == 200:  
        with open(mp3_file_path, 'wb') as file:  
            for chunk in response.iter_content(chunk_size=8192):  
                file.write(chunk)  
        print("Áudio salvo em:", mp3_file_path)  
 
        # Verifica se o arquivo realmente existe e seu tamanho  
        if os.path.exists(mp3_file_path):  
            print(f"O arquivo foi criado com sucesso em {mp3_file_path}")  
            print(f"Tamanho do arquivo: {os.path.getsize(mp3_file_path)} bytes")  
            print(f"Permissões do arquivo: {oct(os.stat(mp3_file_path).st_mode)[-3:]}")  
        else:  
            print(f"Erro: O arquivo {mp3_file_path} não foi encontrado.")  
            exit(1)  
 
        # Converte para WAV  
        try:  
            subprocess.run(['ffmpeg', '-i', mp3_file_path, '-y', wav_file_path],  
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)  
            print(f"Arquivo convertido para WAV: {wav_file_path}")  
        except subprocess.CalledProcessError as e:  
            print(f"Erro ao converter o arquivo para WAV: {e}")  
            exit(1)  
 
        # Transcrição com Whisper  
        print("Tentando carregar o modelo Whisper...")  
        print(f"Versão do Whisper: {whisper.__version__}")  
        modelo = whisper.load_model("base")  
        print("Modelo Whisper carregado com sucesso.")  
 
        # Confirmação do caminho do arquivo antes da transcrição  
        if os.path.exists(wav_file_path):  
            print(f"Arquivo de áudio confirmado para transcrição: {wav_file_path}")  
            try:  
                print(f"Caminho do arquivo usado pelo Whisper: {os.path.abspath(wav_file_path)}")  
                resposta = modelo.transcribe(wav_file_path, language="pt")  
                print("Texto transcrito:", resposta['text'])  
                 
                # Extrai números e letras maiúsculas  
                captcha = re.sub(r'[^A-Z0-9]', '', resposta['text'].upper())  
                print("Captcha:", captcha)  
            except Exception as e:  
                print(f"Erro durante a transcrição do áudio em {wav_file_path}: {e}")  
        else:  
            print(f"Erro: O arquivo de áudio {wav_file_path} não foi encontrado antes da transcrição.")  
    else:  
        print("Erro ao baixar o áudio. Status code:", response.status_code)  
 
except Exception as e:  
    print("Erro encontrado:", str(e))  
 
finally:  
    # Fechar o driver e remover arquivos temporários  
    driver.quit()  
    if os.path.exists(mp3_file_path):  
        os.remove(mp3_file_path)  
    if os.path.exists(wav_file_path):  
        os.remove(wav_file_path)  

