import pandas as pd
from urllib.parse import quote
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import paho.mqtt.client as mqtt
import requests

# Configuração MQTT 
broker = "broker.hivemq.com"
port = 1883
channel_id = "2659682"
write_api_key = "VQGTG8G9AV2U78OI"
topic = f"channels/2659682/publish/VQGTG8G9AV2U78OI"
client_id = "2659682"

def on_connect(client, userdata, flags, rc):
    if rc == 0:
        print("Conectado ao broker MQTT com sucesso")
        client.connected_flag = True
    else:
        print(f"Falha na conexão com o broker MQTT, código de retorno: {rc}")
        client.connected_flag = False

# Publicação no MQTT
def on_publish(client, userdata, mid):
    print(f"Mensagem {mid} publicada com sucesso")

client = mqtt.Client()
mqtt.Client.connected_flag = False  # Adicionar flag de conexão

client.on_connect = on_connect
client.on_publish = on_publish

try:
    client.connect(broker, port, 60)
    client.loop_start()
    while not client.connected_flag:  # Esperar até que a conexão seja estabelecida
        print("Aguardando conexão MQTT...")
        time.sleep(1)
except Exception as e:
    print(f"Erro ao conectar ao broker MQTT: {e}")

def enviar_dados_thingspeak_http(telefone):
    url = f"https://api.thingspeak.com/update?api_key={write_api_key}&field1={telefone}"
    response = requests.get(url)
    if response.status_code == 200:
        print(f"Dados enviados ao ThingSpeak via HTTP: {response.text}")
    else:
        print(f"Falha ao enviar dados ao ThingSpeak via HTTP, código de retorno: {response.status_code}")

# Abrir e enviar mensagem 
navegador = webdriver.Chrome()
navegador.get("https://web.whatsapp.com/")

while len(navegador.find_elements(By.ID, "side")) < 1:
    time.sleep(1)
time.sleep(3)

# leitura da planilha
df = pd.read_excel("C:\\Users\\gabri\\OneDrive\\Área de Trabalho\\Projeto IoT\\diagnostico\\diagnostico.xlsx")
df.fillna(0, inplace=True)

print(df.info())

# Tratar erro de telefone
def limpar_telefone(numero):
    if isinstance(numero, str):
        return numero.replace('-', '').replace('(', '').replace(')', '').replace(' ', '').replace('+', '').replace('WHATS', '')
    return numero

df['Telefone'] = df['Telefone'].apply(limpar_telefone)

# Percorrer toda a planilha
for linha in df.itertuples():
    placa = linha[2]
    cliente = linha[4]
    telefone = linha[7]
    
    try:
        # Verificar o número / inteiro
        if not telefone.isdigit():
            raise ValueError(f"Telefone não é um número válido: {telefone}")
        telefone = int(telefone)

        mensagem_whats = f'Prezado cliente, {cliente}, consta em nosso relatório de diagnóstico que seu veículo de Placa: {placa} está parado há mais de 3 dias sem atualizar data, hora e a posição. Caso esteja funcionando normalmente favor entrar em contato conosco com urgência através dos telefones (16)98876-0515 ou (16)3505-2040. Mensagem enviada automaticamente.'
        
        link_whatsApp = f'https://web.whatsapp.com/send?phone=55{telefone}&text={quote(mensagem_whats)}'
        navegador.get(link_whatsApp)
        time.sleep(5)  # Aumentar o tempo de espera se necessário

        while len(navegador.find_elements(By.ID, "side")) < 1:
            time.sleep(1)
        time.sleep(12)

        navegador.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span').click()
        time.sleep(3)

        # Publicar dados no ThingSpeak via MQTT
        if client.connected_flag:
            payload = f"field1={telefone}"
            result = client.publish(topic, payload)
            status = result.rc
            if status == mqtt.MQTT_ERR_SUCCESS:
                print(f"Dados enviados ao ThingSpeak via MQTT: {payload}")
            else:
                print(f"Falha ao enviar dados ao ThingSpeak via MQTT, código de retorno: {status}")
        else:
            print("Cliente MQTT não está conectado")
            
        enviar_dados_thingspeak_http(telefone)

    except Exception as e:
        print(f'Não foi possível enviar a mensagem para {cliente} com o telefone {telefone}')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{cliente},{telefone}\n')

navegador.quit()