import os
from dotenv import load_dotenv
import requests

# Carrega as variáveis de ambiente do arquivo .env
load_dotenv()

# Função para perguntar algo à API da Groq
def perguntar_groq(pergunta):
    # Configuração da API
    API_KEY = os.getenv("GROQ_API_KEY")  # Lê a chave da variável de ambiente
    API_URL = "https://api.groq.com/openai/v1/chat/completions"  # URL correta

    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json"
    }
    
    data = {
        "model": "mixtral-8x7b-32768",  # Modelo suportado pela Groq
        "messages": [{"role": "user", "content": pergunta}],
        "temperature": 0.7
    }
    
    response = requests.post(API_URL, headers=headers, json=data)
    
    if response.status_code == 200:
        resposta = response.json()
        return resposta["choices"][0]["message"]["content"]
    else:
        return f"Erro: {response.status_code} - {response.text}"
    

