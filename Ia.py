import os
from dotenv import load_dotenv
import requests

# Carrega as variáveis de ambiente do arquivo .env
load_dotenv()

# Constantes para configuração da API
API_KEY = os.getenv("api_grok")  # Lê a chave da variável de ambiente
API_URL = "https://api.groq.com/openai/v1/chat/completions"  # URL da API
MODEL = "mixtral-8x7b-32768"  # Modelo suportado pela Groq
TEMPERATURE = 0.7  # Parâmetro de criatividade da resposta

# Função para fazer uma requisição à API da Groq
def perguntar_groq_geografia(pergunta):
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json"
    }
    
    # Define o perfil do agente como um professor de geografia
    mensagem_sistema = "Você é um professor de geografia especializado na caracterização de todas as regiões do brasil. Responda de forma clara, educativa e detalhada. não precisa se apresentar apenas me envie a resposta"
    
    data = {
        "model": MODEL,
        "messages": [
            {"role": "system", "content": mensagem_sistema},  # Define o perfil do agente
            {"role": "user", "content": pergunta}  # Pergunta do usuário
        ],
        "temperature": TEMPERATURE
    }
    
    try:
        response = requests.post(API_URL, headers=headers, json=data)
        response.raise_for_status()  # Levanta uma exceção para erros HTTP
        resposta = response.json()
        return resposta["choices"][0]["message"]["content"]
    except requests.exceptions.RequestException as e:
        return f"Erro na requisição: {e}"
    except KeyError:
        return "Erro: Resposta da API em formato inesperado."

