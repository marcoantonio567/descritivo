#api = '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'
import requests



# Função para perguntar algo à API da Groq
def perguntar_groq(pergunta):
    # Configuração da API
    API_KEY = "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"  # Substitua pela sua chave real
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

# Exemplo de uso com limite de 50 tokens   
