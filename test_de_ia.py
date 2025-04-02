from langchain.chains import LLMChain
from langchain.prompts import PromptTemplate
from langchain.llms import BaseLLM
from langchain.schema import Generation, LLMResult
from typing import List, Optional, Dict, Any
import requests
import json
import os
import glob


# Configuração da API da Groq
GROQ_API_URL = "https://api.groq.com/openai/v1/chat/completions"
GROQ_API_KEY = "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"

# Criando um wrapper personalizado para o modelo da Groq
class GroqLLM(BaseLLM):
    def _generate(
        self, prompts: List[str], stop: Optional[List[str]] = None
    ) -> LLMResult:
        generations = []
        for prompt in prompts:
            # Configuração dos cabeçalhos da requisição HTTP
            headers = {
                "Authorization": f"Bearer {GROQ_API_KEY}",
                "Content-Type": "application/json"
            }
            
            # Configuração dos dados da requisição
            data = {
                "prompt": prompt,
                "max_tokens": 150,
                "temperature": 0.7
            }
            
            # Fazendo a requisição POST para a API da Groq
            response = requests.post(GROQ_API_URL, headers=headers, data=json.dumps(data))
            
            # Verificando se a requisição foi bem-sucedida
            if response.status_code == 200:
                text = response.json()["choices"][0]["text"]
                generations.append([Generation(text=text)])
            else:
                raise Exception(f"Erro na API da Groq: {response.status_code}, {response.text}")
        
        return LLMResult(generations=generations)

    @property
    def _llm_type(self) -> str:
        return "groq"

# Função para encontrar o último documento em uma pasta
def encontrar_ultimo_documento(pasta):
    # Lista todos os arquivos na pasta
    arquivos = glob.glob(os.path.join(pasta, "*"))
    
    # Filtra apenas arquivos (ignora pastas)
    arquivos = [arquivo for arquivo in arquivos if os.path.isfile(arquivo)]
    
    if not arquivos:
        return "Nenhum arquivo encontrado na pasta."
    
    # Encontra o arquivo mais recente com base na data de modificação
    ultimo_arquivo = max(arquivos, key=os.path.getmtime)
    
    # Retorna o nome do arquivo
    return os.path.basename(ultimo_arquivo)

# Definindo o template do prompt
prompt_template = """
Você é um assistente virtual especializado em ajudar desenvolvedores.
Pergunta: {pergunta}
Resposta:
"""

# Criando o prompt
prompt = PromptTemplate(
    input_variables=["pergunta"],
    template=prompt_template
)

# Inicializando o modelo da Groq
llm = GroqLLM()

# Criando a cadeia (chain) com o prompt e o modelo
chain = LLMChain(llm=llm, prompt=prompt)

# Função para interagir com o chatbot
def chatbot(pergunta):
    # Verifica se a pergunta é sobre o último documento na pasta
    if "último documento" in pergunta.lower() and "pasta" in pergunta.lower():
        # Extrai o caminho da pasta da pergunta (exemplo: "pasta documentos")
        
        pasta = r'C:\\Users\\Usuario\\Desktop\\nova_pasta_descritivo\\TEMPLATES'
        caminho_pasta = pasta
        if os.path.exists(caminho_pasta):
            resposta = encontrar_ultimo_documento(caminho_pasta)
            return f"O último documento na pasta '{pasta}' é: {resposta}"
        else:
            return f"A pasta '{pasta}' não foi encontrada."
    else:
        # Caso contrário, usa o modelo da Groq para responder
        resposta = chain.run(pergunta=pergunta)
        return resposta

# Exemplo de uso
pergunta = "Me dê o nome do último documento que está na pasta de templates"
resposta = chatbot(pergunta)
print(resposta)