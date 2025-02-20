import tkinter as tk

def escolher_Texto_mosaicos():
    resultado = []

    def selecionar_opcao(descricao, botao):
        if descricao in resultado:
            # Se já estiver selecionado, remove e restaura o botão
            resultado.remove(descricao)
            botao.config(bg="SystemButtonFace", relief=tk.RAISED)  # Restaura a cor padrão
        else:
            # Adiciona à lista e muda a aparência do botão
            resultado.append(descricao)
            botao.config(bg="lightgreen", relief=tk.SUNKEN)  # Destaque visual

    def finalizar_selecao(janela):
        janela.destroy()

    janela = tk.Tk()
    janela.title("Menu de Mosaico")

    label = tk.Label(janela, text="CASO NÃO HOUVER MOSAICO \nAPENAS APERTE NO (X)", font=("Arial", 14))
    label.pack(pady=10)

    botoes = [
        ("AB", "AB – Mosaico com predomínio de A sobre B"),
        ("BA", "BA – Mosaico com predomínio de B sobre A"),
        ("BC", "BC – Mosaico com predomínio de B sobre C"),
        ("CB", "CB – Mosaico com predomínio de C sobre B"),
        ("CD", "CD – Mosaico com predomínio de C sobre D"),
        ("DC", "DC – Mosaico com predomínio de D sobre C")
    ]

    for codigo, descricao in botoes:
        # Cria o botão e passa ele diretamente para a lambda
        botao = tk.Button(
            janela,
            text=descricao,
            font=("Arial", 12),
        )
        # Configura o comando do botão após sua criação
        botao.config(command=lambda d=descricao, b=botao: selecionar_opcao(d, b))
        botao.pack(fill=tk.X, pady=5, padx=20)

    botao_concluir = tk.Button(janela, text="Concluir", font=("Arial", 12), command=lambda: finalizar_selecao(janela))
    botao_concluir.pack(fill=tk.X, pady=5, padx=20)

    janela.mainloop()

    if resultado:
        return "<tag>" + "</tag>  <tag>".join(resultado) + "</tag>"
    else:
        return None

