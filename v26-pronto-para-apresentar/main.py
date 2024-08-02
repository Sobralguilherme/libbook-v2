import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox
from openpyxl import Workbook

# Função para conectar ao banco de dados e garantir que a tabela seja criada
def conectar_banco():
    try:
        conn = sqlite3.connect('biblioteca.db')
        criar_tabela(conn)
        return conn
    except sqlite3.Error as e:
        messagebox.showerror("Erro", f"Erro ao conectar ao banco de dados: {e}")
        return None

# Função para criar a tabela se não existir
def criar_tabela(conn):
    try:
        cursor = conn.cursor()
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS livros (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            numero_livro TEXT,
            nome TEXT,
            editora TEXT,
            autor TEXT,
            sinopse TEXT,
            disponibilidade TEXT,
            detalhes_extras TEXT
        );
        """)
        conn.commit()
    except sqlite3.Error as e:
        messagebox.showerror("Erro", f"Erro ao criar tabela: {e}")

# Função para validar a entrada do usuário
def validar_entrada():
    if not entry_numero_livro.get().strip():
        messagebox.showwarning("Aviso", "Número do Livro é obrigatório.")
        return False
    if not entry_nome.get().strip():
        messagebox.showwarning("Aviso", "Nome é obrigatório.")
        return False
    if not entry_editora.get().strip():
        messagebox.showwarning("Aviso", "Editora é obrigatória.")
        return False
    if not entry_autor.get().strip():
        messagebox.showwarning("Aviso", "Autor é obrigatório.")
        return False
    if not entry_sinopse.get().strip():
        messagebox.showwarning("Aviso", "Sinopse é obrigatória.")
        return False
    if not var_disponibilidade.get().strip():
        messagebox.showwarning("Aviso", "Disponibilidade é obrigatória.")
        return False
    return True

# Função para adicionar um novo livro
def adicionar_livro_gui():
    if not validar_entrada():
        return
    conn = conectar_banco()
    if conn is None:
        return
    try:
        cursor = conn.cursor()
        numero_livro = entry_numero_livro.get()
        nome = entry_nome.get()
        editora = entry_editora.get()
        autor = entry_autor.get()
        sinopse = entry_sinopse.get()
        disponibilidade = var_disponibilidade.get()
        detalhes_extras = entry_detalhes_extras.get("1.0", tk.END).strip()
        cursor.execute(
            "INSERT INTO livros (numero_livro, nome, editora, autor, sinopse, disponibilidade, detalhes_extras) VALUES (?,?,?,?,?,?,?)",
            (numero_livro, nome, editora, autor, sinopse, disponibilidade, detalhes_extras)
        )
        conn.commit()
        listar_livros_treeview()
        messagebox.showinfo("Sucesso", "Livro adicionado com sucesso.")
    except sqlite3.Error as e:
        messagebox.showerror("Erro", f"Erro ao adicionar livro: {e}")
    finally:
        conn.close()

# Função para listar livros disponíveis com paginação
def listar_livros_treeview(page=1, per_page=10):
    conn = conectar_banco()
    if conn is None:
        return
    try:
        cursor = conn.cursor()
        offset = (page - 1) * per_page
        cursor.execute("SELECT * FROM livros WHERE disponibilidade = 'sim' LIMIT ? OFFSET ?", (per_page, offset))
        livros = cursor.fetchall()
        conn.close()
        tree.delete(*tree.get_children())
        for livro in livros:
            tree.insert('', 'end', values=livro)
        update_pagination_controls(page)
    except sqlite3.Error as e:
        messagebox.showerror("Erro", f"Erro ao listar livros: {e}")

# Função para atualizar os controles de paginação
def update_pagination_controls(current_page):
    global total_pages
    total_pages = (get_total_records() // 10) + 1
    lbl_pagination.config(text=f"Página {current_page} de {total_pages}")

def get_total_records():
    conn = conectar_banco()
    if conn is None:
        return 0
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM livros WHERE disponibilidade = 'sim'")
        total = cursor.fetchone()[0]
        conn.close()
        return total
    except sqlite3.Error as e:
        messagebox.showerror("Erro", f"Erro ao contar registros: {e}")
        return 0

def previous_page():
    global current_page
    if current_page > 1:
        current_page -= 1
        listar_livros_treeview(current_page)

def next_page():
    global current_page
    if current_page < total_pages:
        current_page += 1
        listar_livros_treeview(current_page)

# Função para pesquisar livros
def pesquisar_livro_gui():
    conn = conectar_banco()
    if conn is None:
        return
    try:
        cursor = conn.cursor()
        numero_livro = entry_numero_livro_pesquisa.get().strip()
        nome = entry_nome_pesquisa.get().strip()
        editora = entry_editora_pesquisa.get().strip()
        autor = entry_autor_pesquisa.get().strip()
        disponibilidade = var_disponibilidade_pesquisa.get().strip()

        query = "SELECT * FROM livros WHERE 1=1"
        params = []

        if numero_livro:
            query += " AND numero_livro = ?"
            params.append(numero_livro)
        if nome:
            query += " AND nome LIKE ?"
            params.append(f"%{nome}%")
        if editora:
            query += " AND editora LIKE ?"
            params.append(f"%{editora}%")
        if autor:
            query += " AND autor LIKE ?"
            params.append(f"%{autor}%")
        if disponibilidade:
            query += " AND disponibilidade = ?"
            params.append(disponibilidade)

        cursor.execute(query, params)
        livros = cursor.fetchall()
        conn.close()
        tree.delete(*tree.get_children())
        for livro in livros:
            tree.insert('', 'end', values=livro)
    except sqlite3.Error as e:
        messagebox.showerror("Erro", f"Erro ao pesquisar livro: {e}")

# Função para marcar a disponibilidade
def marcar_disponibilidade_gui():
    conn = conectar_banco()
    if conn is None:
        return
    try:
        cursor = conn.cursor()
        numero_livro = entry_numero_livro_disponibilidade.get().strip()
        nova_disponibilidade = var_nova_disponibilidade.get().strip()
        cursor.execute("UPDATE livros SET disponibilidade = ? WHERE numero_livro = ?", (nova_disponibilidade, numero_livro))
        conn.commit()
        listar_livros_treeview()
        messagebox.showinfo("Sucesso", "Disponibilidade atualizada com sucesso.")
    except sqlite3.Error as e:
        messagebox.showerror("Erro", f"Erro ao marcar disponibilidade: {e}")
    finally:
        conn.close()

# Função para exportar dados para Excel
def exportar_para_excel_gui():
    workbook = Workbook()
    worksheet = workbook.active
    conn = conectar_banco()
    if conn is None:
        return
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM livros")
        dados = cursor.fetchall()
        conn.close()
        for linha in dados:
            worksheet.append(linha)
        workbook.save("livros.xlsx")
        messagebox.showinfo("Sucesso", "Dados exportados para Excel com sucesso.")
    except sqlite3.Error as e:
        messagebox.showerror("Erro", f"Erro ao exportar para Excel: {e}")

# Função para deletar um livro
def deletar_livro_gui():
    conn = conectar_banco()
    if conn is None:
        return
    try:
        cursor = conn.cursor()
        numero_livro = entry_numero_livro_deletar.get().strip()
        cursor.execute("DELETE FROM livros WHERE numero_livro = ?", (numero_livro,))
        conn.commit()
        listar_livros_treeview()
        messagebox.showinfo("Sucesso", "Livro deletado com sucesso.")
    except sqlite3.Error as e:
        messagebox.showerror("Erro", f"Erro ao deletar livro: {e}")
    finally:
        conn.close()

# Função para carregar dados do livro para edição
def carregar_dados_livro_gui():
    conn = conectar_banco()
    if conn is None:
        return
    try:
        cursor = conn.cursor()
        numero_livro = entry_numero_livro_editar.get().strip()
        cursor.execute("SELECT * FROM livros WHERE numero_livro = ?", (numero_livro,))
        livro_encontrado = cursor.fetchone()
        conn.close()
        if livro_encontrado:
            entry_numero_livro.delete(0, tk.END)
            entry_nome.delete(0, tk.END)
            entry_editora.delete(0, tk.END)
            entry_autor.delete(0, tk.END)
            entry_sinopse.delete(0, tk.END)
            entry_detalhes_extras.delete("1.0", tk.END)

            entry_numero_livro.insert(0, livro_encontrado[1])
            entry_nome.insert(0, livro_encontrado[2])
            entry_editora.insert(0, livro_encontrado[3])
            entry_autor.insert(0, livro_encontrado[4])
            entry_sinopse.insert(0, livro_encontrado[5])
            var_disponibilidade.set(livro_encontrado[6])
            entry_detalhes_extras.insert("1.0", livro_encontrado[7])
        else:
            messagebox.showinfo("Resultado da Pesquisa", "Livro não encontrado.")
    except sqlite3.Error as e:
        messagebox.showerror("Erro", f"Erro ao carregar dados do livro: {e}")

# Função para editar um livro
def editar_livro_gui():
    if not validar_entrada():
        return
    conn = conectar_banco()
    if conn is None:
        return
    try:
        cursor = conn.cursor()
        numero_livro = entry_numero_livro_editar.get().strip()
        novo_numero_livro = entry_numero_livro.get().strip()
        nome = entry_nome.get().strip()
        editora = entry_editora.get().strip()
        autor = entry_autor.get().strip()
        sinopse = entry_sinopse.get().strip()
        disponibilidade = var_disponibilidade.get().strip()
        detalhes_extras = entry_detalhes_extras.get("1.0", tk.END).strip()
        cursor.execute("""
        UPDATE livros SET 
            numero_livro = ?, 
            nome = ?, 
            editora = ?, 
            autor = ?, 
            sinopse = ?, 
            disponibilidade = ?, 
            detalhes_extras = ? 
        WHERE numero_livro = ?""",
        (novo_numero_livro, nome, editora, autor, sinopse, disponibilidade, detalhes_extras, numero_livro))
        conn.commit()
        listar_livros_treeview()
        messagebox.showinfo("Sucesso", "Livro editado com sucesso.")
    except sqlite3.Error as e:
        messagebox.showerror("Erro", f"Erro ao editar livro: {e}")
    finally:
        conn.close()

# Função para filtrar livros dinamicamente
def filtrar_livros_gui(event):
    conn = conectar_banco()
    if conn is None:
        return
    try:
        cursor = conn.cursor()
        numero_livro = entry_numero_livro_pesquisa.get().strip()
        nome = entry_nome_pesquisa.get().strip()
        editora = entry_editora_pesquisa.get().strip()
        autor = entry_autor_pesquisa.get().strip()
        disponibilidade = var_disponibilidade_pesquisa.get().strip()

        query = "SELECT * FROM livros WHERE 1=1"
        params = []

        if numero_livro:
            query += " AND numero_livro LIKE ?"
            params.append(f"%{numero_livro}%")
        if nome:
            query += " AND nome LIKE ?"
            params.append(f"%{nome}%")
        if editora:
            query += " AND editora LIKE ?"
            params.append(f"%{editora}%")
        if autor:
            query += " AND autor LIKE ?"
            params.append(f"%{autor}%")
        if disponibilidade:
            query += " AND disponibilidade = ?"
            params.append(disponibilidade)

        cursor.execute(query, params)
        livros = cursor.fetchall()
        conn.close()
        tree.delete(*tree.get_children())
        for livro in livros:
            tree.insert('', 'end', values=livro)
    except sqlite3.Error as e:
        messagebox.showerror("Erro", f"Erro ao filtrar livros: {e}")

# Função principal para iniciar o GUI
def main():
    global tree, entry_numero_livro, entry_nome, entry_editora, entry_autor, entry_sinopse, var_disponibilidade, entry_detalhes_extras
    global entry_numero_livro_pesquisa, entry_nome_pesquisa, entry_editora_pesquisa, entry_autor_pesquisa, var_disponibilidade_pesquisa
    global entry_numero_livro_disponibilidade, var_nova_disponibilidade, entry_numero_livro_deletar, entry_numero_livro_editar
    global lbl_pagination, current_page, total_pages

    current_page = 1

    root = tk.Tk()
    root.title("Sistema de Cadastro de Livros")
    root.geometry("800x600")  # Definir tamanho inicial da janela

    # Configuração para redimensionamento
    for i in range(24):  # Ajuste para o número de linhas que você tem
        root.grid_rowconfigure(i, weight=1)
    for i in range(3):  # Ajuste para o número de colunas que você tem
        root.grid_columnconfigure(i, weight=1)

    # Widgets de entrada
    tk.Label(root, text="Número do Livro").grid(row=0, column=0, sticky='nsew')
    entry_numero_livro = tk.Entry(root)
    entry_numero_livro.grid(row=0, column=1, sticky='nsew')

    tk.Label(root, text="Nome").grid(row=1, column=0, sticky='nsew')
    entry_nome = tk.Entry(root)
    entry_nome.grid(row=1, column=1, sticky='nsew')

    tk.Label(root, text="Editora").grid(row=2, column=0, sticky='nsew')
    entry_editora = tk.Entry(root)
    entry_editora.grid(row=2, column=1, sticky='nsew')

    tk.Label(root, text="Autor").grid(row=3, column=0, sticky='nsew')
    entry_autor = tk.Entry(root)
    entry_autor.grid(row=3, column=1, sticky='nsew')

    tk.Label(root, text="Sinopse").grid(row=4, column=0, sticky='nsew')
    entry_sinopse = tk.Entry(root)
    entry_sinopse.grid(row=4, column=1, sticky='nsew')

    tk.Label(root, text="Disponibilidade").grid(row=5, column=0, sticky='nsew')
    var_disponibilidade = tk.StringVar()
    entry_disponibilidade = ttk.Combobox(root, textvariable=var_disponibilidade, values=["sim", "não"])
    entry_disponibilidade.grid(row=5, column=1, sticky='nsew')

    tk.Label(root, text="Detalhes Extras").grid(row=6, column=0, sticky='nsew')
    entry_detalhes_extras = tk.Text(root, height=5, width=30)
    entry_detalhes_extras.grid(row=6, column=1, sticky='nsew')

    # Adicionar scrollbar para Text widget
    scrollbar_detalhes = tk.Scrollbar(root, command=entry_detalhes_extras.yview)
    entry_detalhes_extras.config(yscrollcommand=scrollbar_detalhes.set)
    scrollbar_detalhes.grid(row=6, column=2, sticky='ns')

    # Árvore para listar livros
    tree_frame = tk.Frame(root)
    tree_frame.grid(row=7, column=0, columnspan=3, sticky='nsew')

    tree_scrollbar_y = tk.Scrollbar(tree_frame, orient=tk.VERTICAL)
    tree_scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)

    tree_scrollbar_x = tk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
    tree_scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)

    tree = ttk.Treeview(tree_frame, columns=('ID', 'Numero do Livro', 'Nome', 'Editora', 'Autor', 'Sinopse', 'Disponibilidade', 'Detalhes Extras'), 
                        show='headings', yscrollcommand=tree_scrollbar_y.set, xscrollcommand=tree_scrollbar_x.set)
    tree.pack(fill=tk.BOTH, expand=True)

    tree_scrollbar_y.config(command=tree.yview)
    tree_scrollbar_x.config(command=tree.xview)

    tree.heading('ID', text='ID')
    tree.heading('Numero do Livro', text='Número do Livro')
    tree.heading('Nome', text='Nome')
    tree.heading('Editora', text='Editora')
    tree.heading('Autor', text='Autor')
    tree.heading('Sinopse', text='Sinopse')
    tree.heading('Disponibilidade', text='Disponibilidade')
    tree.heading('Detalhes Extras', text='Detalhes Extras')

    # Paginação
    lbl_pagination = tk.Label(root, text="")
    lbl_pagination.grid(row=8, column=0, columnspan=2, sticky='nsew')

    tk.Button(root, text="Página Anterior", command=previous_page).grid(row=9, column=0, sticky='nsew')
    tk.Button(root, text="Próxima Página", command=next_page).grid(row=9, column=1, sticky='nsew')

    # Botões
    tk.Button(root, text="Adicionar Livro", command=adicionar_livro_gui).grid(row=10, column=0, pady=5, sticky='nsew')
    tk.Button(root, text="Listar Livros Disponíveis", command=lambda: listar_livros_treeview(1)).grid(row=10, column=1, pady=5, sticky='nsew')

    tk.Label(root, text="Número do Livro para Pesquisar").grid(row=11, column=0, sticky='nsew')
    entry_numero_livro_pesquisa = tk.Entry(root)
    entry_numero_livro_pesquisa.grid(row=11, column=1, sticky='nsew')
    tk.Label(root, text="Nome para Pesquisar").grid(row=12, column=0, sticky='nsew')
    entry_nome_pesquisa = tk.Entry(root)
    entry_nome_pesquisa.grid(row=12, column=1, sticky='nsew')
    tk.Label(root, text="Editora para Pesquisar").grid(row=13, column=0, sticky='nsew')
    entry_editora_pesquisa = tk.Entry(root)
    entry_editora_pesquisa.grid(row=13, column=1, sticky='nsew')
    tk.Label(root, text="Autor para Pesquisar").grid(row=14, column=0, sticky='nsew')
    entry_autor_pesquisa = tk.Entry(root)
    entry_autor_pesquisa.grid(row=14, column=1, sticky='nsew')
    tk.Label(root, text="Disponibilidade para Pesquisar").grid(row=15, column=0, sticky='nsew')
    var_disponibilidade_pesquisa = tk.StringVar()
    entry_disponibilidade_pesquisa = ttk.Combobox(root, textvariable=var_disponibilidade_pesquisa, values=["sim", "não"])
    entry_disponibilidade_pesquisa.grid(row=15, column=1, sticky='nsew')

    # Adicionar eventos para filtros dinâmicos
    entry_numero_livro_pesquisa.bind("<KeyRelease>", filtrar_livros_gui)
    entry_nome_pesquisa.bind("<KeyRelease>", filtrar_livros_gui)
    entry_editora_pesquisa.bind("<KeyRelease>", filtrar_livros_gui)
    entry_autor_pesquisa.bind("<KeyRelease>", filtrar_livros_gui)
    entry_disponibilidade_pesquisa.bind("<<ComboboxSelected>>", filtrar_livros_gui)

    tk.Button(root, text="Número do Livro para Marcar Disponibilidade").grid(row=17, column=0, sticky='nsew')
    entry_numero_livro_disponibilidade = tk.Entry(root)
    entry_numero_livro_disponibilidade.grid(row=17, column=1, sticky='nsew')

    tk.Label(root, text="Nova Disponibilidade").grid(row=18, column=0, sticky='nsew')
    var_nova_disponibilidade = tk.StringVar()
    entry_nova_disponibilidade = ttk.Combobox(root, textvariable=var_nova_disponibilidade, values=["sim", "não"])
    entry_nova_disponibilidade.grid(row=18, column=1, sticky='nsew')
    tk.Button(root, text="Marcar Disponibilidade", command=marcar_disponibilidade_gui).grid(row=19, column=0, pady=5, sticky='nsew')

    tk.Button(root, text="Exportar para Excel", command=exportar_para_excel_gui).grid(row=19, column=1, pady=5, sticky='nsew')

    tk.Label(root, text="Número do Livro para Deletar").grid(row=20, column=0, sticky='nsew')
    entry_numero_livro_deletar = tk.Entry(root)
    entry_numero_livro_deletar.grid(row=20, column=1, sticky='nsew')
    tk.Button(root, text="Deletar Livro", command=deletar_livro_gui).grid(row=21, column=0, pady=5, sticky='nsew')

    tk.Label(root, text="Número do Livro para Editar").grid(row=22, column=0, sticky='nsew')
    entry_numero_livro_editar = tk.Entry(root)
    entry_numero_livro_editar.grid(row=22, column=1, sticky='nsew')
    tk.Button(root, text="Carregar Dados para Edição", command=carregar_dados_livro_gui).grid(row=23, column=0, pady=5, sticky='nsew')
    tk.Button(root, text="Editar Livro", command=editar_livro_gui).grid(row=23, column=1, pady=5, sticky='nsew')

    listar_livros_treeview(current_page)

    root.mainloop()

if __name__ == "__main__":
    main()
