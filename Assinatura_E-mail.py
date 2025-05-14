import os
import sys
from ldap3 import Server, Connection, SUBTREE

# Configuração LDAP (ajuste conforme o seu ambiente)
LDAP_SERVER = "meu.dominio.com.br"
BASE_DN = "OU=Empresa,DC=meu,DC=dominio,DC=com,DC=br"
BIND_DN = "CN=ContaAdministrativa,OU=Usuario,OU=Servico,OU=Empresa,OU=dominio,DC=ad,DC=zortea,DC=com,DC=br"
BIND_PASSWORD = "Senha"

# Caminho do arquivo de template RTF
TEMPLATE_FILE = r"\\192.168.1.1\meu\diretorio\template.rtf"

if not os.path.exists(TEMPLATE_FILE):
    print(f"Arquivo de template RTF não encontrado: {TEMPLATE_FILE}")
    sys.exit(1)

# Solicita que o operador informe o e-mail do usuário
user_email = input("Digite o e-mail do usuário para busca: ").strip()

# Conectar-se ao servidor LDAP
server = Server(LDAP_SERVER)
conn = Connection(server, BIND_DN, BIND_PASSWORD, auto_bind=True)

# Executa a consulta LDAP
conn.search(BASE_DN, f"(mail={user_email})", search_scope=SUBTREE,
            attributes=["displayName", "department", "homePhone", "mobile", "sAMAccountName"])

# Verifica se há resultado
if not conn.entries:
    print(f"Usuário não encontrado com o e-mail: {user_email}")
    sys.exit(1)

# Recupera os atributos do usuário
user_data = conn.entries[0]
display_name = user_data.displayName.value if user_data.displayName else ""
department = user_data.department.value if user_data.department else ""
home_phone = user_data.homePhone.value if user_data.homePhone else "+55 (00) 0000-0000"
mobile = user_data.mobile.value if user_data.mobile else ""
sam_account_name = user_data.sAMAccountName.value if user_data.sAMAccountName else "usuario_desconhecido"

# Lê o conteúdo do template
with open(TEMPLATE_FILE, "r", encoding="utf-8") as file:
    template_content = file.read()

# Substitui os placeholders
template_content = template_content.replace("%%DisplayName%%", display_name)
template_content = template_content.replace("%%Department%%", department)
template_content = template_content.replace("%%HomePhone%%", home_phone)
template_content = template_content.replace("%%Mobile%%", mobile)

# Define o diretório de saída
output_folder = os.path.join(os.path.expanduser("~"), "Desktop")
os.makedirs(output_folder, exist_ok=True)

# Define o nome do arquivo de saída
output_file = os.path.join(output_folder, f"{sam_account_name}_signature.rtf")

# Salva o arquivo atualizado
with open(output_file, "w", encoding="utf-8") as file:
    file.write(template_content)

print(f"Assinatura RTF atualizada para {display_name} gerada em: {output_file}")