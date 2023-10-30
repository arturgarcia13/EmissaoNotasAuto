import subprocess

# Lista de bibliotecas necessárias
dependencias = [
    'pandas',
    'numpy',
    'selenium',
    'undetected-chromedriver',
    'auto-py-to-exe',
    'pyautogui'
]

# Função para verificar e instalar bibliotecas ausentes
def verificar_e_instalar_dependencias():
    for lib in dependencias:
        try:
            __import__(lib)
            print(f'{lib} está instalado.')
        except ImportError:
            print(f'{lib} não está instalado. Instalando...')
            subprocess.call(['pip', 'install', lib])
            print(f'{lib} foi instalado com sucesso.')

if __name__ == "__main__":
    print("Verificando e instalando dependências...")
    verificar_e_instalar_dependencias()
    print("Todas as dependências foram verificadas e instaladas.")
