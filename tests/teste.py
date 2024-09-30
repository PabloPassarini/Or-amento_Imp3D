import os

# Caminho original
caminho = r"f:\Projetos\Programação\Gasto3D\src\main.py"

# Divide o caminho
parte1, parte2 = os.path.split(caminho)  # parte1 é 'f:\Projetos\Programação\Gasto3D\src' e parte2 é 'main.py'
parte1, parte3 = os.path.split(parte1)  # parte1 agora é 'f:\Projetos\Programação\Gasto3D' e parte3 é 'src'

# O resultado final é a parte1
caminho_sem_dois_ultimos = parte1

print(caminho_sem_dois_ultimos)