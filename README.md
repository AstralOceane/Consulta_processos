# Automação de Consulta e Salvamento de Dados em Python

Este é um exemplo de código em Python que automatiza a consulta de processos em um site e salva as informações em planilhas do Excel usando o Selenium e a biblioteca openpyxl.

## Pré-requisitos

- Python 3.x
- pip install -r Requirements.txt (para instalar todas as biblitoecas necessarias para rodar o codigo)

## Configuração

1. Clone este repositório para o seu sistema:

git clone https://github.com/AstralOceane/Consulta_processos.git

2. Preencha as informações necessárias nas variáveis :

    # Função para enviar e-mail
        de = 'seu_email@gmail.com'
        senha = 'sua_senha'
        para = 'destinatario@gmail.com'
    # Digitar o número da OAB
    numero_oab = numero_do_processo

    # Selecionar o estado
    codigo_estado = 'SP'

3. Execute o código Python:

python app.py

4. Os dados serão coletados e armazenados em planilhas do Excel no arquivo "dados.xlsx".

## Personalização

Você pode personalizar o código de acordo com suas necessidades. Você pode ajustar os seletores XPath para se adequar ao site que deseja consultar e também personalizar o nome da planilha e as colunas na qual deseja armazenar os dados.

## Licença

Este projeto é licenciado sob a Licença MIT - veja o arquivo [LICENSE](LICENSE) para detalhes.

## Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para abrir problemas (issues) ou enviar pull requests com melhorias.

## Autor

Astral Oceane - [https://github.com/AstralOceane](https://github.com/AstralOceane)

## Agradecimentos

- Agradeço se quiser contribuir com melhorias ou idéias ou um estágio onde possa melhorar minhas habilidades :) no mais continuarei treinando realizando exercícios.
- Video onde peguei a ideia para o codigo https://www.youtube.com/watch?v=mhJp7npBtfk&ab_channel=DevAprender%7CJhonatandeSouza
- Agradecimentos ao meu amigo Bruno, que deu diversas ideias de melhorias para este projeto.
