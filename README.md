# 🔎 Sistema Automatizado de Monitoramento e Prospecção via Web Scraping com Interface Web e Cibersegurança

Este projeto em Python automatiza a coleta e o monitoramento de processos públicos disponíveis em páginas do sistema SEI (Sistema Eletrônico de Informações), com foco em licenciamento ambiental. A ferramenta foi desenvolvida com o objetivo de otimizar o acompanhamento de andamentos administrativos e facilitar a prospecção comercial inteligente, evitando o acesso manual repetitivo a dezenas (ou centenas) de páginas.

Com base em uma planilha de entrada contendo links públicos, o sistema realiza web scraping estruturado para extrair informações-chave, como:

Data da última atualização

Unidade responsável

Descrição do andamento

Interessado

Nome do empreendimento

# 🚀 Funcionalidades

Execução programada e automatizada com verificação em horários úteis e dias úteis (com filtro de feriados).

Coleta estruturada de dados, com uso de proxy (ScraperAPI) para garantir anonimato e mascaramento do IP, reduzindo riscos de rastreio.

Geração de relatórios em Excel com os dados extraídos de forma consolidada.

Envio automático de e-mails via Outlook, com corpo de mensagem personalizado e anexo do relatório em Excel.

Interface web (Flask) para visualização online dos dados extraídos, incluindo última verificação e relatórios.

Execução manual via navegador com apenas um clique para forçar a atualização dos dados.

Tratamento de erros e exceções com robustez para manter estabilidade em execução contínua.

# 🛡️ Cibersegurança Aplicada

Para aumentar a privacidade e a estabilidade do scraping, foi integrada uma infraestrutura com mascaramento de IP via ScraperAPI, garantindo que os acessos realizados ao sistema público não exponham o local real da requisição. Essa medida evita possíveis bloqueios ou rastreamentos indesejados, assegurando maior segurança para a aplicação.

# 📂 Infraestrutura Técnica

Backend: Python + Flask

Agendamento e Looping: threading, datetime, time

Scraping: requests, BeautifulSoup, ftfy, ScraperAPI

Manipulação de dados: pandas, openpyxl

Envio de e-mails: win32com.client (Outlook)

Interface web: HTML com render_template e rotas interativas

# ❗ Observação Importante
⚠️ O arquivo de Excel gerado com os dados reais não foi incluído neste repositório por tratar-se de propriedade intelectual da empresa responsável.
Este código é disponibilizado como uma alternativa segura e eficaz para consultas automatizadas em sistemas públicos, oferecendo uma base sólida para aplicação em outras realidades e instituições.
