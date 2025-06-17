# Web.Scraping
Criei um script em Python que realiza web scraping diretamente nessas páginas públicas. A ferramenta lê uma planilha com os links, acessa cada página automaticamente, coleta as informações mais relevantes e organiza tudo de forma estruturada.

No meu dia a dia, sempre me deparei com um desafio: acompanhar o andamento de diversos processos administrativos por meio de links que levavam a páginas de um site público. A cada novo processo, era necessário acessar manualmente a página, identificar a data da última atualização, a unidade responsável, a descrição do andamento e ainda verificar quem era o interessado. Esse trabalho, repetitivo e sujeito a erros, consumia um tempo precioso.

Esses dados são transformados em um relatório em Excel, pronto para ser consultado. E para completar a automação, o sistema ainda prepara e envia um e-mail formal, com texto personalizado, relatório anexado e a identidade visual da organização integrada — tudo de forma automática via Outlook.

💡 Com essa solução, deixei de gastar horas com tarefas repetitivas e passei a entregar informações confiáveis, atualizadas e bem apresentadas em questão de minutos.

Essa automação surgiu de uma dor real e mostrou como a tecnologia, aplicada com foco e simplicidade, pode transformar uma rotina pesada em um processo inteligente e eficiente.

📈 Hoje, acompanhar processos deixou de ser esforço — virou estratégia.


# Tecnologias e bibliotecas utilizadas: 🧰

Python – linguagem principal de desenvolvimento

pandas – leitura, manipulação e exportação de planilhas Excel

requests – requisição HTTP para acessar páginas públicas

BeautifulSoup (bs4) – extração e parsing de conteúdo HTML

datetime – manipulação e formatação de datas

time – controle de intervalo entre requisições (delay)

matplotlib.pyplot – suporte visual (se necessário para gráficos)

win32com.client – automação de envio de e-mails via Outlook

os – verificação de arquivos e caminhos no sistema
