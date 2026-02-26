📂 DOCUMENTAÇÃO TÉCNICA E ARQUITETURA
Sistema de Movimentações de Headcount
1. Visão Geral do Sistema
Aplicação Web Cloud-Native desenvolvida para modernizar e auditar as trocas e requisições de vagas entre centros de custos. O sistema elimina o uso de planilhas descentralizadas, centralizando os dados em um banco relacional em nuvem e integrando-se nativamente com o Microsoft Power BI via API REST.

2. Stack Tecnológico e Lógica de Construção
O projeto foi desenvolvido sob a ótica da Engenharia de Dados moderna:

Frontend (Interface): Python utilizando o framework Streamlit.

Backend (Banco de Dados): Supabase (PostgreSQL). Escolhido por ser um banco de dados relacional em nuvem altamente escalável.

Módulo de Notificações: Automação em Python (smtplib e email.mime) para disparo assíncrono de e-mails corporativos.

Gestão de Parâmetros: Leitura em cache de tabelas locais para otimização de memória do servidor.

3. Parâmetros de Conexão e Credenciais (Data Vault)
Abaixo estão os dados técnicos da infraestrutura. (Nota: Senhas reais estão omitidas neste documento por política de segurança da informação).

A. Banco de Dados (Supabase - PostgreSQL API)

Projeto ID: pthrbdtrvwboagzishee

URL Base da API (REST): https://pthrbdtrvwboagzishee.supabase.co/rest/v1/

Tabelas Endpoint: movimentacoes e solicitacoes_postos

Parâmetro de Consulta: ?select=* (Retorna todas as colunas)

Autenticação (Headers):

apikey: sb_publishable_S8aJBUpboihn_6JC_biwBQ_jxUV3HT3

Authorization: Bearer sb_publishable_S8aJBUpboihn_6JC_biwBQ_jxUV3HT3

B. Serviço de Disparo de E-mails (SMTP Google)

Servidor SMTP: smtp.gmail.com | Porta: 587 (TLS)

Remetente Autenticado: kamilacrisc@gmail.com

4. Segurança e Governança
Nenhuma credencial de banco de dados, senhas de usuários corporativos ou chaves de e-mail estão expostas no código-fonte. Todas as variáveis sensíveis são injetadas em tempo de execução através do cofre criptografado Streamlit Secrets. O repositório de código permanece público para fins de portfólio, mas totalmente blindado contra acessos indevidos.

5. Integração com Power BI (Método API REST - Bypassing SSL)
Para evitar falhas de validação de certificado de segurança (SSL) exigidas nativamente pelo Power BI em bancos PostgreSQL na nuvem, a arquitetura utiliza a API REST do Supabase. Este método garante atualizações automáticas Cloud-to-Cloud sem necessidade de Gateways físicos.

Fase 1: Extração e Transformação (Power BI Desktop)
Utilizar o conector Web no Power BI.

Selecionar a opção Avançado.

URL da Parte: Inserir o endpoint completo da tabela desejada (ex: https://pthrbdtrvwboagzishee.supabase.co/rest/v1/movimentacoes?select=*).

Cabeçalhos HTTP (Parâmetros de Segurança):

Adicionar cabeçalho 1: apikey = [Sua_Chave_Publishable]

Adicionar cabeçalho 2: Authorization = Bearer [Sua_Chave_Publishable]

No Power Query:

A extração retornará uma List (Lista).

Clicar em Para Tabela (To Table) > OK.

Clicar no ícone de expansão no cabeçalho da coluna Column1.

Desmarcar a opção "Usar o nome da coluna original como prefixo" e confirmar.

Fase 2: Configuração de Atualização (Power BI Service / Nuvem)
Após publicar o relatório no Workspace web, é obrigatório reconfigurar as credenciais para burlar o teste nativo da Microsoft que não envia os cabeçalhos HTTP.

Acessar as Configurações do Modelo Semântico (Dataset).

Expandir a aba Credenciais da fonte de dados e clicar em Editar credenciais.

Configuração Exata Exigida:

Método de Autenticação: Anônimo (As senhas já estão embutidas no código do Power Query).

Nível de Privacidade: Organizacional.

⚠️ Checkbox Obrigatório: Marcar a opção Ignorar conexão de teste (Skip test connection). Sem esta marcação, a Microsoft tentará validar o link puro e retornará erro 400 (Bad Request).

Clicar em Entrar / Aplicar.

Configurar os horários na aba Atualizar (Scheduled Refresh). A partir deste momento, o painel está autônomo na nuvem.
