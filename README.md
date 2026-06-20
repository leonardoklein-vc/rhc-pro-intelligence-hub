# 🧬 RHC Pro v2: Enterprise Clinical Data Hub & BI

![Architecture](https://img.shields.io/badge/Architecture-Serverless_SPA-blue?style=for-the-badge)
![Pattern](https://img.shields.io/badge/Design_Pattern-Library%2FShell-8b5cf6?style=for-the-badge)
![Google Apps Script](https://img.shields.io/badge/Backend-Google_Apps_Script-4285F4?style=for-the-badge&logo=google)
![JavaScript](https://img.shields.io/badge/Frontend-Vanilla_JS-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black)
![DataViz](https://img.shields.io/badge/DataViz-Chart.js-FF6384?style=for-the-badge)

RHC Pro v2 é uma SPA serverless no Google Workspace para governança de dados clínicos e BI Oncológico. Transforma dezenas de planilhas em um Data Hub central com auditoria automatizada, filtros em milissegundos via Smart Cache e renderização de relatórios PDF in-browser. Escalabilidade total com arquitetura Library/Shell.

---

## 🚨 O Desafio Operacional (O Caos Analógico)
O fluxo do Registro Hospitalar lidava com uma carga de dados oncológicos complexa e descentralizada. 
* **Fragmentação Extrema:** As etapas de trabalho não conversavam entre si. Dados eram organizados em dezenas de guias isoladas.
* **Lentidão Analítica:** Levantamentos de KPIs e indicadores de saúde (como a Lei dos 60 Dias) demoravam semanas de mineração manual.
* **Falta de Memória Institucional:** Nenhuma rastreabilidade das ações tomadas, dificultando a governança clínica e a auditoria de erros.
<img width="1804" height="872" alt="ChatGPT Image 20 de jun  de 2026, 17_57_54" src="https://github.com/user-attachments/assets/9f081a57-8189-47ad-8c1e-1d16de115192" />

---

## 💡 A Arquitetura da Solução (Modular & Serverless)

Para garantir manutenção ágil e escalabilidade, o ecossistema foi dividido em três camadas distintas, espelhadas na estrutura deste repositório:

1. **RHC Core (Library):** O motor de regras de negócio. Um script centralizado que as outras planilhas consomem como biblioteca. Atualizações aqui refletem instantaneamente em toda a operação.
2. **Client Shells (Queries):** Planilhas operacionais que atuam apenas como "terminais burros" (Cascas), conectando-se à biblioteca central via funções de ponte.
3. **RHC Pro Hub (V2):** O painel executivo central, abrigando o Gestor de Bases, Repositório de POPs e o poderoso motor de Business Intelligence.

<img width="1014" height="585" alt="Captura de tela 2026-06-20 180254" src="https://github.com/user-attachments/assets/8eb5d33c-8fa3-48c2-998b-64816ea79840" />
---

## ✨ Módulos e Engenharia de Destaque

### 📊 1. Motor de Business Intelligence (Client-Side Rendering)
Contornar o limite de tempo de execução dos servidores do Google exige mover o peso do processamento para a máquina do usuário.
* **Smart Caching (V2):** Os dados mastigados pelo servidor são injetados no `localStorage` do navegador. Trocar anos de análise ou aplicar filtros cruzados complexos (Idade + Etnia + Grupos Topográficos CID + Desfecho Clínico) acontece em milissegundos, sem recarregar a página.
* **DataViz Dinâmico:** Integração profunda com `Chart.js` para renderizar gráficos de Evolução Histórica e Radares de Proporcionalidade demográfica oncológica.
* **PDF In-Browser Engine:** Um motor de relatórios que clona o DOM, injeta CSS otimizado para impressão e utiliza `html2pdf.js` para compilar relatórios executivos direto na memória RAM do usuário, sem consumir cota de servidor.* 

<img width="1819" height="865" alt="ChatGPT Image 20 de jun  de 2026, 18_08_07" src="https://github.com/user-attachments/assets/dd04a0c1-2c24-484e-99b1-1c334e5aab14" />


### 🛡️ 2. Auditoria Contínua e "Máquina do Tempo"
Para garantir que nenhum paciente oncológico seja "perdido" no fluxo de trabalho mensal:
* **Cruzamento Hierárquico:** O robô varre o Google Drive recursivamente, mapeando bases ativas, inativas e *Queries Brutas*.
* **Tolerância Temporal Matemática:** O algoritmo converte meses e anos em índices absolutos `(Ano * 12) + Mês`. Ele busca pacientes faltantes justificando-os matematicamente se foram tratados no passado (-12 meses) ou lançados no futuro (+2 meses), isolando cirurgicamente as reais perdas para ação de compliance.

### 🏥 3. Gestor Clínico (Workflow de Registro)
Uma interface UI injetada nas planilhas de *Queries* (via Library) para padronizar o *data entry* das equipes.
* Preenchimento guiado do Estadiamento (TNM Patológico), Razões de Não Tratamento e Clínica de Início.
* **Sistema de Bloqueio (Soft-Locking):** Salva estados globais da linha em atendimento, impedindo conflitos de concorrência e sobrescrita de dados já auditados.*
<img width="1070" height="551" alt="Captura de tela 2026-06-20 181500" src="https://github.com/user-attachments/assets/e5cae3dc-e6df-4bc5-a475-b957351d221d" />


### 💾 4. Central de Backups e Cofre de Snippets
* **Rotina Automática ZIP:** Algoritmos em background consolidam os arquivos do Drive em lotes ZIP organizados por ano/sessão, gerenciando automaticamente a retenção de 30 dias.
* **Repositório Interno:** Uma IDE embutida no próprio painel onde a equipe técnica salva *Snippets* e scripts úteis de consulta SQL/GAS diretamente no banco local.
<img width="891" height="475" alt="Captura de tela 2026-06-20 181717" src="https://github.com/user-attachments/assets/6c027e8e-25db-4ead-9f57-4fadfab0a661" />

---

## 🎨 UI/UX: Theming Engine Corporativo
A interface da SPA foi projetada para reduzir a fadiga cognitiva de operadores com um motor de temas via Variáveis CSS. Permite alternar em tempo real entre mais de 15 esquemas visuais (Dark Mode, Light Clean, Alto Contraste para acessibilidade e as paletas oficiais da instituição).

---

## 🚀 Stack Tecnológico
* **Core Backend:** Google Apps Script (ES6+ V8 Engine).
* **Banco de Dados:** Google Sheets (como BD Relacional) + Google Drive API.
* **Frontend:** HTML5, CSS3, Vanilla JS (Zero frameworks pesados, foco absoluto em performance).
* **Bibliotecas Client-Side:** `Chart.js` (Visualização) e `html2pdf.js` (Exportação Documental).
