# 🛠️ PQ Web Simulator

**Simulador de Power Query para Excel Web (Add-in Gratuito)**

O **PQ Web Simulator** é um Add-in desenvolvido para suprir a ausência de ferramentas avançadas de transformação de dados no Excel Web. Ele permite realizar operações comuns do Power Query (como manter colunas, filtrar, dividir texto e transformar casos) diretamente na interface web, utilizando a API JavaScript do Excel.

Este projeto foi idealizado para auxiliar analistas de dados e entusiastas que trabalham exclusivamente no ambiente online, oferecendo uma alternativa leve, acessível e com uma experiência de usuário (UX) aprimorada com feedback visual e sonoro.

## 🚀 Funcionalidades

- **✂️ Manter Colunas:** Seleciona e extrai colunas específicas da tabela ativa usando a função dinâmica `CHOOSECOLS`.
- **🔍 Aplicar Filtro:** Filtra linhas com base em critérios exatos de coluna e valor, projetando o resultado em uma nova área.
- **➗ Dividir Texto:** Separa conteúdos de células selecionadas por delimitadores personalizados.
- **⬆️⬇️ Transformar Texto:** Converte rapidamente seleções para MAIÚSCULA ou minúscula.
- **🎨 Design System Adaptativo:** Interface moderna com **10 temas visuais** (Dark Tech, Cyberpunk, Royal Gold, Ocean, Sunset, Forest, Light, Midnight, Cherry e Slate) e persistência de preferência via LocalStorage.
- **🔊 Feedback Sonoro UI:** Sistema de áudio sutil para confirmação de ações (sucesso), erros e interações (hover), elevando a experiência do usuário.
- **✨ Animações CSS:** Título com efeito de brilho dinâmico (que muda de cor conforme o tema) e ícone de ferramenta com rotação contínua.

## 💻 Tecnologias Utilizadas

- **Core:** HTML5, CSS3 (Variáveis CSS, Flexbox/Grid, Keyframes), JavaScript (ES6+)
- **API:** Microsoft Office.js (Excel Web Add-ins)
- **Deploy:** GitHub Pages
- **Versionamento:** Git & GitHub

## 📦 Instalação e Uso

Como este é um projeto hospedado no GitHub Pages, você pode utilizá-lo diretamente no Excel Web:

1. Acesse o [Excel Online](https://www.office.com/launch/excel).
2. Vá na guia **Página Inicial** > **Suplementos** > **Meus Suplementos** (ou "Obter Suplementos").
3. Selecione a aba **URL Personalizada** (ou "Desenvolvedor", dependendo da versão).
4. Insira o link do manifesto: `https://dw-dev-cloud.github.io/simulador-web-pq/manifest.xml`
5. O painel lateral **"PQ Web Simulator"** será carregado automaticamente.

## 🏗️ Estrutura do Projeto

```text
pq-web-simulator/
├── index.html          # Estrutura principal e injeção do Office.js
├── taskpane.css        # Design system com 10 temas, variáveis CSS e animações
├── functions.js        # Lógica de negócios, handlers de eventos e sistema de áudio
├── manifest.xml        # Configuração do Add-in para o Microsoft Office
├── error.mp3           # Feedback sonoro de erro
├── success.mp3         # Feedback sonoro de sucesso
└── hover.mp3           # Feedback sonoro de hover