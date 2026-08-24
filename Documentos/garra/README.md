# 🦅 Garra Master - Automação em Perícia Balística

O **Garra Master** é um motor de Inteligência Artificial e Fotogrametria desenvolvido em Python para automatizar, padronizar e acelerar a redação de Laudos de Perícia Criminal na área de Balística Forense.

Utilizando a API do Google Gemini, o script faz a triagem autônoma de ofícios e evidências, aplica visão computacional para extração de calibres e comprimentos de cano, e injeta os laudos diretamente em templates do Microsoft Word (`.docx`).

## ⚙️ Principais Funcionalidades
- **Triagem Inteligente (IA):** Separa automaticamente fotos de documentos de evidências.
- **Extração Criminológica Blindada:** Lê ofícios em PDF ou JPG, extraindo dados e limpando quesitos.
- **Fotogrametria Integrada:** Lê papel milimetrado para inferir calibres e dimensões de canos.
- **Proteção de Cadeia de Custódia:** Identifica a foto do invólucro original e gera um sumário do material.
- **Freio de Emergência:** Trava automática em caso de esgotamento de créditos da API.

## 🛠️ Como instalar e usar
1. Clone o repositório: `git clone https://github.com/SEU_USUARIO/garra-master-balistica.git`
2. Instale as dependências: `pip install pymupdf chromadb ollama python-docx google-genai pillow`
3. Substitua a chave da API no arquivo `servo.py`.
4. Organize os casos em pastas contendo o Ofício e as fotos das evidências e rode `python servo.py`.
