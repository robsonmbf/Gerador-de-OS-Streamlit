# 📋 Gerador de Ordem de Serviço - NR01

Sistema web para geração automatizada de Ordens de Serviço conforme Norma Regulamentadora NR-01, desenvolvido em Python com Streamlit.

## 🚀 Funcionalidades

- ✅ **Upload de Planilha**: Carregue sua planilha Excel com dados dos funcionários
- ✅ **Validação Automática**: Verifica se a estrutura da planilha está correta
- ✅ **Interface Intuitiva**: Design moderno e responsivo
- ✅ **Filtragem Inteligente**: Filtre funcionários por setor
- ✅ **Personalização**: Configure riscos e medidas específicas para cada função
- ✅ **Geração Automatizada**: Cria documentos Word (.docx) profissionais
- ✅ **Conformidade NR-01**: Segue todos os requisitos da norma regulamentadora
- ✅ **Download Instantâneo**: Baixe as OS geradas imediatamente

## 📋 Estrutura da Planilha

A planilha Excel deve conter as seguintes colunas obrigatórias:

| Coluna | Descrição |
|--------|-----------|
| `Nome` | Nome completo do funcionário |
| `Setor` | Setor/departamento de trabalho |
| `Função` | Cargo/função exercida |
| `Data de Admissão` | Data de admissão na empresa |
| `Empresa` | Nome da empresa |
| `Unidade` | Unidade/filial |
| `Descrição de Atividades` | Descrição detalhada das atividades |

## 🛠️ Tecnologias Utilizadas

- **Python 3.8+**
- **Streamlit** - Framework para aplicações web
- **Pandas** - Manipulação de dados
- **Python-docx** - Geração de documentos Word
- **OpenPyXL** - Leitura de planilhas Excel

## 📦 Instalação Local

```bash
# Clone o repositório
git clone https://github.com/robsonmbf/Gerador-de-OS-Streamlit.git
cd Gerador-de-OS-Streamlit

# Instale as dependências
pip install -r requirements.txt

# Execute a aplicação
streamlit run app_os_generator.py
```

## 🌐 Deploy Online

A aplicação está disponível online no Streamlit Cloud:
**[🔗 Acesse o Sistema](https://gerador-os.streamlit.app)**

## 📖 Como Usar

### 1. Prepare sua Planilha
- Baixe o modelo de exemplo no sistema
- Preencha com os dados dos seus funcionários
- Salve no formato Excel (.xlsx)

### 2. Carregue os Dados
- Faça upload da planilha no sistema
- O sistema validará automaticamente a estrutura
- Visualize as estatísticas dos dados carregados

### 3. Configure a OS
- Selecione o funcionário desejado
- Configure os riscos ocupacionais específicos
- Defina as medidas de proteção e EPIs
- Adicione informações sobre avaliações ambientais

### 4. Gere o Documento
- Clique em "Gerar Ordem de Serviço"
- Baixe o documento Word gerado
- O arquivo seguirá o padrão da NR-01

## 📄 Conteúdo da Ordem de Serviço

O documento gerado inclui:

- ✅ **Identificação** do funcionário e empresa
- ✅ **Tarefas da função** detalhadas
- ✅ **Agentes de riscos ocupacionais** (físicos, químicos, biológicos, ergonômicos, acidentes)
- ✅ **Possíveis danos à saúde** relacionados aos riscos
- ✅ **Medidas de prevenção e controle** dos riscos
- ✅ **EPIs obrigatórios** para a função
- ✅ **Procedimentos de emergência** e primeiros socorros
- ✅ **Orientações sobre grave e iminente risco**
- ✅ **Aspectos legais** conforme CLT e NR-01

## ⚖️ Conformidade Legal

O sistema gera documentos em total conformidade com:
- **NR-01** - Disposições Gerais e Gerenciamento de Riscos Ocupacionais
- **Art. 158 da CLT** - Obrigações do empregado
- **Portaria SEPRT nº 6.730/2020** - Aprovação da NR-01

## 🤝 Contribuições

Contribuições são bem-vindas! Para contribuir:

1. Fork o projeto
2. Crie uma branch para sua feature (`git checkout -b feature/MinhaFeature`)
3. Commit suas mudanças (`git commit -m 'Add: MinhaFeature'`)
4. Push para a branch (`git push origin feature/MinhaFeature`)
5. Abra um Pull Request

## 📝 Licença

Este projeto está sob a licença MIT. Veja o arquivo [LICENSE](LICENSE) para mais detalhes.

## 🆘 Suporte

Para dúvidas ou problemas:
- Abra uma [Issue](https://github.com/robsonmbf/Gerador-de-OS-Streamlit/issues)
- Entre em contato através do GitHub

## 📊 Status do Projeto

- ✅ **Versão Atual**: 2.0
- ✅ **Status**: Em produção
- ✅ **Última atualização**: Setembro 2025

## 🔄 Próximas Funcionalidades

- [ ] Integração com banco de dados
- [ ] Histórico de OS geradas  
- [ ] Templates personalizáveis
- [ ] Assinatura digital
- [ ] Notificações por email
- [ ] Relatórios gerenciais

---

Desenvolvido com ❤️ para facilitar a gestão de segurança do trabalho nas empresas brasileiras.