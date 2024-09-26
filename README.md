# Gerador de Certificados

## Descrição

Este projeto é um gerador de certificados desenvolvido em Python, utilizando a biblioteca Tkinter para a criação de uma interface gráfica amigável. O programa permite que os usuários gerem certificados personalizados para alunos, baseando-se em dados importados de um arquivo Excel. Este gerador é um exemplo de RPA (Robotic Process Automation), que visa automatizar o processo de emissão de certificados, tornando-o mais eficiente e menos propenso a erros.

## Funcionalidades

- **Importação de Dados**: O sistema lê dados de um arquivo Excel que contém informações dos alunos, como CPF, nome, RG, datas de início e fim do curso, e e-mail.
- **Interface Gráfica**: Uma interface intuitiva que apresenta os dados em uma tabela (Treeview), permitindo fácil visualização e seleção de registros.
- **Edição de Campos**: Os dados selecionados na tabela podem ser facilmente transferidos para campos de entrada, permitindo que os usuários visualizem e editem informações antes de gerar certificados.
- **Filtragem de Dados**: Os usuários podem pesquisar e filtrar dados de alunos com base no CPF.
- **Geração de Certificados**: O sistema utiliza um modelo de documento Word para gerar certificados personalizados, substituindo campos marcados por informações do aluno.
- **Geração em Massa**: Permite a criação de vários certificados de uma só vez, economizando tempo e esforço na emissão de documentos.

## Tecnologias Utilizadas

- **Python**: A linguagem principal utilizada para o desenvolvimento do aplicativo.
- **Tkinter**: Biblioteca padrão do Python para criação de interfaces gráficas.
- **Pandas**: Biblioteca poderosa para manipulação e análise de dados, utilizada para ler e processar os dados do Excel.
- **python-docx**: Biblioteca para criar e modificar documentos Word, utilizada para gerar os certificados.

## Estrutura do Projeto

1. **Interface Gráfica**: Criada com Tkinter, onde são apresentadas as opções de pesquisa e os dados dos alunos em uma tabela.
2. **Funções de Manipulação**: Funções para importar dados, gerar certificados e lidar com eventos da interface gráfica, como cliques e entradas.
3. **Modelos de Documentos**: Um documento Word pré-formatado é usado como modelo para os certificados, permitindo a personalização com as informações dos alunos.
