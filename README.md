# Processador de Notas
---

Arquivo criado com o intuito de agilizar o processamento das notas do simulado e gerar um arquivo de texto que alimentará o **Escolar Manager**. Para que ele funcione, algumas diretrizes deverão ser seguidas:

## 1. Formato da Planilha de Origem

Por ser a versão 1.0 do programa, muito refinamento ainda precisa ser feito, principalmente rem relação à base de dados, por esse motivo estou utilizando Excel. Nesse sentido, ele está editado de maneira específica para que possa ser lido pelo processador de notas.
- O arquivo Excel precisa ter 4 planilhas: `1ª Série`, `2ª Série`, `3ª Série` e `Lista de Alunos`;
- As planilhas de `Série` precisam ter as notas gerais do simulado na `coluna N`;
- Os nomes dos `alunos` precisam iniciar na `linha 7`;
- A lista de alunos contém somente duas colunas: `número de matrícula` na coluna `A` e `Nome Completo do Aluno` na `coluna B`.

> **Detalhe importante** A planilha **Lista de Alunos** não precisa de nenhum cabeçalho, e as informações devem começar a ser listadas na **linha 1**.

## 2. Formato de Saída

O arquivo de saída é um `.txt` com o título que é uma combinação entre a série escolhida e o nome dado pelo digitador. Ele estará disponível no caminho `...\Meus Arquivos\Documents\resultados`

## 3. Dependências
1. python 3.x
2. tkinter
3. tkinterbootstrap
4. os
5. re
6. threading
7. sys
8. subprocess
9. pandas

---


> Mais informações serão adicionadas conforme novas atualizações forem sendo feitas