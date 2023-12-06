[Versão em inglês](README.md)

Solução desenvolvida em **VBA** para **automatizar a atualização de uma planilha de controle de pagamentos** sobre um empreendimento.

## Como funciona
A solução funciona com a utilização de no mínimo duas planilhas: 
- a planilha de controle (a que receberá os dados);
- e a planilha que contém os dados de recebimento (nossa fonte de dados).

Ao abrir a planilha de controle, haverá na na primeira coluna e primeira linha da aba  **Clientes** o botão **Atualizar**, que ao ser acionado, irá requisitar o arquivo que contém os dados de pagamento (na demonstração, usamos o arquivo *RECEBIMENTO OUTUBRO DE 2022.xlsx*).

Selecionado o arquivo, o algoritmo irá identificar pelo nome das colunas as informações necessárias, como número do apartamento, data do pagamento e valor pago.
Essas informações então são transpostas para os devidos campos na própria aba **Clientes**.

Ao final da transposição, será criada uma nova aba contendo os dados colhidos e o status da transposição. Caso haja algum registro da planilha de origem que não foi transposta para a planilha de controle, o status indicará *'Unidade não encontrada'* ou *'Competência não encontrada'*.