# Copiador automatico

 Programa que pega os dados de vários arquivos .txt e os copia em planilhas de vários arquivos .xlsx.

 Criado por Pedro Reis (pedropcr.42@gmail.com)

Antes de executar o script, você precisa baixar o módulo openpyxl do Python

 Orientações para uso:
*Este programa copia automaticamente dados de arquivos no formato .txt para arquivos no formato .xlsx;
*Você precisa criar arquivos com 'nomes base' seguidos de um número sequencial, ou seja, todos os arquivos de mesmo formato devem ter o mesmo\nnome e variar apenas no número;

Por exemplo: suponha que você tenha vários arquivos xlsx para preencher com números de telefone.
Em seu diretório, crie um arquivo com uma planilha modelo, e o copie até dar o número de arquivos que precisa.
Nomeie os arquivos do seguinte modo: telefones01.xlsx, telefones02.xlsx e assim por diante.
Faça o mesmo com os arquivos .txt.

É importante que nos arquivos txt os dados estejam 'juntos', isto é, que estejam um embaixo do outro sem linhas em branco. 
Este programa irá pegar os dados linha por linha, por isso os organize com antecedência nos arquivos de textos de modo que o programa só precise pegar os dados e copiá-los nas células fornecidas.

*Com os passos acima concluídos, basta informar o menor número sequencial (ponto de início) e o maior número (fim) dos arquivos xlsx. Vale lembrar que o arquivo do número que marca o fim também será preenchido!

*Para um exemplo mais prático, use os arquivos 'exemplo' que vem junto com o repósitório.
Quando executar o script,

Em 'nome base' do arquivo .xlsx digite: exemploPlanilha

Em 'nome base' do arquivo .txt digite: exemploDados

Em 'menor número' digite: 1

Em 'maior número' digite: 3

Em 'número arquivo txt' digite: 1

Em 'letras das colunas' digite: b d

Em 'números das linhas' digite: 1 2 3 4 5 6


Pronto! Depois disso, abra os arquivos .xlsx e veja se o que está escrito nas células não são os mesmos dados dos arquivos .txt!
