# Resume

Automação para leitura de arquivos de Nota de Corretagem do Banco Inter e CEI.
Para funcionar é preciso preencher o arquivo `configs.json`.

A pasta `files_inter` é onde vc coloca todas as notas de corretagem (arquivos .xls) baixadas do HomeBroker do Banco Inter.
A nota vem com o seguinte nome "&#95;NotaCor&#95;diamesano_seucodigo.xls".

A pasta `files_cei` é onde vc coloca todas as notas de corretagem (arquivos .xls) baixadas do CEI.

## config `inter_file` dentro de configs.json:

`cod_cli` = seu codigo que está no final do nome do arquivo de corretagem do inter

`read_file_by_file` = usado para ler arquivo por arquivo na forma como é listado. Ativando esse parâmetro. Atualmente você nao garante a veracidade dos valores finais, pois como alguns ativos tiveram desdobramentos, isso vai causar um wrong value (estamos corrigindo isso). Caso nao ative ele vai ler em ordem crescente, basta informar os respectivos years no `initial_year` e no `final_year`. 

`developments` = é onde se insere os desdobramentos que os ativos tiveram

## Usage

Basta inserir os arquivos nas respectivas pastas, preencher o arquivo `configs.json` e dar o play

## Contributing
Feel free to improve the project