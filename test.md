# Melhorias ao iniciar o jogo e ao executar script de download

Olá, eu realizei algumas alterações nos códigos manualmente.
Foi uma alteração no `GAMES_DATA`, contendo agora exemplos de jogo reais que eu posso testar.

Porém, eu alterei a tipagem de cada entrada no `GAMES_DATA`, agora é possível que um jogo tenha um executável, ou então uma lista de executáveis que podem ser executados para o jogo em questão pelo campo `executable_file`, como necessitamos do nome do processo para detectar que ele abriu ou não, fiz outra alteração em `process_name`, que pode receber um único nome de processo ou então uma lista de processos.

Adeque minha regra de negócio atual dos meus códigos para refletirem essa nova alteração e manter minhas demais funcionalidade intactas.

Também fiz uma nova alteração na lógica que faz a descompatação dos jogos que veem em partes (pieces.csv), iremos utilizar a tecnologia do 7zip como no comando a seguir `"%ProgramFiles%\7-Zip\7z.exe" x "G:\EAg\net-for-speed-subway-2\parts\net-for-speed-subway-2.zip.001" -o"G:\EAg\net-for-speed-subway-2\game"`, o exemplo não contem pastas reais apenas para efeito de exemplificação.

Em `latancy-guild-electron\src\installer\scripts` possui os scripts que são executados no instalador do meu programa gerado pelo electron (Inno Setup), seu código fonte pode ser localizado em `latancy-guild-electron\src\installer\installer.iss`. Crie dois novos scripts, um silencioso e outro com logs, igual ao utilizado pelo Netbird na mesma pasta de scripts, esses dois scripts faram o download do 7zip através do link `https://www.7-zip.org/a/7z2501-x64.msi`. E então faça atualização do meu arquivo `Installer.iss` com os novos passos da instalação para persistir o 7zip no sistema.
