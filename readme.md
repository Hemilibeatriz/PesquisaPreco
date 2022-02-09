Algoritmo de Automação de Processos desenvolvido em um ambiente virtual (Pycharm) para:

    A partir de uma planilha já cadastrada com os produtos que se quer verificar os precos (dentro de um limite de valores, tanto mínimo quanto máximo) e utilizando para esta verificação o navegador web com o Selenium, assim que o produto se enquadrar na faixa de valores escolhida, o link do mesmo será enviado para um e-mail específico.

    O arquivo possui:        

    pasta .idea = criado automáticamente pelo Pycharm na hora da criação do Ambiente Virtual

    main.py = arquivo a ser rodado

    buscas.xlsx = arquivo em que contém os produtos que eu quero pesquisar, os termos banidos (se tiver esse termo o código exclui este resultado, e os preços mínimo e máximo (que é a faixa de valores escolhida))

    Ofertas.xlsx = que é o resultado da execução do main.py, ela é atualizada a cada execução e nela são salvos os produtos identificados. Contendo o nome, valor e link.

    read.md = este arquivo aqui :)

Observação:
Você PRECISA do driver do Chrome (ChromeDriver) para que o Selenium possa funcionar bem. Para instalá-lo basta saber qual a versão do seu Google Chrome.

Para verificar qual a versão do navegador Google Chrome, clique no menu na barra de ferramentas Ajuda e selecione “Sobre o Google Chrome”. O número da versão atual aparecerá.

Pesquise na internet por "chromedriver Version 97.0.4692.71" (coloque a versão que voce encontrou na etapa anterior).

O primeiro link deverá ser este: https://chromedriver.chromium.org/downloads clique nele e escolha a versão correta.

Com o ChromeDriver baixado, escolha um local para estalação. Este local deve ver um PATH, para saber todos os seus PATHS vá no CMD e digite echo %path%

No meu caso, ele será instalado em C:\Users\hemil\AppData\Local\Programs\Python\Python37.

Sempre que der erro de versão no seu ChromeDriver, você deve atualiza-lo :)
