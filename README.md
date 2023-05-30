# SOELIST
## Como dados de entrada são fornecidos:
  ### Banco de dados do SCADA, contendo o SUBNAM e o POINTNAM de todos os alarmes do SCADA;
  ### Neste banco de dados também encontram-se os textos para o estado OPEN (0) e CLOSE (1);
  ### Ao partir deste banco de dados, é criado um dataframe sostat, contendo:
    #### SUBNAM.POINTNAM STEXT0, STEXT1
  
  ###Também e fornecido um arquivo txt, chamado scratch.txt, contendo:
    #### Index, Start Time, Stop Time, Event, Description, Active
    #### A coluna Event deste txt, está truncada em 48 caracteres, preenchidos por espaços em branco caso necessário
    #### Devido a isso, um elemento da coluna Event pode conter mais que um correspondente no data frame sostat estruturado anteriormente
## Agora, como saída, é obtido o arquivo SOELIST.xlsx, contendo:
  ### Event Time, contendo a coluna Star Time do txt, só que no formato (dd/mm/yyyy hh:mm:ss.000)
  ### Precious Time é preenchido por valores em branco
  ### Status é preenchido com o STEXT0 caso a coluna Description do txt esteja Close e STEXT1 para OPEN
  ### Tagname contendo todos os correspondentes do txt no data frame sostat
  
## Ainda neste arquivo SOELIST.xlsx
  ### As linhas em vermelho representam os pontos não encontrados no banco de dados do SCADA
  ### As linhas em verde representam os pontos com mais de uma correspondencia no BD do SCADA
  ### Comentários
  ### Pontos com corresponência única no data frame sostat ou BD scada
