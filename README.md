# planilhas
An basic helper to assists repart your multiple excel spreadsheets in a multiple excel files for better organization or factor.


Basically pom.xml load all code dependencies and SepararPlanilhasExcel the methods copiarConteudo and copiarValorCelula.

he was write for an specific file, to filter from than 2023 and ignore this rule for the header, copiarValorCelula was designed to bring the data faithfully while respecting the root file, basically a formatting class with an condition structure adjusting data.

In the main method we have a path for your file caminhoArquivoOriginal for the excel file you want to copy data and pastaDestino for new files destine path.

.rowCacheSize(100) & bufferSize(4096) was also defined to optimize and control how much code reads/writes from Excel, bringing accessible performance.


The rest is more of the same, logs, variables and instances.
