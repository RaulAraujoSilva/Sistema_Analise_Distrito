Com base no conteúdo técnico extraído dos documentos, seguem os conceitos teóricos sobre o Poder Calorífico Superior (PCS) e a Cromatografia Gasosa aplicada ao gás natural:

### Definição do PCS e Importância Comercial
O Poder Calorífico Superior (PCS) é a grandeza que representa a quantidade de energia contida em um determinado volume de gás. A sua determinação é fundamental para a conversão do volume medido (em metros cúbicos) para a quantidade de energia (em Joules, kWh ou MMBtu), visto que o faturamento e a comercialização do gás natural, bem como o pagamento de royalties, baseiam-se na energia entregue e não apenas no volume físico movimentado. A máxima "você fatura o que mede" reflete a necessidade de precisão nesta variável para garantir justiça nas transações entre produtores, transportadores, distribuidores e consumidores.

### Determinação via Cromatografia Gasosa em Linha
A metodologia padrão para a determinação do PCS e da composição do gás natural é a cromatografia gasosa. Para sistemas de medição com grandes volumes transacionados (tipicamente Categorias A e B), é obrigatório o uso de cromatógrafos em linha (instalados diretamente no processo).

O funcionamento teórico do cromatógrafo em linha envolve:
1.  **Amostragem:** Coleta de uma amostra representativa do fluxo de gás, condicionada para remover líquidos e partículas.
2.  **Separação:** A amostra é injetada em colunas de separação, impulsionada por um gás de arraste (como Hélio, Hidrogênio ou Nitrogênio). As colunas retardam seletivamente os componentes do gás com base em suas propriedades físico-químicas.
3.  **Detecção:** Um detector, tipicamente de Condutividade Térmica (TCD), identifica os componentes à medida que eluem da coluna, gerando um sinal elétrico proporcional à concentração de cada um.
4.  **Cálculo:** Um controlador processa os sinais para determinar a fração molar de cada componente.

### Relação entre PCS e Composição do Gás
O PCS não é medido diretamente, mas sim calculado inferencialmente a partir da composição química do gás determinada pelo cromatógrafo. A fórmula teórica para o cálculo do Poder Calorífico da amostra é dada pelo somatório do produto entre a concentração de cada componente ($x_i$) e o seu respectivo poder calorífico tabelado ($H_i$):

$$PCS_{amostra} = \sum (x_i \cdot H_i)$$

A composição típica do gás natural brasileiro, que define o seu PCS, apresenta as seguintes faixas de concentração molar:
*   **Metano ($CH_4$):** 60% a 90% (principal componente energético).
*   **Etano ($C_2H_6$):** 0% a 20%.
*   **Propano ($C_3H_8$):** 0% a 20%.
*   **Butano ($C_4H_{10}$):** 0% a 20%.
*   **Gases Inertes:** Dióxido de Carbono ($CO_2$, 0-8%) e Nitrogênio ($N_2$, 0-5%), que não contribuem para a energia e reduzem o PCS.

### Variação do PCS
O PCS do gás natural não é um valor estático; ele sofre variações ao longo do tempo (diárias e mensais). Essas oscilações ocorrem devido a alterações na composição do fluido proveniente dos campos de produção ou de processos de tratamento. O monitoramento contínuo dessas variações é essencial para o fechamento volumétrico e energético diário e mensal.

### Regulamentação Aplicável (INMETRO)
A qualidade do gás e os instrumentos utilizados para sua medição são regulados no Brasil pelo **Regulamento Técnico Metrológico (RTM)** aprovado pela **Portaria INMETRO nº 188/2021**. Esta portaria:
*   Estabelece os requisitos técnicos e metrológicos para cromatógrafos a gás em linha.
*   Aplica-se à medição fiscal, transferência de custódia, apropriação e controle operacional.
*   Exige que o analisador seja capaz de quantificar hidrocarbonetos leves (C1 a C5), gases inertes ($N_2$, $CO_2$) e a fração de hidrocarbonetos pesados ($C_6+$).
*   Determina que a calibração dos cromatógrafos seja realizada utilizando Material de Referência Certificado (MRC/gás padrão) em condições ambientais controladas.