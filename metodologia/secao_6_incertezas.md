Com base nos fundamentos teóricos e regulatórios apresentados no material, seguem os conceitos sobre Incertezas de Medição aplicados à infraestrutura de gás natural:

**Definição e Importância Regulatória**
A incerteza de medição é definida como um parâmetro que caracteriza a dispersão dos valores atribuídos a um mensurando, com base nas informações utilizadas. Diferentemente do erro (que é a diferença entre o valor medido e um valor de referência ou "verdadeiro"), a incerteza quantifica a dúvida associada ao resultado da medição. No contexto regulatório de óleo e gás, a estimativa da incerteza é fundamental para gerenciar o risco de medições incorretas e garantir a conformidade com o Regulamento Técnico de Medição (RTM), devendo ser calculada na fase de projeto e atualizada após calibrações ou alterações significativas no sistema.

**Classificação das Incertezas (GUM)**
A avaliação da incerteza segue as diretrizes do GUM (Guia para a Expressão de Incerteza de Medição) e classifica-se em dois tipos:
*   **Incerteza Tipo A:** É aquela avaliada por métodos estatísticos, expressa como um desvio padrão derivado de uma distribuição de probabilidade de valores medidos em uma série de observações repetidas.
*   **Incerteza Tipo B:** É aquela avaliada por meios que não a análise estatística de séries de observações. Baseia-se em conhecimentos prévios, como dados fornecidos em certificados de calibração, incertezas herdadas de padrões de referência, especificações técnicas de fabricantes, dados de medições anteriores ou constantes físicas de materiais.

**Combinação de Incertezas e Método RSS**
Para determinar a incerteza total de um sistema, deve-se combinar as incertezas padrão de todas as grandezas de entrada. Quando as fontes de incerteza são independentes (não correlacionadas), utiliza-se o método da Raiz da Soma dos Quadrados (RSS - *Root Sum Squares*). O modelo matemático envolve elevar ao quadrado a contribuição de cada incerteza componente, somá-las e extrair a raiz quadrada do resultado para obter a incerteza combinada.

**Incerteza Expandida e Fator de Abrangência**
A incerteza combinada representa um desvio padrão. Para fins práticos e regulatórios, utiliza-se a **Incerteza Expandida**, que define um intervalo dentro do qual se espera encontrar a maior parte da distribuição de valores que podem ser atribuídos ao mensurando.
Para obtê-la, multiplica-se a incerteza combinada por um **fator de abrangência ($k$)**. Conforme as normas aplicáveis ao setor, adota-se **$k=2$**, o que corresponde a um nível de confiança de aproximadamente **95%**.

**Limites de Referência e Aplicação**
A estimativa de incerteza deve considerar o pior cenário operacional (condições normais, mínimas e máximas). Os limites aceitáveis de incerteza e tolerância de erro variam conforme a criticidade da aplicação:
*   **Medição Fiscal e Transferência de Custódia:** Exigem maior rigor metrológico, tipicamente associados a limites de controle mais estreitos (ex: referência a 1% em contextos de falha presumida ou validação).
*   **Medição para Apropriação:** Permite tolerâncias ligeiramente maiores (ex: referência a 3% em contextos operacionais), dado o menor impacto financeiro direto comparado à transferência de custódia.

**Memorial de Cálculo para Medição de Gás**
O memorial de cálculo de incerteza é o documento formal que demonstra a adequação do sistema. O processo de elaboração deve seguir as etapas lógicas do GUM:
1.  **Definição do Modelo Matemático:** Estabelecer a equação que relaciona a vazão com as variáveis de entrada (ex: pressão diferencial, pressão estática, temperatura, densidade, fator do medidor).
2.  **Identificação das Fontes:** Listar todas as variáveis que contribuem para a incerteza (instrumentação, calibração, fluido, instalação).
3.  **Cálculo dos Coeficientes de Sensibilidade:** Determinar matematicamente o quanto a variação de cada grandeza de entrada influencia o resultado final da vazão.
4.  **Quantificação:** Calcular as incertezas padrão e combinada.
5.  **Resultado Final:** Apresentar a incerteza expandida da vazão (volumétrica ou mássica) em porcentagem e unidades da grandeza, declarando os graus de liberdade e o fator de abrangência utilizado.