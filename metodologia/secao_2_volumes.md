Com base no conteúdo técnico extraído do documento, seguem os conceitos teóricos sobre Medição Volumétrica e Transferência de Custódia em gás natural:

### Fundamentos da Medição Fiscal e Transferência de Custódia

**Conceito de Transferência de Custódia**
A transferência de custódia descreve a mudança de propriedade que ocorre quando produtos, neste caso o gás natural, são comprados e vendidos. Este processo é regido por termos contratuais e regulamentações, onde comprador e vendedor devem cumprir obrigações predeterminadas para garantir a precisão e transparência das transações.
Existem dois tipos principais:
1.  **Comercial:** Ocorre quando as obrigações são estabelecidas em contrato particular e os aspectos metrológicos são negociados entre as partes.
2.  **Legal:** Ocorre quando a transação é regulamentada por órgãos oficiais (como ANP e INMETRO), atendendo a interesses da sociedade e requerendo aprovação metrológica formal.

**Medição Fiscal**
A medição fiscal é definida como uma combinação de regulamentos, protocolos, leis, sistemas e dispositivos que habilita duas partes a concordarem com o valor do produto transferido. Ela é aplicada em transações de grandes volumes, produtos de alto valor agregado e produtos estratégicos. No Brasil, o regulamento técnico determina que os sistemas devem gerenciar o risco de medições incorretas através de um sistema de gestão da medição.

### Medição Volumétrica e Unidades de Referência

Embora a medição de vazão possa ser realizada através de tecnologias mássicas (como Coriolis) ou volumétricas (como Turbina, Ultrassom ou Placa de Orifício), a legislação brasileira exige a totalização volumétrica nas condições de referência.

**Unidades e Condições de Referência**
Para que os volumes de gás natural sejam comparáveis e utilizados em transações comerciais, o volume medido nas condições de operação (pressão e temperatura da linha) deve ser convertido para uma condição base padronizada.
*   **Vazão Não Corrigida:** Medição realizada nas condições reais de operação (pressão e temperatura de escoamento).
*   **Vazão Corrigida (Standard):** Volume convertido para as condições de referência.
*   **Condições de Referência no Brasil:** Pressão base de **101,325 kPa (1 atm)** e Temperatura base de **20 °C**.

A conversão do volume operado para o volume corrigido utiliza a seguinte relação teórica, aplicada pelos computadores de vazão:

$$V_{ref} = V_{op} \times F_P \times F_T \times F_{pv} \times F_M$$

Onde:
*   $V_{ref}$: Volume nas condições de referência.
*   $V_{op}$: Volume nas condições de operação.
*   $F_P$: Fator de correção da pressão.
*   $F_T$: Fator de correção da temperatura.
*   $F_{pv}$: Fator de compressibilidade ($1/Z$, corrigindo o comportamento não ideal do gás).
*   $F_M$: Fator de correção do medidor (curva de calibração).

### Fatores que Afetam a Medição Volumétrica

A precisão da medição volumétrica é influenciada diretamente por variáveis físicas e composicionais:

1.  **Pressão e Temperatura:** Variações nestas grandezas alteram a densidade do gás. A medição precisa destas variáveis é fundamental para a correção do volume para as condições de referência.
2.  **Qualidade e Composição do Gás:** A determinação da composição (via cromatografia em linha ou laboratório) é essencial para o cálculo da densidade relativa, do Fator de Compressibilidade ($Z$) e do Poder Calorífico (necessário para a determinação da energia entregue).
3.  **Regime de Escoamento:** O perfil de velocidade do fluido (laminar, turbulento ou transição), definido pelo Número de Reynolds, afeta a exatidão de medidores lineares e deprimogênios.
4.  **Ruído e Instalação:** Ruídos sônicos (gerados por válvulas reguladoras) e distorções no perfil de fluxo afetam tecnologias como medidores ultrassônicos e placas de orifício, exigindo trechos retos adequados ou condicionadores de fluxo.

### Critérios de Concordância e Tolerâncias (Falha Presumida)

Para garantir a confiabilidade metrológica, são estabelecidos limites de tolerância para a validação dos sistemas de medição. O conceito de "Falha Presumida" é utilizado para determinar quando um medidor não está mais apto para uso sem manutenção ou nova calibração.

**Limites de Tolerância para Deriva (Drift):**
*   **Medição Fiscal e Transferência de Custódia:** A deriva do medidor não deve exceder **1%** em valor absoluto.
*   **Medição de Apropriação:** A tolerância é de até **3%** em valor absoluto.
*   **Medidores Padrão de Referência:** O desvio máximo entre calibrações sucessivas não deve ser maior que **0,02%**.

Caso a deriva ultrapasse esses limites, o medidor é considerado em falha presumida e deve ser retirado de operação para manutenção e recalibração. Além disso, a incerteza da medição deve ser estimada e mantida dentro dos requisitos regulatórios, considerando todas as fontes de erro (instrumentação, condições ambientais e fluido).