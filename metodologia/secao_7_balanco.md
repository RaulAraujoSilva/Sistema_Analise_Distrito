Seguem os conceitos teóricos fundamentais sobre o Balanço de Massa (ou Volumétrico) aplicados a distritos de distribuição de gás natural:

**Definição e Cálculo do Balanço**

O balanço em um distrito de distribuição consiste na comparação entre o volume total de gás que entra no sistema (geralmente através de um City-Gate ou Ponto de Entrega) e o somatório dos volumes medidos nos pontos de saída (consumidores e outras ramificações). A discrepância entre esses valores é quantificada através da diferença percentual, calculada pela seguinte fórmula:

$$Dif\% = \frac{\text{Volume de Entrada} - \Sigma(\text{Volumes de Saída})}{\text{Volume de Entrada}} \times 100$$

**Conceito de Bandas de Incerteza**

Para uma análise técnica rigorosa, não se devem comparar apenas os valores nominais medidos, uma vez que todo sistema de medição possui uma incerteza associada. Deve-se calcular a "banda de variação" ou intervalo de abrangência para os volumes de entrada e de saída.

Os limites dessa banda são definidos pelas seguintes equações, onde $V$ é o volume medido e $U\%$ é a incerteza expandida do sistema de medição:

*   **Volume Mínimo Probável:** $V_{min} = V \times (1 - U\%)$
*   **Volume Máximo Probável:** $V_{max} = V \times (1 + U\%)$

**Critério de Aceitação e Interpretação**

A validação do balanço baseia-se na sobreposição das bandas de incerteza da entrada e da saída.

1.  **Balanço Aceitável (Fechamento Metrológico):** Ocorre quando há interseção entre as faixas de valores prováveis da entrada e da saída (exemplo: o volume mínimo da entrada é menor ou igual ao volume máximo da saída).
    *   *Interpretação:* A diferença numérica encontrada pode ser explicada estatisticamente pelas incertezas intrínsecas aos instrumentos de medição. Não há evidência estatística de perdas físicas ou erros sistemáticos graves.

2.  **Balanço Não Aceitável (Diferença Significativa):** Ocorre quando as bandas de incerteza não se tocam (há um "gap" entre elas).
    *   *Interpretação:* A diferença encontrada é superior à soma das incertezas combinadas. Isso indica que a discrepância não é apenas metrológica, apontando para a existência de um problema real no sistema.

**Causas de Divergências e Ações Recomendadas**

Quando o balanço não fecha (as bandas não se sobrepõem), as causas prováveis incluem:

*   **Perdas Técnicas:** Vazamentos na rede de distribuição.
*   **Derivas Instrumentais:** Medidores (de entrada ou saída) operando fora da classe de exatidão ou com calibração vencida.
*   **Erros de Parametrização:** Configurações incorretas nos computadores de vazão (ex: densidade base, composição do gás).
*   **Condições Operacionais:** Falhas na compensação de pressão e temperatura (conversão PTZ) ou vazamentos em válvulas de segurança.

**Ações Recomendadas:**
Ao identificar um desequilíbrio que excede as bandas de incerteza, deve-se iniciar uma investigação que inclua a verificação de vazamentos físicos, a auditoria dos sistemas de medição (calibração e parametrização), a análise histórica dos volumes para identificar o momento do desvio e a inspeção das condições de instalação dos instrumentos.