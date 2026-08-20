# Painel SAPIEM — modalidades de aposentadoria por município

Ferramenta de consulta para o exame de legalidade das concessões: reúne as
modalidades parametrizadas no SAPIEM, os requisitos convertidos em dias, a
fundamentação legal com o texto de cada dispositivo e um conferidor que compara
os dados do processo contra a exigência da regra na data da DIB.

## Arquivos

```
Painel_SAPIEM.html          abra com dois cliques — é o painel
municipios/
  cachoeirinha.json         dados de um município
extrair_sapiem.py           gera o .json a partir da planilha + PDFs das normas
LEIA-ME.md                  este arquivo
```

Cachoeirinha já vem embutida no HTML e abre sem nenhuma configuração. Os demais
municípios entram pela aba **Municípios**, botão *Carregar município (.json)* —
o arquivo fica salvo no navegador e não precisa ser recarregado a cada sessão.

## Onde guardar

Pasta sincronizada do **OneDrive Business**, dentro do bloco *Referência
Jurídica*. Sugestão de estrutura:

```
Referência Jurídica/
  SAPIEM/
    Painel_SAPIEM.html
    municipios/
      cachoeirinha.json
      <outro-municipio>.json
    fontes/
      cachoeirinha/
        SAPIEM_Cachoeirinha_17jan23_rev__Eduardo.xlsx
        ELOM_01_2022.pdf
        Lei-complementar-83-2022-Cachoeirinha-RS.pdf
        relatorio_cachoeirinha.md
```

Guardar as fontes junto importa: quando a lei mudar, você regera o `.json` a
partir delas em vez de recomeçar.

**Não funciona** abrir o painel por um link do SharePoint ou pela visualização
web do OneDrive — esses ambientes não executam o JavaScript do arquivo. Precisa
ser o arquivo local, na pasta sincronizada. Se a Coordenadoria de TI puder
hospedar arquivos estáticos na intranet, aí sim vale migrar: vira um endereço
interno para toda a equipe, com atualização em um lugar só.

## Acrescentar um município

1. **Reúna as fontes.** A planilha do SAPIEM e, para cada norma, o PDF. Prefira
   sempre a **versão consolidada do LeisMunicipais** à do site da Câmara: tem
   camada de texto (extração automática) e já incorpora as alterações
   posteriores. PDF digitalizado não tem texto extraível — o script avisa, e aí
   o texto precisa ser transcrito e passado como `.txt`.

2. **Rode o script.**

   ```
   python extrair_sapiem.py planilha.xlsx \
       --aba "NomeDaAba" \
       --id nome-do-municipio \
       --municipio "Nome do Município" \
       --norma "LOM=emenda_lom.pdf" \
       --norma "LCM 00/0000=lei_complementar.pdf" \
       --versao-planilha "dd/mm/aaaa" \
       --saida municipios/nome-do-municipio.json
   ```

   Precisa de `openpyxl` e `pdfplumber`:
   `pip install openpyxl pdfplumber`

   O `--norma` aceita `.pdf`, `.txt` e `.md`, e pode ser repetido. A sigla à
   esquerda do `=` tem que ser **exatamente como aparece na coluna
   EMBASAMENTO_LEGAL** da planilha — é assim que o script casa cada modalidade
   com o artigo certo. Se o relatório disser "embasamento legal não
   reconhecido", quase sempre é a sigla.

3. **Leia o `relatorio_<id>.md`.** Ele traz a cobertura dos textos, as
   inconsistências internas da planilha e uma tabela com os requisitos de cada
   linha em anos, ao lado da base legal — é o material para você conferir
   contra a lei.

4. **Complete o `.json` à mão.** Quatro campos dependem de leitura e ficam
   vazios de propósito:

   - `progressao` — tabelas de idade mínima e pontuação das regras de
     transição, se o município as tiver;
   - `correlatos` — dispositivos úteis de consulta que nenhuma modalidade cita;
   - `achados` — o cruzamento da planilha com a lei, cada item com
     `situacao` igual a `confere`, `lacuna` ou `conferir`;
   - `vigenciaConferidaEm` — a data em que você checou se a norma sofreu
     alteração. Enquanto vazio, o painel exibe "vigência não conferida" na
     tarja superior.

   Use `municipios/cachoeirinha.json` como modelo.

5. **Carregue no painel**, aba Municípios.

## Formato do campo `progressao`

```json
"pontuacao_geral": {
  "rotulo": "Transição por pontuação — regra geral",
  "aplicaA": {"regra": "pontuacao", "excetoGrupo": "professor"},
  "idade":  {"ate": {"M": 61, "F": 56}, "apos": {"ano": 2024, "M": 62, "F": 57}},
  "pontos": {"base": {"M": 96, "F": 86}, "primeiroAcrescimo": 2024,
             "incremento": 1, "teto": {"M": 105, "F": 100}},
  "fonte": "Art. 92-B, V e §§ 1º, 2º e 3º, da LOM."
}
```

`aplicaA` seleciona a que modalidades a tabela se aplica, por `regra`, `grupo`
e `excetoGrupo`. O painel recalcula idade e pontos pelo ano da DIB e mostra a
`fonte` ao lado do resultado — a exigência aplicada fica sempre rastreável até
o dispositivo.

## Convenções

| Grandeza | Convenção |
|---|---|
| 1 ano | 365 dias |
| 1 mês | 30 dias |
| Idade | verificada por data-aniversário (nascimento + N anos ≤ DIB); o total em dias é informativo |
| Pontuação | (idade em dias + tempo de contribuição em dias) ÷ 365 |
| Marco de ingresso | até 31/12/2003 · até 20/06/2022 · posterior |

## Limites

O painel **não decide**. "Compatível" significa apenas que sexo, marco de
ingresso e requisitos de idade e tempo não excluem a modalidade — laudo
pericial, enquadramento de deficiência, efetiva exposição a agentes nocivos e
exercício em magistério continuam dependendo do exame dos documentos.

Os dados vêm da planilha do SAPIEM na versão informada, que pode estar
desatualizada em relação à lei. O campo `vigenciaConferidaEm` existe para essa
verificação não se perder.

O marco de ingresso de 20/06/2022 está fixo no código, por ser a data de
Cachoeirinha. Municípios com data de corte diferente exigem ajuste na função
`apurar()` do painel e em `montar_modalidades()` do script.
