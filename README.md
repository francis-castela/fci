# Gerador de Agenda - Teatro Municipal de Itajaí

Aplicação web para montar agenda de eventos e exportar arte em PNG no formato retrato 1080x1440.

## Recursos

- Cadastro manual de eventos com 5 campos:
  - Dia/Mês
  - Horário
  - Nome do evento
  - Produtor
  - Onde comprar
- Pré-visualização em canvas com navegação por páginas.
- Exportação de:
  - Página atual
  - Todas as páginas (quando há mais eventos do que cabem em uma única imagem)
- Importação de planilha:
  - CSV
  - XLS
  - XLSX
- Download de modelo XLSX para preenchimento.
- Personalização de cores com opção de restaurar padrão.
- Persistência em cookies para continuar depois sem perder dados.

## Estrutura do projeto

```text
fci/
  index.html
  README.md
  assets/
    css/
      styles.css
    js/
      app.js
    fonts/
      Montserrat-VariableFont_wght.ttf
      Montserrat-Italic-VariableFont_wght.ttf
    img/
      logo-teatro.png
      logo-fundacao-cultural.jpg
```

## Como usar

1. Abra o arquivo `index.html` no navegador.
2. No Passo 1, defina o título da agenda (sempre convertido para maiúsculo).
3. Opcionalmente, ajuste cores no bloco "Personalizar cores".
4. No Passo 2, adicione eventos manualmente ou importe uma planilha.
5. No Passo 3, revise a lista de eventos.
6. No Passo 4, exporte a página atual ou todas as páginas.

## Importação de planilha

Padrão esperado de colunas:

- A: Dia/Mês
- B: Horário
- C: Nome do evento
- D: Produtor
- E: Onde comprar

Dicas:

- Cada linha representa um evento.
- Linhas incompletas são ignoradas.
- Pode usar o botão "Baixar modelo XLSX" para começar.

## Persistência de dados

O sistema salva automaticamente em cookies:

- Título
- Cores
- Eventos

Ao reabrir a página, os dados são restaurados automaticamente.

## Créditos

Desenvolvido por Francis Castela para uso interno do Teatro Municipal de Itajaí.
Versão 1.0 - 04/2026.
