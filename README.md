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
  - Todas as páginas (em arquivo ZIP quando há mais de uma página)
- Importação de planilha:
  - CSV
  - XLS
  - XLSX
- Download de modelo XLSX para preenchimento.
- Personalização de cores com opção de restaurar padrão.
- Persistência automática em localStorage para continuar depois sem perder dados.

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
    vendor/
      xlsx.full.min.js (opcional, fallback local)
      jszip.min.js (opcional, fallback local)
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

O sistema salva automaticamente em localStorage:

- Título
- Cores
- Eventos

Ao reabrir a página, os dados são restaurados automaticamente.

## Bibliotecas e fallback

O projeto usa carregamento dinâmico de bibliotecas para planilhas e ZIP:

- XLSX (importação e template)
- JSZip (exportação de múltiplas páginas em um único .zip)

Ordem de carregamento:

1. Arquivo local em `assets/vendor/`
2. CDN (jsDelivr)

Se quiser funcionamento offline completo, adicione os arquivos:

- `assets/vendor/xlsx.full.min.js`
- `assets/vendor/jszip.min.js`

## Checklist rápido de testes

- Criar 1 evento e exportar página atual em PNG.
- Criar eventos suficientes para 2+ páginas e exportar todas (ZIP).
- Importar um CSV válido e conferir contagem de eventos.
- Importar arquivo com linhas incompletas e validar aviso de linhas ignoradas.
- Baixar o modelo XLSX.
- Recarregar a página e validar restauração de título, cores e eventos.
- Alternar tema claro/escuro e validar persistência.

## Troubleshooting

- Importação/modelo XLSX não funciona:
  - Verifique conexão com internet ou adicione `xlsx.full.min.js` em `assets/vendor/`.
- Exportação ZIP não funciona:
  - Verifique conexão com internet ou adicione `jszip.min.js` em `assets/vendor/`.
  - Sem JSZip, o sistema faz fallback para download individual de PNGs.
- Favicon não aparece:
  - Faça recarga forçada do navegador para limpar cache (`Ctrl+F5`).

## Changelog

### 1.1 - 04/2026

#### Added

- Drag and drop dos eventos na tabela de cadastrados.
- Painel Organizar por acima das colunas, com seleção de campo e botão de organização ascendente.

#### Changed

- Fluxo de reordenação: ao reorganizar eventos (drag and drop ou organizador), a tabela e a prévia são atualizadas imediatamente.
- Persistência: a ordem dos eventos reorganizados passa a ser salva automaticamente no localStorage.
- Ordenação por Dia/Mês em ordem crescente quando o formato é válido (dd/mm).
- Ordenação por Horário em ordem crescente quando o formato é válido (hh:mm).
- Ordenação por Evento, Produtor e Onde comprar em ordem alfabética ascendente.

#### Fixed

- Consistência da experiência após organização: estado visual e estado salvo ficam sincronizados com a nova ordem.

### 1.0 - 04/2026

#### Added

- Lançamento inicial com as opções básicas de criação, importação e exportação da agenda.

## Créditos

Desenvolvido por Francis Castela para uso interno do Teatro Municipal de Itajaí.
Versão 1.1 - 04/2026.
