# Etiquetas Sid

Programa local em Python para preparar etiquetas de pacientes em folhas A4, com 2 colunas e 4 linhas de 99,1 × 67,7 mm.

## Executar

Instale Python com suporte a Tkinter e execute na pasta do projeto:

```powershell
python -m pip install -r requirements.txt
python etiquetas_sid.py
```

Também é possível usar `iniciar.bat` após instalar as dependências. Para criar o executável, execute `gerar_exe.bat` e copie `dist/EtiquetasSid.exe` para a pasta principal.

## Arquivos

- `BD/BD.xlsx`: primeira aba, cabeçalhos na primeira linha entre A e P. Colunas obrigatórias: `NÚMERO VIDAS`, `NOME COMPLETO`, `DATA NASCIMENTO`.
- `responsaveis.txt`: UTF-8, um funcionário por linha. Edite no Bloco de Notas e clique em Recarregar responsáveis.
- `config.json`: margens e espaços em milímetros, editáveis também pela interface.

Números informados com zeros à esquerda encontram o número correspondente sem zeros na planilha. A etiqueta mostra o número da planilha. Registros equivalentes duplicados impedem a leitura para evitar seleção ambígua. Datas devem ser datas do Excel ou texto em dd/mm/aaaa.

## Usar

1. Digite números separados por espaços, vírgulas ou ponto e vírgula e clique em Adicionar pacientes.
2. Selecione uma ou várias etiquetas com Ctrl/Shift, ou clique em Selecionar todas. Preencha PERÍODO REF. e escolha RESP. ENVIO. Clique em Aplicar às selecionadas antes de trocar a seleção. Na edição em grupo, campos vazios preservam os valores existentes; campos com valores iguais aparecem preenchidos.
3. Para várias etiquetas do mesmo paciente, adicione seu número novamente.
4. Escolha a posição inicial de 1 a 8 (da esquerda para a direita, de cima para baixo).
5. Gere o PDF e imprima pelo leitor de PDF em A4, tamanho real/100%, sem ajustar à página.

DATA RECOLHIMENTO e RESP. RECOLHIMENTO ficam em branco para preenchimento à caneta. As etiquetas excedentes continuam na página seguinte.

As margens iniciais são estimadas, pois a embalagem não informa o gabarito. Faça primeiro um teste em papel comum e compare com a folha adesiva contra a luz. Ajuste o alinhamento pela interface antes de usar etiquetas. A impressão física e as margens da impressora precisam ser validadas no equipamento.

O programa não modifica a planilha nem envia dados pela rede. PDFs contêm informações dos pacientes e ficam no local escolhido pelo usuário.

O número do paciente aparece sem casas decimais. O ícone da janela e do executável usa `imgs/icon.png`, incluído automaticamente na compilação.

A interface abre no modo escuro. Use o botão Modo escuro no topo para alternar para o tema claro. A etiqueta mantém o fundo branco, com rótulos e valores em negrito e divisórias mais escuras para melhorar a leitura.
