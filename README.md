# BEVI para 2TECH

Conversor local para Windows para padronizar tabelas de comissionamento BEVI no layout 2TECH.

## Como usar no modo app

1. Abra `dist\BEVI_2TECH_C6.exe`.
2. O navegador vai abrir uma tela local do conversor.
3. Arraste ou selecione a tabela de comissão recebida da BEVI.
4. Preencha a vigência inicial se quiser sobrescrever a data da BEVI. Para arquivos do Banco PAN, NEO CRÉDITO e FINANTO, ela é obrigatória porque as tabelas não trazem data de vigência.
5. Clique em `Converter para 2TECH` e depois em `Baixar arquivo gerado`.

O executável abre uma interface local em `localhost`, então não precisa instalar Python nas máquinas dos usuários.

## Recompilar o executável

Execute `build_exe.bat` depois de qualquer alteração no código.

## Regras implementadas

- Bancos mapeados: `C6BANK`, `DIGIO`, `DAYCOVAL`, `SAFRA`, `PAN`, `NEO CREDITO`, `BANRISUL` e `FINANTO`.
- Taxas e percentuais: gravados com vírgula decimal para importação na Check.
- Convênio: reconhecido pelo nome da tabela, atualmente `INSS` ou `SIAPE`.
- Tipo de contrato: coluna `Forma Contrato` da BEVI, com regras específicas para os layouts PAN e NEO.
- Formalização: `FÍSICO` quando o nome contém `WEB PLUS`; caso contrário, `DIGITAL`.
- Prazo inicial e final: coluna `Prazo`. Aceita prazo único (`108`) ou faixa na mesma coluna (`1 a 108`).
- `À Vista (Empresa)`: coluna `Comissão Total`.
- `À Vista (Repasse 1)`: `À Vista (Empresa) * 45%`, salvo como valor, sem fórmula.
- Nome da tabela: descrição banco + nome sem o prazo final + taxa inicial/final.
- Nome da tabela: remove trecho de taxa interna como `TX 1,850~1,35` quando aparecer.
- Ticket no nome: `MAIOR 20K` vira valor inicial `20000`; `6K A 20K` vira `6000` a `20000`; `ATE 6K` vira `0` a `6000`.
- Refinanciamento: gera também a variante `REFIN + MARGEM`, exceto para Digio, Banco do Brasil, NEO CRÉDITO e FINANTO.
- PAN: lê arquivos `.xlsx`, usa `Flat Repassada` como comissão e exige vigência inicial.
- NEO CRÉDITO: monta o nome com convênio, tipo e código da tabela; exige vigência inicial e não duplica refinanciamento.

## Próximos bancos

Cada banco deve entrar como uma nova função de conversão, mantendo a interface igual.
