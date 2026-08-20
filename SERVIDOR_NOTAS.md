# Notas fiscais lidas do servidor remoto

O programa le as DANFEs direto da pasta do servidor, sem precisar copiar nada
para o PC local.

| Campo | Valor |
| --- | --- |
| Servidor | `10.121.19.8` |
| Pasta dentro do servidor | `C:\Users\marfimbh1\Desktop\NOTAS` |
| Caminho usado pelo programa | `\\10.121.19.8\C$\Users\marfimbh1\Desktop\NOTAS` |
| Usuario | `marfimbh1` |
| Senha | **nao fica no repositorio** - veja abaixo |

Isso ja esta gravado no `config.json`, no bloco `servidor`. Nada precisa ser
digitado, exceto a senha.

## A senha

A senha **nao** vai para o `config.json` (esse arquivo sobe para o GitHub).
Ela fica so no PC que roda o programa, em `credenciais_rede.json`, cifrada
com a conta do Windows (DPAPI, via `pywin32`).

Formas de informar a senha, na ordem em que o programa procura:

1. Variavel de ambiente `ROMANEIO_SENHA_REDE`.
2. Arquivo `credenciais_rede.json`, ao lado do programa.
3. Na primeira leitura o programa pergunta a senha numa janelinha e guarda
   sozinho no arquivo acima. E o caminho mais simples: abrir o programa,
   digitar a senha uma vez, pronto.

Para trocar ou apagar depois: **Configuracoes > Servidor das notas**, campo
`Senha` (botao `Esquecer` limpa).

## Se der "Acesso negado" (erro 5)

O caminho `C$` e o compartilhamento administrativo do Windows: ele so aceita
contas que sejam **administradoras do servidor**. Se `marfimbh1` nao for
administradora em `10.121.19.8`, use um compartilhamento normal:

1. No servidor: botao direito na pasta `NOTAS` > Propriedades >
   Compartilhamento > Compartilhar, liberando leitura para `marfimbh1`.
2. No programa: **Configuracoes > Servidor das notas**
   - `Compartilhamento`: `NOTAS` (o nome que aparece no servidor)
   - `Pasta no servidor`: deixar em branco
3. `Testar conexao`.

## Outros avisos comuns

| Mensagem | O que fazer |
| --- | --- |
| Servidor nao encontrado (53) | Servidor desligado, IP trocado ou fora da VPN/rede |
| Usuario ou senha incorretos (1326) | Corrigir a senha em Configuracoes > Servidor das notas |
| Ja existe conexao com outro usuario (1219) | Menu **Arquivo > Reconectar ao servidor** |
| Compartilhamento nao encontrado (67) | Nome do compartilhamento errado |

## Voltar a ler de uma pasta local

Pelo botao `...` ao lado da pasta, na tela principal: ao escolher uma pasta do
proprio computador o programa pergunta se e para parar de ler do servidor.

Ou por **Configuracoes > Servidor das notas**, desmarcando `Ler as notas de um
servidor da rede`.

## Menu Arquivo > Reconectar ao servidor

Derruba a conexao atual e entra de novo. Serve para o erro 1219 (o Windows ja
tinha uma conexao com esse servidor por outro usuario) e para quando o
servidor foi reiniciado com o programa aberto.
