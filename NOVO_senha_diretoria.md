# Login da versão diretoria — E-MAIL + senha padrão

A versão diretoria abre com LOGIN por E-MAIL + SENHA PADRÃO. O código verifica se o
e-mail está na lista de liberados; se estiver, a senha padrão abre o dashboard completo.

## Como configurar (no topo do sales_dashboard_v10.py)
    EMAILS_LIBERADOS = ["lucas123@gmail", "diretor@alcon.com"]   # e-mails autorizados
    SENHA_PADRAO     = "Dados@"                                  # uma senha para todos

(Já vem pré-preenchido com o e-mail/senha de teste — TROQUE pelos reais.)

Ou por linha de comando (não deixa nada escrito no arquivo):
    python sales_dashboard_v10.py --emails "a@x.com,b@y.com" --senha-padrao "SuaSenha"

## Como funciona (envelope por e-mail)
- O HTML da diretoria é cifrado UMA vez com uma chave de dados aleatória (AES-256).
- Para CADA e-mail liberado, o gerador cria um "envelope" que abre essa chave usando
  a combinação (e-mail + senha padrão), via PBKDF2-SHA256.
- No login, digita-se e-mail + senha. Só abre se existir um envelope para aquele
  e-mail com aquela senha. E-mail fora da lista NÃO abre — nem com a senha certa —
  e não dá pra burlar pelo F12 (não existe envelope para ele).
- O e-mail é normalizado (minúsculas, sem espaços), então maiúsculas/espaços não atrapalham.
- Mensagem única "E-mail não liberado ou senha incorreta" (não revela quais e-mails valem).

## Testado (com lucas123@gmail / Dados@)
- lucas123@gmail + Dados@  → abre ✅
- LUCAS123@GMAIL (com espaço) + Dados@ → abre ✅ (normaliza)
- lucas123@gmail + senha errada → barra 🔒
- intruso@gmail + Dados@ (senha certa, e-mail não liberado) → barra 🔒

## Funcionalidades 100% mantidas
O HTML decifrado é byte a byte idêntico ao diretoria sem senha — mesmo dashboard,
mesmos gráficos e filtros. A criptografia só embrulha o arquivo fechado.

## Modo alternativo (só senha, sem e-mail)
Deixe EMAILS_LIBERADOS = [] e use SENHA_ADMIN / SENHA_DIRETOR (duas senhas que abrem).

## Sem CDN, offline
cryptojs_bundle.js vai embutido; abre no file:// sem internet. A geração usa a
'cryptography' do Python (Anaconda). Se faltar a lib ou o bundle, gera SEM senha (avisa).

## Segurança — honesto
Protege o arquivo FECHADO. Depois que a pessoa certa abre, o F12 dela vê o que ela já
pode ver. A lista de e-mails restringe QUEM abre; senha forte reforça contra quebra.
