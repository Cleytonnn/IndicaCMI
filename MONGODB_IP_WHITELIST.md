# ⚙️ Desbloqueando MongoDB Atlas para Vercel

## Passo 1: Acesse MongoDB Atlas
👉 Abra: https://cloud.mongodb.com/v2

**Clique em: "Log in with Email"**
- Email: `cleytonbaumgratz_db_user` ou o email da sua conta
- Senha: sua senha

## Passo 2: Acesse seu Cluster
1. Após fazer login, clique no seu projeto
2. Clique em **"Cluster0"**

## Passo 3: Vá para Network Access
1. No menu ESQUERDO, procure por **"Network Access"** (ou "Security" > "Network Access")
2. Clique em **"ADD IP ADDRESS"**

## Passo 4: Adicione 0.0.0.0/0
1. Na janela que abriu, clique em **"ALLOW ACCESS FROM ANYWHERE"**
   - Isso coloca `0.0.0.0/0` automaticamente
2. Clique em **"Confirm"**

## Passo 5: Aguarde (1-2 minutos)
- Status mudará para **"ACTIVE"** em verde

## ✅ Pronto!
Sua aplicação agora conectará com MongoDB!

---

### Se tiver dúvida, captura de tela dos steps:
1. Login → Email e Senha
2. Dashboard → Seu cluster
3. Menu esquerdo → "Network Access"
4. Botão → "ADD IP ADDRESS"
5. Checkbox → "ALLOW ACCESS FROM ANYWHERE"
6. Botão → "Confirm"

**Aviso:** Deixar `0.0.0.0/0` é prático para desenvolvimento. Em produção, adicione apenas IPs específicos!
