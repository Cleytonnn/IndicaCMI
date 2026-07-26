# Configuração MongoDB Atlas

## 1. Criar conta (GRATUITO)
1. Acesse https://www.mongodb.com/cloud/atlas
2. Clique em "Try Free"
3. Crie sua conta com email

## 2. Criar Cluster
1. Após login, clique em "Create a Deployment"
2. Escolha "M0 Free" (gratuito)
3. Escolha provedor e região
4. Clique "Create Cluster" (aguarde 5-10 min)

## 3. Gerar String de Conexão
1. Clique em "Connect"
2. Escolha "Drivers"
3. Selecione "Node.js"
4. Copie a string de conexão
5. Substitua `<username>:<password>` por suas credenciais

## 4. Variáveis de Ambiente - Vercel
1. Acesse: https://vercel.com/indica2/indica-cmi/settings/environment-variables
2. Clique em "Add"
3. Name: `MONGODB_URI`
4. Value: Cole a string de conexão do MongoDB
5. Clique em "Save" e "Redeploy"

## 5. Pronto!
Sua aplicação agora usa MongoDB em produção! 🎉
