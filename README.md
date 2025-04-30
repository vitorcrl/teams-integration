# Teams Manual Integration API

Este projeto é uma API simples para integrar manualmente uma aplicação com o Microsoft Teams. Ele utiliza o OAuth2 para autenticação e permite o envio de mensagens a um canal do Teams por meio de uma rota HTTP.

## Funcionalidades

- Autenticação via OAuth2 com Microsoft Identity Platform
- Geração manual do access token
- Envio de mensagens para um canal específico do Microsoft Teams

## Tecnologias Utilizadas

- Node.js
- Express
- TypeScript
- Axios
- simple-oauth2
- dotenv

## Instalação

1. Clone o repositório:

```bash
git clone https://github.com/seu-usuario/teams-integration-api.git
cd teams-integration-api
```

2. Instale as dependências:

```bash
npm install
```

3. Crie um arquivo `.env` com as seguintes variáveis:

```env
CLIENT_ID=seu_client_id
CLIENT_SECRET=seu_client_secret
TENANT_ID=seu_tenant_id
REDIRECT_URI=http://localhost:3000/auth/callback
```

> Estas informações você consegue ao registrar sua aplicação no [portal do Azure](https://portal.azure.com/).

## Uso

1. Inicie o servidor:

```bash
npm run start:dev
```

2. Acesse [http://localhost:3000](http://localhost:3000) no navegador.

3. Clique em "Clique aqui para conectar" para iniciar o processo de autenticação com sua conta do Microsoft Teams.

4. Após o login, você será redirecionado para uma página que confirma o sucesso e fornece um link para enviar uma mensagem ao Teams.

5. Antes de testar o envio de mensagem:
   - Copie o token gerado (mostrado no terminal) e substitua manualmente na variável `token` na rota `/send-message`.
   - Substitua os valores de `teamId` e `channelId` pelos seus valores reais.

6. Acesse [http://localhost:3000/send-message](http://localhost:3000/send-message) para enviar uma mensagem de teste ao canal do Teams.

## Observações

- Este projeto é **manual e experimental**. O token expira e precisa ser renovado periodicamente.
- Ideal para testes ou provas de conceito. Para produção, implemente refresh token e armazenamento seguro.

## Licença

MIT
