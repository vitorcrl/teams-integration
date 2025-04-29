import express, { Request, Response } from "express";
import axios from "axios";
import { AuthorizationCode, ModuleOptions } from "simple-oauth2";
import dotenv from "dotenv";

dotenv.config();

const app = express();
const port = 3000;

//  Configuração do OAuth2
const oauthConfig: ModuleOptions<"client_id"> = {
  client: {
    id: process.env.CLIENT_ID as string,
    secret: process.env.CLIENT_SECRET as string,
  },
  auth: {
    tokenHost: "https://login.microsoftonline.com",
    authorizePath: `/${process.env.TENANT_ID}/oauth2/v2.0/authorize`,
    tokenPath: `/${process.env.TENANT_ID}/oauth2/v2.0/token`,
  },
};

const oauthClient = new AuthorizationCode(oauthConfig);

//  Rota inicial
app.get("/", (_req: Request, res: Response) => {
  const authorizationUri = oauthClient.authorizeURL({
    redirect_uri: process.env.REDIRECT_URI,
    scope: "User.Read Chat.ReadWrite ChannelMessage.Send",
  });

  res.send(`
    <h2>Conectar ao Microsoft Teams</h2>
    <a href="${authorizationUri}">Clique aqui para conectar</a>
  `);
});

//  Callback do login
app.get("/auth/callback", async (req: Request, res: Response) => {
  const code = req.query.code as string;

  try {
    const tokenParams = {
      code,
      redirect_uri: process.env.REDIRECT_URI,
      scope: "User.Read Chat.ReadWrite ChannelMessage.Send",
    };

    const accessToken = await oauthClient.getToken(tokenParams as any);
    const token = accessToken.token.access_token;

    console.log("Access Token:", token);

    res.send(`
      <h3>Autenticado com sucesso!</h3>
      <p>Access token obtido! (confira no terminal)</p>
      <a href="/send-message">Enviar mensagem ao Teams</a>
    `);
  } catch (error: any) {
    console.error("Erro ao obter token:", error.response?.data || error.message);
    res.status(500).send("Erro ao autenticar");
  }
});

//  Rota de envio de mensagem (teste manual)
app.get("/send-message", async (_req: Request, res: Response) => {
  const token = "COLE_SEU_TOKEN_AQUI";
  const teamId = "SEU_TEAM_ID";
  const channelId = "SEU_CHANNEL_ID";

  try {
    await axios.post(
      `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${channelId}/messages`,
      {
        body: {
          content: "Mensagem enviada via Node.js + TypeScript 🎯",
        },
      },
      {
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
      }
    );

    res.send("Mensagem enviada com sucesso!");
  } catch (error: any) {
    console.error("Erro ao enviar mensagem:", error.response?.data || error.message);
    res.status(500).send("Erro ao enviar mensagem.");
  }
});

//  Inicializa o servidor
app.listen(port, () => {
  console.log(`✅ App rodando em http://localhost:${port}`);
});
