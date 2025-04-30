import express, { Request, Response } from "express";
import axios from "axios";
import { AuthorizationCode, ModuleOptions } from "simple-oauth2";
import dotenv from "dotenv";
import expressSession from "express-session";
declare module "express-session" {
  interface SessionData {
    token?: string;
  }
}

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
app.use(
  expressSession({
    secret: "teams-integration-secret",
    resave: false,
    saveUninitialized: true,
  })
);
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
    if (typeof token !== "string") {
      throw new Error("Access token inválido");
    }
    // salva o token na sessão
    req.session.token = token as string;

    // pega os teams que o usuário participa
    const teamsResponse = await axios.get("https://graph.microsoft.com/v1.0/me/joinedTeams", {
      headers: {
        Authorization: `Bearer ${token}`,
      },
    });

    const teams = teamsResponse.data.value;

    const html = teams
      .map(
        (team: any) => `
        <li>
          ${team.displayName} - ID: ${team.id} 
          <a href="/select-team/${team.id}">Selecionar</a>
        </li>
      `
      )
      .join("");

    res.send(`
      <h3>Times disponíveis:</h3>
      <ul>${html}</ul>
    `);
  } catch (error: any) {
    console.log(error)
    console.error("Erro ao autenticar:", error.response?.data || error.message);
    res.status(500).send("Erro ao autenticar");
  }
});


//  Rota de envio de mensagem (teste manual)
app.get("/send-message", async (_req: Request, res: Response) => {
  const token = process.env.TOKEN;
  const teamId = process.env.TEAM_ID;
  const channelId = process.env.CHANNEL_ID;

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

app.get("/select-team/:teamId", async (req: Request, res: Response) => {
  const { teamId } = req.params;
  const token = req.session.token;

  if (!token) return res.redirect("/");

  try {
    // pega os canais desse team
    const channelsResponse = await axios.get(
      `https://graph.microsoft.com/v1.0/teams/${teamId}/channels`,
      {
        headers: {
          Authorization: `Bearer ${token}`,
        },
      }
    );

    const channels = channelsResponse.data.value;

    const html = channels
      .map(
        (channel: any) => `
        <li>
          ${channel.displayName} - ID: ${channel.id}
          <a href="/send-message/${teamId}/${channel.id}">Enviar mensagem</a>
        </li>
      `
      )
      .join("");

    res.send(`
      <h3>Canais disponíveis em ${teamId}</h3>
      <ul>${html}</ul>
    `);
  } catch (error: any) {
    console.error("Erro ao buscar canais:", error.response?.data || error.message);
    res.status(500).send("Erro ao buscar canais");
  }
});

app.get("/send-message/:teamId/:channelId", async (req: Request, res: Response) => {
  const token = req.session.token;
  const { teamId, channelId } = req.params;

  if (!token) return res.redirect("/");

  try {
    await axios.post(
      `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${channelId}/messages`,
      {
        body: {
          content: "Mensagem enviada via Node.js + OAuth2 🚀",
        },
      },
      {
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
      }
    );

    res.send("✅ Mensagem enviada com sucesso!");
  } catch (error: any) {
    console.error("Erro ao enviar mensagem:", error.response?.data || error.message);
    res.status(500).send("Erro ao enviar mensagem.");
  }
});


//  Inicializa o servidor
app.listen(port, () => {
  console.log(`✅ App rodando em http://localhost:${port}`);
});
