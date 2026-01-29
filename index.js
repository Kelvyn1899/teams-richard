import express from "express";
import fetch from "node-fetch";

const app = express();
app.use(express.json());

const {
  TENANT_ID,
  CLIENT_ID,
  CLIENT_SECRET
} = process.env;

const ORGANIZER =
  "sistemas3.criciuma@borgesesilvaservicosmedicos.onmicrosoft.com";

/**
 * 🔑 Obter token do Microsoft Graph
 */
async function getAccessToken() {
  const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;

  const params = new URLSearchParams();
  params.append("client_id", CLIENT_ID);
  params.append("client_secret", CLIENT_SECRET);
  params.append("grant_type", "client_credentials");
  params.append("scope", "https://graph.microsoft.com/.default");

  const res = await fetch(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: params
  });

  const data = await res.json();

  if (!res.ok) {
    console.error("❌ ERRO TOKEN:", data);
    throw new Error("Falha ao obter access token");
  }

  console.log("✅ Token obtido com sucesso");
  return data.access_token;
}

/**
 * 📅 Criar reunião no Teams + calendário
 */
app.post("/teams/reuniao", async (req, res) => {
  console.log("📥 Webhook recebido do Bitrix:", req.body);

  const {
    participantName,
    participantEmail,
    startDateTime, // ISO sem timezone (ex: 2026-02-01T14:00:00)
    endDateTime,
    subject,
    description
  } = req.body;

  try {
    const token = await getAccessToken();

    const url = `https://graph.microsoft.com/v1.0/users/${ORGANIZER}/events`;

    const payload = {
      subject,
      body: {
        contentType: "HTML",
        content: description
      },
      start: {
        dateTime: startDateTime,
        timeZone: "America/Sao_Paulo"
      },
      end: {
        dateTime: endDateTime,
        timeZone: "America/Sao_Paulo"
      },
      attendees: [
        {
          emailAddress: {
            address: participantEmail,
            name: participantName
          },
          type: "required"
        }
      ],
      isOnlineMeeting: true,
      onlineMeetingProvider: "teamsForBusiness"
    };

    console.log("📤 Payload enviado ao Graph:", payload);

    const response = await fetch(url, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify(payload)
    });

    const data = await response.json();

    if (!response.ok) {
      console.error("❌ ERRO GRAPH:", {
        status: response.status,
        error: data
      });

      return res.status(response.status).json({
        error: "Erro ao criar reunião",
        graphStatus: response.status,
        graphError: data
      });
    }

    console.log("✅ Reunião criada com sucesso:", {
      eventId: data.id,
      joinUrl: data.onlineMeeting?.joinUrl
    });

    return res.json({
      eventId: data.id,
      joinUrl: data.onlineMeeting?.joinUrl
    });

  } catch (err) {
    console.error("❌ ERRO GERAL:", err);
    return res.status(500).json({ error: err.message });
  }
});

app.listen(3000, () => {
  console.log("🚀 Proxy Teams ativo na porta 3000");
});
