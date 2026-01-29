require('dotenv').config();
const express = require('express');
const axios = require('axios');
const cors = require('cors');

const app = express();
const PORT = 3000;

const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const BITRIX_WEBHOOK_URL = process.env.BITRIX_WEBHOOK_URL;

const ORGANIZER_EMAIL = 'sistemas3.criciuma@borgesesilvaservicosmedicos.onmicrosoft.com';

app.use(cors());
app.use(express.json());

// Converte data Bitrix -> ISO Graph
function bitrixToISO(dateStr) {
  const [datePart, timePart] = dateStr.split(' ');
  const [day, month, year] = datePart.split('/');
  return `${year}-${month}-${day}T${timePart}`;
}

// Token OAuth Graph
async function getAccessToken() {
  const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;

  const params = new URLSearchParams();
  params.append('client_id', CLIENT_ID);
  params.append('client_secret', CLIENT_SECRET);
  params.append('grant_type', 'client_credentials');
  params.append('scope', 'https://graph.microsoft.com/.default');

  const response = await axios.post(url, params);
  return response.data.access_token;
}

app.post('/create-meeting', async (req, res) => {
  const {
    nome,
    email,
    data,
    titulo,
    descricao,
    cardId,
    entityTypeId
  } = req.body;

  if (!nome || !email || !data || !titulo || !cardId || !entityTypeId) {
    return res.status(400).json({ error: 'Campos obrigatórios ausentes' });
  }

  try {
    console.log('📥 Payload recebido:', req.body);

    const token = await getAccessToken();

    const start = bitrixToISO(data);
    const end = bitrixToISO(data); // mantém como estava

    const eventPayload = {
      subject: titulo,
      body: {
        contentType: 'HTML',
        content: descricao || ''
      },
      start: {
        dateTime: start,
        timeZone: 'America/Sao_Paulo'
      },
      end: {
        dateTime: end,
        timeZone: 'America/Sao_Paulo'
      },
      attendees: [
        {
          emailAddress: {
            address: email,
            name: nome
          },
          type: 'required'
        }
      ],
      isOnlineMeeting: true,
      onlineMeetingProvider: 'teamsForBusiness'
    };

    const graphUrl = `https://graph.microsoft.com/v1.0/users/${ORGANIZER_EMAIL}/events`;

    const response = await axios.post(graphUrl, eventPayload, {
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json'
      }
    });

    const joinUrl = response.data.onlineMeeting?.joinUrl || '';

    console.log('✅ Reunião criada:', {
      id: response.data.id,
      joinUrl
    });

    // ============================
    // 🔹 NOVO TRECHO – ATUALIZA BITRIX
    // ============================
    const bitrixPayload = {
      id: cardId,
      entityTypeId,
      fields: {
        ufCrm67_1705518728148: joinUrl
      }
    };

    const bitrixResponse = await axios.post(
      BITRIX_WEBHOOK_URL,
      bitrixPayload
    );

    console.log('✅ Link da reunião salvo no Bitrix:', {
      cardId,
      field: 'ufCrm67_1705518728148',
      joinUrl
    });

    // ============================

    res.json({
      success: true,
      eventId: response.data.id,
      joinUrl,
      bitrix: bitrixResponse.data
    });

  } catch (error) {
    console.error('❌ Erro no fluxo:', error.response?.data || error.message);
    res.status(500).json({
      error: 'Erro ao criar reunião ou atualizar Bitrix',
      details: error.response?.data || error.message
    });
  }
});

app.listen(PORT, () => {
  console.log(`🚀 Proxy Teams rodando na porta ${PORT}`);
});
