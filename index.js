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

    // Suporte a múltiplos nomes
    const nomes = nome
      .split(',')
      .map(n => n.trim())
      .filter(n => n.length > 0);

    // Suporte a múltiplos e-mails
    const emails = email
      .split(',')
      .map(e => e.trim())
      .filter(e => e.length > 0);

    if (emails.length === 0) {
      return res.status(400).json({ error: 'Nenhum e-mail válido informado' });
    }

    // Associa nome + e-mail pelo índice
    const attendees = emails.map((address, index) => ({
      emailAddress: {
        address,
        name: nomes[index] || nomes[nomes.length - 1] || 'Convidado'
      },
      type: 'required'
    }));

    console.log('👥 Convidados processados:', attendees);

    const token = await getAccessToken();

    const start = bitrixToISO(data);
    const end = bitrixToISO(data);

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
      attendees,
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
      joinUrl,
      totalAttendees: attendees.length
    });

    // Preenche o campo do Bitrix com o link da reunião
    await axios.post(BITRIX_WEBHOOK_URL, {
      id: cardId,
      entityTypeId,
      fields: {
        ufCrm67_1705518728148: joinUrl
      }
    });

    console.log('✅ Link da reunião salvo no Bitrix');

    res.json({
      success: true,
      eventId: response.data.id,
      joinUrl,
      convidados: attendees
    });

  } catch (error) {
    console.error('❌ Erro Graph:', error.response?.data || error.message);
    res.status(500).json({
      error: 'Erro ao criar reunião',
      details: error.response?.data || error.message
    });
  }
});

app.listen(PORT, () => {
  console.log(`🚀 Proxy Teams rodando na porta ${PORT}`);
});
